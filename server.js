const express = require('express');
const multer = require('multer');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const Anthropic = require('@anthropic-ai/sdk');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');
const bcrypt = require('bcryptjs');
const crypto = require('crypto');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 150 * 1024 * 1024 } });

// Security headers (HSTS, X-Frame-Options, X-Content-Type-Options, etc.)
app.use(helmet({ contentSecurityPolicy: false })); // CSP off since inline scripts are used across the existing dashboards

// Restrict cross-origin API access to known app domains only
const ALLOWED_ORIGINS = [
  'https://azr-operations.com',
  'https://azhar-ai-la1l.onrender.com',
  'http://localhost:3000'
];
app.use(cors({
  origin: function (origin, callback) {
    if (!origin || ALLOWED_ORIGINS.includes(origin)) return callback(null, true);
    return callback(new Error('Not allowed by CORS'));
  },
  credentials: true
}));

// Brute-force protection on login: 10 attempts per 15 minutes per IP
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 10,
  message: { error: 'Too many login attempts. Please try again in 15 minutes.' },
  standardHeaders: true,
  legacyHeaders: false
});

app.use(express.json({ limit: '150mb' }));
app.use(express.urlencoded({ extended: true, limit: '150mb' }));

const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

//  PostgreSQL 
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: process.env.DATABASE_URL ? { rejectUnauthorized: false } : false
});

//  Fallback file storage (if DB not available) 
const DATA_DIR = path.join(__dirname, '.data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
const DISPATCH_FILE  = path.join(DATA_DIR, 'dispatch.json');
const REJECTION_FILE = path.join(DATA_DIR, 'rejection.json');
// ── Learned lookups for rejection auto-classification, built once from 7
// months of already-staff-classified history (Jan-Jul 2026). Lets a RAW
// transport file (only has transport's own freeform reason text — no
// Final Root Cause / ORG-BU / Internal-External columns at all) get those
// 3 columns auto-filled instead of requiring manual re-classification.
// Regenerate by re-running the same learning script against a newer master
// export if the categories drift significantly over time. ──
var REJECTION_LOOKUPS = { reasonToRootCause: {}, orgBuFine: {}, orgBuLoose: {}, rootCauseToInternalExternal: {} };
try {
  REJECTION_LOOKUPS = JSON.parse(fs.readFileSync(path.join(__dirname, 'rejection_lookups.json'), 'utf8'));
  console.log('Loaded rejection classification lookups:', Object.keys(REJECTION_LOOKUPS.reasonToRootCause).length, 'reason patterns,',
    Object.keys(REJECTION_LOOKUPS.orgBuFine).length + Object.keys(REJECTION_LOOKUPS.orgBuLoose).length, 'org/type combos,',
    Object.keys(REJECTION_LOOKUPS.rootCauseToInternalExternal).length, 'root-cause categories');
} catch (e) { console.error('Could not load rejection_lookups.json — auto-classification will fall back to rules only:', e.message); }

// Generalizes a raw reason note into a matchable template — same normalization
// used when the lookup was built, so dates/numbers don't prevent a match
// (e.g. "CUSTOMER NEED DELIVERY ON 19-JAN-2026" and "...23-FEB-2026" both hit
// the same learned entry).
function normalizeReasonText(raw) {
  var t = String(raw || '').toUpperCase().trim().replace(/\s+/g, ' ');
  t = t.replace(/\b\d{1,2}[-\/][A-Z]{3}[-\/]\d{2,4}\b/g, '<DATE>');
  t = t.replace(/\b\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}\b/g, '<DATE>');
  t = t.replace(/\b\d{4,}\b/g, '<NUM>');
  return t.trim();
}

// Every clean category name this system can ever produce — used to make
// autoClassifyRootCause() idempotent. This matters because rebuilding the
// dashboard total re-runs this same function over already-classified
// historical rows pulled back out of the DB (see rejection_rows below); a
// clean name fed back in must return unchanged, not get re-titlecased into a
// near-duplicate (e.g. "PO Box" -> "Po Box").
var KNOWN_CLEAN_CATEGORIES = new Set(['INTERNAL TRANSFER']);
Object.values(REJECTION_LOOKUPS.reasonToRootCause).forEach(function(c) { KNOWN_CLEAN_CATEGORIES.add(String(c).toUpperCase()); });
[
  'Merchandiser Unavailable on Route', 'Pending Goods Return Voucher Not Ready (GRV)', 'Customer System Down',
  'Duplicate Item Received Under Separate LPO', 'Returned — Weather/Road Closure', 'Declined — No Stock Requirement',
  'Declined — LPO Cancelled in Customer System', 'Returned — Receiving Closed for the Day', 'Declined — Insufficient Storage Space',
  'Declined — No Delivery Schedule Confirmed', 'Returned — Customer Payment Not Ready (Cheque/Cash)',
  'Declined — LPO Not Reflected in Customer System', 'Declined — Customer System Under Maintenance',
  'Declined — Insufficient Storage Capacity Today', 'Declined — Outside Receiving Hours'
].forEach(function(c) { KNOWN_CLEAN_CATEGORIES.add(c.toUpperCase()); });

// Translates raw, informally-typed root-cause text from transport staff into
// professional wording for anything leadership sees. Matches on distinctive
// substrings so minor wording/typo variants ("REFUSED DUE TO" vs "REFUSED TO
// ACCEPT DUE TO") merge into the same clean category instead of showing as
// separate duplicate rows.
function polishRootCause(raw) {
  var s = toStr(raw);
  if (!s) return s;
  var u = s.toUpperCase().trim();
  if (u.indexOf('MERCHANDISER') !== -1 && u.indexOf('ROUTE') !== -1) return 'Merchandiser Unavailable on Route';
  if (u.indexOf('FIRST GRV') !== -1 || u.indexOf('COLLECT GRV') !== -1) return 'Pending Goods Return Voucher Not Ready (GRV)';
  if (u.indexOf('SYSTEM NOT WORKING') !== -1) return 'Customer System Down';
  if (u.indexOf('SAME ITEM') !== -1 && u.indexOf('LPO') !== -1) return 'Duplicate Item Received Under Separate LPO';
  if (u.indexOf('HEAVY RAIN') !== -1 || (u.indexOf('ROAD CLOSURE') !== -1 && u.indexOf('RETURN') !== -1)) return 'Returned — Weather/Road Closure';
  if (u.indexOf('NO NEED STOCK') !== -1) return 'Declined — No Stock Requirement';
  if (u.indexOf('LPO DELETED') !== -1) return 'Declined — LPO Cancelled in Customer System';
  if (u.indexOf('RECEIVING CLOSED') !== -1) return 'Returned — Receiving Closed for the Day';
  if (u.indexOf('NO SPACE') !== -1) return 'Declined — Insufficient Storage Space';
  if (u.indexOf('NO SCHEDULE') !== -1) return 'Declined — No Delivery Schedule Confirmed';
  if (u.indexOf('PAYMENT') !== -1 && (u.indexOf('NOT READY') !== -1 || u.indexOf('NOT READ') !== -1)) return 'Returned — Customer Payment Not Ready (Cheque/Cash)';
  if (u.indexOf('LPO') !== -1 && u.indexOf('NOT') !== -1 && /L\w{0,3}ECTED/.test(u)) return 'Declined — LPO Not Reflected in Customer System';
  if (u.indexOf('SYSTEM UPDATING') !== -1) return 'Declined — Customer System Under Maintenance';
  if (u.indexOf('UNABLE TO ACCOMMODATE') !== -1) return 'Declined — Insufficient Storage Capacity Today';
  if (u.indexOf('RECEIVING TIME OVER') !== -1) return 'Declined — Outside Receiving Hours';
  // Fallback for anything not yet mapped: tidy spacing + Title Case, so it's
  // never worse than today even before someone adds a proper mapping for it.
  return s.replace(/\s+/g, ' ').trim().toLowerCase().replace(/\b\w/g, function(c) { return c.toUpperCase(); });
}

// Root cause: try the learned lookup (from real staff-classified history)
// first — it covers far more phrasings than the 14 hardcoded rules below.
// Falls through to the hardcoded rules, then a tidy-up fallback, exactly as
// before, so this is a strict improvement, never a regression.
function autoClassifyRootCause(raw) {
  var s = toStr(raw);
  if (!s) return { category: s, matched: false };
  if (KNOWN_CLEAN_CATEGORIES.has(s.toUpperCase())) return { category: s, matched: true }; // already a final category — pass through unchanged
  var key = normalizeReasonText(s);
  var learned = REJECTION_LOOKUPS.reasonToRootCause[key];
  if (learned) return { category: learned, matched: true };
  var polished = polishRootCause(s);
  // polishRootCause() either hit a real rule (a confident match, same as the
  // learned lookup) or fell through to its raw tidy-up-and-titlecase fallback
  // (genuinely unclassified — belongs in needsReview). Detect which by
  // comparing against what that fallback alone would have produced.
  var titleCaseOnly = s.replace(/\s+/g, ' ').trim().toLowerCase().replace(/\b\w/g, function(c) { return c.toUpperCase(); });
  var hitRealRule = polished !== titleCaseOnly;
  return { category: polished, matched: hitRealRule };
}

// Food/Non-Food (+ special sub-brand labels like "Shark Ninja-DSN", "Group Seb")
// derived from Organization + Order Type (+ Customer Group as a tiebreaker),
// learned from history. Falls back to the row's own raw Type text when no
// confident match exists, so an unusual combo never gets guessed wrong.
function autoClassifyOrgBu(org, orderType, custGroup, rawTypeStr) {
  var fineKey = [org, orderType, custGroup].join('|');
  var looseKey = [org, orderType].join('|');
  var label = REJECTION_LOOKUPS.orgBuFine[fineKey] || REJECTION_LOOKUPS.orgBuLoose[looseKey] || null;
  if (!label) return { label: rawTypeStr || '', isFood: null, isNonFood: null, matched: false };
  var upper = label.toUpperCase();
  var isFood = upper === 'FOOD';
  // Special sub-brand labels (Shark Ninja-DSN / Shark Ninja-E-Com / Group Seb) count
  // as Non-Food for the dashboard's Food/Non-Food split, same as the team's own rule.
  var isNonFood = !isFood;
  return { label: label, isFood: isFood, isNonFood: isNonFood, matched: true };
}

// Internal/External derived from the (now-known) Final Root Cause — "Customer
// Related" is the only category that's External; every other category
// (Sales/GRV/Merchandising, Warehouse, Transport, CST, IT, Force Majeure,
// Partial Pick) is Internal. This mirrors a 100%-clean split found in 7
// months of history, so it's a hard rule, not a guess.
//
// The learned lookup only covers category text that actually appeared in
// that history. The 14 older hardcoded polishRootCause() categories use
// different wording, so they get their own explicit mapping here — based on
// the same External="customer's fault" / Internal="AKI's/weather/etc" logic
// the historical data follows, not independently re-verified per phrase.
var HARDCODED_CATEGORY_TO_IE = {
  'MERCHANDISER UNAVAILABLE ON ROUTE': 'Internal',
  'PENDING GOODS RETURN VOUCHER NOT READY (GRV)': 'Internal',
  'CUSTOMER SYSTEM DOWN': 'External',
  'DUPLICATE ITEM RECEIVED UNDER SEPARATE LPO': 'External',
  'RETURNED — WEATHER/ROAD CLOSURE': 'Internal',
  'DECLINED — NO STOCK REQUIREMENT': 'External',
  'DECLINED — LPO CANCELLED IN CUSTOMER SYSTEM': 'External',
  'RETURNED — RECEIVING CLOSED FOR THE DAY': 'External',
  'DECLINED — INSUFFICIENT STORAGE SPACE': 'External',
  'DECLINED — NO DELIVERY SCHEDULE CONFIRMED': 'External',
  'RETURNED — CUSTOMER PAYMENT NOT READY (CHEQUE/CASH)': 'External',
  'DECLINED — LPO NOT REFLECTED IN CUSTOMER SYSTEM': 'External',
  'DECLINED — CUSTOMER SYSTEM UNDER MAINTENANCE': 'External',
  'DECLINED — INSUFFICIENT STORAGE CAPACITY TODAY': 'External',
  'DECLINED — OUTSIDE RECEIVING HOURS': 'External'
};
function deriveInternalExternal(rootCauseCategory) {
  var upperCat = String(rootCauseCategory || '').toUpperCase();
  var detailed = REJECTION_LOOKUPS.rootCauseToInternalExternal[upperCat];
  if (detailed) return { detailed: detailed, binary: (detailed === 'Customer Related') ? 'External' : 'Internal', matched: true };
  var fallbackBinary = HARDCODED_CATEGORY_TO_IE[upperCat];
  if (fallbackBinary) return { detailed: null, binary: fallbackBinary, matched: true };
  return { detailed: null, binary: null, matched: false };
}

// ── Shared aggregation core: takes ONE classified rejection/delivery record
// and accumulates it into the running org/month totals. Used by BOTH the
// upload endpoint (looping over freshly-classified new rows) and the row
// correction endpoint (looping over ALL stored rows after one is edited) —
// having one shared function means both paths always compute totals exactly
// the same way, with zero risk of drift between them. ──
function accumulateRejectionRow(orgMap, monthMap, rec) {
  var org = rec.org, root = rec.root, cust = rec.cust, addr = rec.addr, area = rec.area,
      mo = rec.mo, day = rec.day, isFood = rec.isFood, isNF = rec.isNF, srcStr = rec.srcStr,
      val = rec.val, rej = rec.rej, del = rec.del;
      if (org) {
        if (!orgMap[org]) orgMap[org]={tDel:0,tRej:0,val:0,food_rej:0,food_del:0,nonfood_rej:0,nonfood_del:0,ext_rej:0,ext_del:0,int_rej:0,int_del:0,food_val:0,nonfood_val:0,del:new Array(12).fill(0),rej:new Array(12).fill(0),reasons:{},custs:{},areas:{},food_reasons:{},food_custs:{},nonfood_reasons:{},nonfood_custs:{},ext_reasons:{},ext_custs:{},int_reasons:{},int_custs:{},detail:{},food_detail:{},nonfood_detail:{},ext_detail:{},int_detail:{}};
        // Full per-org, per-month breakdown (food/nonfood/external/internal) — without this,
        // ORG + MONTH + SOURCE filters combined would fall back to all-org month totals,
        // which is what caused the >100% "Contribution to Rejection Rate" bug.
        if (mo) {
          if(!orgMap[org].byMonth) orgMap[org].byMonth={};
          if(!orgMap[org].byMonth[mo]) orgMap[org].byMonth[mo]={
            tRej:0,tDel:0,val:0,food_rej:0,food_del:0,nonfood_rej:0,nonfood_del:0,
            ext_rej:0,ext_del:0,int_rej:0,int_del:0,food_val:0,nonfood_val:0,
            reasons:{},custs:{},detail:{},
            food_reasons:{},food_custs:{},food_detail:{},
            nonfood_reasons:{},nonfood_custs:{},nonfood_detail:{},
            ext_reasons:{},ext_custs:{},ext_detail:{},
            int_reasons:{},int_custs:{},int_detail:{}
          };
        }
        var mb = mo ? orgMap[org].byMonth[mo] : null;
        if (del) {
          orgMap[org].tDel++; if(mo)orgMap[org].del[mo-1]++;
          if(isFood)orgMap[org].food_del++; else if(isNF)orgMap[org].nonfood_del++;
          if(srcStr==='EXTERNAL')orgMap[org].ext_del++; else if(srcStr==='INTERNAL')orgMap[org].int_del++;
          if(mb){
            mb.tDel++;
            if(isFood)mb.food_del++; else if(isNF)mb.nonfood_del++;
            if(srcStr==='EXTERNAL')mb.ext_del++; else if(srcStr==='INTERNAL')mb.int_del++;
          }
        }
        if (rej) {
          orgMap[org].tRej++; orgMap[org].val+=val; if(mo)orgMap[org].rej[mo-1]++;
          if(mb){
            mb.tRej++; mb.val+=val;
            if(root) mb.reasons[root]=(mb.reasons[root]||0)+1;
            if(cust) mb.custs[cust]=(mb.custs[cust]||0)+1;
            if(cust||root){
              var monthDetailKey = (cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
              mb.detail[monthDetailKey]=(mb.detail[monthDetailKey]||0)+1;
              if(isFood)  mb.food_detail[monthDetailKey]=(mb.food_detail[monthDetailKey]||0)+1;
              if(isNF)    mb.nonfood_detail[monthDetailKey]=(mb.nonfood_detail[monthDetailKey]||0)+1;
              if(srcStr==='EXTERNAL') mb.ext_detail[monthDetailKey]=(mb.ext_detail[monthDetailKey]||0)+1;
              if(srcStr==='INTERNAL') mb.int_detail[monthDetailKey]=(mb.int_detail[monthDetailKey]||0)+1;
            }
            if(isFood){ mb.food_rej++; mb.food_val+=val; if(root)mb.food_reasons[root]=(mb.food_reasons[root]||0)+1; if(cust)mb.food_custs[cust]=(mb.food_custs[cust]||0)+1; }
            else if(isNF){ mb.nonfood_rej++; mb.nonfood_val+=val; if(root)mb.nonfood_reasons[root]=(mb.nonfood_reasons[root]||0)+1; if(cust)mb.nonfood_custs[cust]=(mb.nonfood_custs[cust]||0)+1; }
            if(srcStr==='EXTERNAL'){ mb.ext_rej++; if(root)mb.ext_reasons[root]=(mb.ext_reasons[root]||0)+1; if(cust)mb.ext_custs[cust]=(mb.ext_custs[cust]||0)+1; }
            else if(srcStr==='INTERNAL'){ mb.int_rej++; if(root)mb.int_reasons[root]=(mb.int_reasons[root]||0)+1; if(cust)mb.int_custs[cust]=(mb.int_custs[cust]||0)+1; }
          }
          if(isFood){orgMap[org].food_rej++;orgMap[org].food_val+=val;}
          else if(isNF){orgMap[org].nonfood_rej++;orgMap[org].nonfood_val+=val;}
          if(srcStr==='EXTERNAL')orgMap[org].ext_rej++; else if(srcStr==='INTERNAL')orgMap[org].int_rej++;
          if(root)orgMap[org].reasons[root]=(orgMap[org].reasons[root]||0)+1;
          if(cust)orgMap[org].custs[cust]=(orgMap[org].custs[cust]||0)+1;
          if(area)orgMap[org].areas[area]=(orgMap[org].areas[area]||0)+1;
          // Detail: customer+address+rootcause combo
          if(cust||root){
            var detailKey = (cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
            orgMap[org].detail[detailKey]=(orgMap[org].detail[detailKey]||0)+1;
            if(isFood)  orgMap[org].food_detail[detailKey]=(orgMap[org].food_detail[detailKey]||0)+1;
            if(isNF)    orgMap[org].nonfood_detail[detailKey]=(orgMap[org].nonfood_detail[detailKey]||0)+1;
            if(srcStr==='EXTERNAL') orgMap[org].ext_detail[detailKey]=(orgMap[org].ext_detail[detailKey]||0)+1;
            if(srcStr==='INTERNAL') orgMap[org].int_detail[detailKey]=(orgMap[org].int_detail[detailKey]||0)+1;
          }
          // Per-type breakdown
          if(isFood){
            if(root)orgMap[org].food_reasons[root]=(orgMap[org].food_reasons[root]||0)+1;
            if(cust)orgMap[org].food_custs[cust]=(orgMap[org].food_custs[cust]||0)+1;
          } else if(isNF){
            if(root)orgMap[org].nonfood_reasons[root]=(orgMap[org].nonfood_reasons[root]||0)+1;
            if(cust)orgMap[org].nonfood_custs[cust]=(orgMap[org].nonfood_custs[cust]||0)+1;
          }
          // Per-source breakdown
          if(srcStr==='EXTERNAL'){
            if(root)orgMap[org].ext_reasons[root]=(orgMap[org].ext_reasons[root]||0)+1;
            if(cust)orgMap[org].ext_custs[cust]=(orgMap[org].ext_custs[cust]||0)+1;
          } else if(srcStr==='INTERNAL'){
            if(root)orgMap[org].int_reasons[root]=(orgMap[org].int_reasons[root]||0)+1;
            if(cust)orgMap[org].int_custs[cust]=(orgMap[org].int_custs[cust]||0)+1;
          }
        }
      }
      if (mo) {
        if (!monthMap[mo]) monthMap[mo]={days:{},tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{},food_reasons:{},food_custs:{},nonfood_reasons:{},nonfood_custs:{},ext_reasons:{},ext_custs:{},int_reasons:{},int_custs:{},food_rej:0,nonfood_rej:0,ext_rej:0,int_rej:0,detail:{},food_detail:{},nonfood_detail:{},ext_detail:{},int_detail:{},data:{}};
        if (del) monthMap[mo].tDel++;
        if (rej) {
          monthMap[mo].tRej++; monthMap[mo].val+=val;
          if(root)monthMap[mo].reasons[root]=(monthMap[mo].reasons[root]||0)+1;
          if(cust)monthMap[mo].custs[cust]=(monthMap[mo].custs[cust]||0)+1;
          if(area)monthMap[mo].areas[area]=(monthMap[mo].areas[area]||0)+1;
          if(day)monthMap[mo].days[day]=1;
          // Per-type breakdown for month
          if(isFood){
            monthMap[mo].food_rej++;
            if(root)monthMap[mo].food_reasons[root]=(monthMap[mo].food_reasons[root]||0)+1;
            if(cust)monthMap[mo].food_custs[cust]=(monthMap[mo].food_custs[cust]||0)+1;
          } else if(isNF){
            monthMap[mo].nonfood_rej++;
            if(root)monthMap[mo].nonfood_reasons[root]=(monthMap[mo].nonfood_reasons[root]||0)+1;
            if(cust)monthMap[mo].nonfood_custs[cust]=(monthMap[mo].nonfood_custs[cust]||0)+1;
          }
          // Per-source breakdown for month
          if(srcStr==='EXTERNAL'){
            monthMap[mo].ext_rej++;
            if(root)monthMap[mo].ext_reasons[root]=(monthMap[mo].ext_reasons[root]||0)+1;
            if(cust)monthMap[mo].ext_custs[cust]=(monthMap[mo].ext_custs[cust]||0)+1;
          } else if(srcStr==='INTERNAL'){
            monthMap[mo].int_rej++;
            if(root)monthMap[mo].int_reasons[root]=(monthMap[mo].int_reasons[root]||0)+1;
            if(cust)monthMap[mo].int_custs[cust]=(monthMap[mo].int_custs[cust]||0)+1;
          }
          // Month detail
          if(cust||root){
            var mdk=(cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
            monthMap[mo].detail[mdk]=(monthMap[mo].detail[mdk]||0)+1;
            if(isFood)  monthMap[mo].food_detail[mdk]=(monthMap[mo].food_detail[mdk]||0)+1;
            if(isNF)    monthMap[mo].nonfood_detail[mdk]=(monthMap[mo].nonfood_detail[mdk]||0)+1;
            if(srcStr==='EXTERNAL') monthMap[mo].ext_detail[mdk]=(monthMap[mo].ext_detail[mdk]||0)+1;
            if(srcStr==='INTERNAL') monthMap[mo].int_detail[mdk]=(monthMap[mo].int_detail[mdk]||0)+1;
          }
        }
        if (day) {
          if(!monthMap[mo].data[day])monthMap[mo].data[day]={tDel:0,tRej:0,val:0,reasons:{},custs:{},areas:{},food_reasons:{},food_custs:{},nonfood_reasons:{},nonfood_custs:{},ext_reasons:{},ext_custs:{},int_reasons:{},int_custs:{},food_rej:0,nonfood_rej:0,ext_rej:0,int_rej:0,detail:{},food_detail:{},nonfood_detail:{},ext_detail:{},int_detail:{}};
          if(del)monthMap[mo].data[day].tDel++;
          if(rej){
            monthMap[mo].data[day].tRej++; monthMap[mo].data[day].val+=val;
            if(root)monthMap[mo].data[day].reasons[root]=(monthMap[mo].data[day].reasons[root]||0)+1;
            if(cust)monthMap[mo].data[day].custs[cust]=(monthMap[mo].data[day].custs[cust]||0)+1;
            if(area)monthMap[mo].data[day].areas[area]=(monthMap[mo].data[day].areas[area]||0)+1;
            if(isFood){
              monthMap[mo].data[day].food_rej++;
              if(root)monthMap[mo].data[day].food_reasons[root]=(monthMap[mo].data[day].food_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].food_custs[cust]=(monthMap[mo].data[day].food_custs[cust]||0)+1;
            } else if(isNF){
              monthMap[mo].data[day].nonfood_rej++;
              if(root)monthMap[mo].data[day].nonfood_reasons[root]=(monthMap[mo].data[day].nonfood_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].nonfood_custs[cust]=(monthMap[mo].data[day].nonfood_custs[cust]||0)+1;
            }
            if(srcStr==='EXTERNAL'){
              monthMap[mo].data[day].ext_rej++;
              if(root)monthMap[mo].data[day].ext_reasons[root]=(monthMap[mo].data[day].ext_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].ext_custs[cust]=(monthMap[mo].data[day].ext_custs[cust]||0)+1;
            } else if(srcStr==='INTERNAL'){
              monthMap[mo].data[day].int_rej++;
              if(root)monthMap[mo].data[day].int_reasons[root]=(monthMap[mo].data[day].int_reasons[root]||0)+1;
              if(cust)monthMap[mo].data[day].int_custs[cust]=(monthMap[mo].data[day].int_custs[cust]||0)+1;
            }
            // Day detail
            if(cust||root){
              var ddk=(cust||'Unknown')+'|||'+(addr||'No address')+'|||'+(root||'Unknown');
              monthMap[mo].data[day].detail[ddk]=(monthMap[mo].data[day].detail[ddk]||0)+1;
              if(isFood)  monthMap[mo].data[day].food_detail[ddk]=(monthMap[mo].data[day].food_detail[ddk]||0)+1;
              if(isNF)    monthMap[mo].data[day].nonfood_detail[ddk]=(monthMap[mo].data[day].nonfood_detail[ddk]||0)+1;
              if(srcStr==='EXTERNAL') monthMap[mo].data[day].ext_detail[ddk]=(monthMap[mo].data[day].ext_detail[ddk]||0)+1;
              if(srcStr==='INTERNAL') monthMap[mo].data[day].int_detail[ddk]=(monthMap[mo].data[day].int_detail[ddk]||0)+1;
            }
          }
        }
      }

}

// ── Shared post-processing: turns raw orgMap/monthMap totals into the
// exact JSON shape the dashboard expects (top10 reasons, top customers,
// detail breakdowns, etc). Used by both the upload endpoint and the row
// correction endpoint, so a manual edit produces identically-shaped output
// to a fresh upload. ──
function buildDashboardOutputs(orgMap, monthMap, totalRej, totalDel, totalVal) {
    function fmtVal(v){return v>=1000000?'AED '+(v/1000000).toFixed(2)+'M':'AED '+Math.round(v/1000)+'K';}
    function top10(obj){return Object.keys(obj).map(function(l){return{l:l,n:obj[l]};}).sort(function(a,b){return b.n-a.n;}).slice(0,10);}
    function top8c(obj){return Object.keys(obj).map(function(n){return{n:n,c:obj[n],v:''};}).sort(function(a,b){return b.c-a.c;}).slice(0,20);}
    function top6a(obj){return Object.keys(obj).map(function(a){return{a:a,n:obj[a]};}).sort(function(a,b){return b.n-a.n;}).slice(0,6);}
    function topDetail(obj){return Object.keys(obj).map(function(k){var p=k.split('|||');return{cust:p[0],addr:p[1],root:p[2],n:obj[k]};}).sort(function(a,b){return b.n-a.n;}).slice(0,100);}

    var allR={},allC={},allA={},allDetail={},allFoodDetail={},allNFDetail={},allExtDetail={},allIntDetail={},allDel=new Array(12).fill(0),allRej=new Array(12).fill(0);
    var allFoodRej=0,allNFRej=0,allExtRej=0,allIntRej=0,allFoodDel=0,allNFDel=0,allFoodVal=0,allNFVal=0;
    var allFoodR={},allFoodC={},allNFR={},allNFC={},allExtR={},allExtC={},allIntR={},allIntC={};
    Object.keys(orgMap).forEach(function(org){
      var v=orgMap[org];
      Object.keys(v.reasons).forEach(function(k){allR[k]=(allR[k]||0)+v.reasons[k];});
      Object.keys(v.custs).forEach(function(k){allC[k]=(allC[k]||0)+v.custs[k];});
      Object.keys(v.areas).forEach(function(k){allA[k]=(allA[k]||0)+v.areas[k];});
      Object.keys(v.detail||{}).forEach(function(k){allDetail[k]=(allDetail[k]||0)+v.detail[k];});
      Object.keys(v.food_detail||{}).forEach(function(k){allFoodDetail[k]=(allFoodDetail[k]||0)+v.food_detail[k];});
      Object.keys(v.nonfood_detail||{}).forEach(function(k){allNFDetail[k]=(allNFDetail[k]||0)+v.nonfood_detail[k];});
      Object.keys(v.ext_detail||{}).forEach(function(k){allExtDetail[k]=(allExtDetail[k]||0)+v.ext_detail[k];});
      Object.keys(v.int_detail||{}).forEach(function(k){allIntDetail[k]=(allIntDetail[k]||0)+v.int_detail[k];});
      Object.keys(v.food_reasons||{}).forEach(function(k){allFoodR[k]=(allFoodR[k]||0)+v.food_reasons[k];});
      Object.keys(v.food_custs||{}).forEach(function(k){allFoodC[k]=(allFoodC[k]||0)+v.food_custs[k];});
      Object.keys(v.nonfood_reasons||{}).forEach(function(k){allNFR[k]=(allNFR[k]||0)+v.nonfood_reasons[k];});
      Object.keys(v.nonfood_custs||{}).forEach(function(k){allNFC[k]=(allNFC[k]||0)+v.nonfood_custs[k];});
      Object.keys(v.ext_reasons||{}).forEach(function(k){allExtR[k]=(allExtR[k]||0)+v.ext_reasons[k];});
      Object.keys(v.ext_custs||{}).forEach(function(k){allExtC[k]=(allExtC[k]||0)+v.ext_custs[k];});
      Object.keys(v.int_reasons||{}).forEach(function(k){allIntR[k]=(allIntR[k]||0)+v.int_reasons[k];});
      Object.keys(v.int_custs||{}).forEach(function(k){allIntC[k]=(allIntC[k]||0)+v.int_custs[k];});
      v.del.forEach(function(d,i){allDel[i]+=d;}); v.rej.forEach(function(r,i){allRej[i]+=r;});
      allFoodRej+=(v.food_rej||0); allNFRej+=(v.nonfood_rej||0);
      allExtRej+=(v.ext_rej||0); allIntRej+=(v.int_rej||0);
      allFoodDel+=(v.food_del||0); allNFDel+=(v.nonfood_del||0);
      allFoodVal+=(v.food_val||0); allNFVal+=(v.nonfood_val||0);
    });

    var monthsOut={};
    Object.keys(monthMap).forEach(function(mo){
      var md=monthMap[mo]; var dataOut={};
      Object.keys(md.data).forEach(function(day){
        var dd=md.data[day];
        dataOut[day]={tDel:dd.tDel,tRej:dd.tRej,val:fmtVal(dd.val),reasons:top10(dd.reasons),custs:top8c(dd.custs),areas:top6a(dd.areas),food_rej:dd.food_rej||0,nonfood_rej:dd.nonfood_rej||0,ext_rej:dd.ext_rej||0,int_rej:dd.int_rej||0,food_reasons:top10(dd.food_reasons||{}),food_custs:top8c(dd.food_custs||{}),nonfood_reasons:top10(dd.nonfood_reasons||{}),nonfood_custs:top8c(dd.nonfood_custs||{}),ext_reasons:top10(dd.ext_reasons||{}),ext_custs:top8c(dd.ext_custs||{}),int_reasons:top10(dd.int_reasons||{}),int_custs:top8c(dd.int_custs||{}),detail:topDetail(dd.detail||{}),food_detail:topDetail(dd.food_detail||{}),nonfood_detail:topDetail(dd.nonfood_detail||{}),ext_detail:topDetail(dd.ext_detail||{}),int_detail:topDetail(dd.int_detail||{})};
      });
      monthsOut[mo]={days:Object.keys(md.days).map(Number).sort(function(a,b){return a-b;}),tDel:md.tDel,tRej:md.tRej,val:fmtVal(md.val),food_rej:md.food_rej||0,nonfood_rej:md.nonfood_rej||0,ext_rej:md.ext_rej||0,int_rej:md.int_rej||0,reasons:top10(md.reasons),custs:top8c(md.custs||{}),areas:top6a(md.areas||{}),food_reasons:top10(md.food_reasons||{}),food_custs:top8c(md.food_custs||{}),nonfood_reasons:top10(md.nonfood_reasons||{}),nonfood_custs:top8c(md.nonfood_custs||{}),ext_reasons:top10(md.ext_reasons||{}),ext_custs:top8c(md.ext_custs||{}),int_reasons:top10(md.int_reasons||{}),int_custs:top8c(md.int_custs||{}),detail:topDetail(md.detail||{}),food_detail:topDetail(md.food_detail||{}),nonfood_detail:topDetail(md.nonfood_detail||{}),ext_detail:topDetail(md.ext_detail||{}),int_detail:topDetail(md.int_detail||{}),data:dataOut};
    });

    var orgsOut={all:{tDel:totalDel,tRej:totalRej,val:fmtVal(totalVal),food_rej:allFoodRej,food_del:allFoodDel,nonfood_rej:allNFRej,nonfood_del:allNFDel,ext_rej:allExtRej,int_rej:allIntRej,food_val:fmtVal(allFoodVal),nonfood_val:fmtVal(allNFVal),del:allDel,rej:allRej,reasons:top10(allR),custs:top8c(allC),areas:top6a(allA),detail:topDetail(allDetail),food_detail:topDetail(allFoodDetail||{}),nonfood_detail:topDetail(allNFDetail||{}),ext_detail:topDetail(allExtDetail||{}),int_detail:topDetail(allIntDetail||{}),food_reasons:top10(allFoodR),food_custs:top8c(allFoodC),nonfood_reasons:top10(allNFR),nonfood_custs:top8c(allNFC),ext_reasons:top10(allExtR),ext_custs:top8c(allExtC),int_reasons:top10(allIntR),int_custs:top8c(allIntC)}};
    Object.keys(orgMap).forEach(function(org){
      var v=orgMap[org];
      var byMonthOut = {};
      if (v.byMonth) {
        Object.keys(v.byMonth).forEach(function(mo){
          var mData = v.byMonth[mo];
          byMonthOut[mo] = {
            tRej: mData.tRej||0, tDel: mData.tDel||0, val: fmtVal(mData.val||0),
            food_rej: mData.food_rej||0, food_del: mData.food_del||0, food_val: fmtVal(mData.food_val||0),
            nonfood_rej: mData.nonfood_rej||0, nonfood_del: mData.nonfood_del||0, nonfood_val: fmtVal(mData.nonfood_val||0),
            ext_rej: mData.ext_rej||0, ext_del: mData.ext_del||0,
            int_rej: mData.int_rej||0, int_del: mData.int_del||0,
            reasons: top10(mData.reasons||{}), custs: top8c(mData.custs||{}), detail: topDetail(mData.detail||{}),
            food_reasons: top10(mData.food_reasons||{}), food_custs: top8c(mData.food_custs||{}), food_detail: topDetail(mData.food_detail||{}),
            nonfood_reasons: top10(mData.nonfood_reasons||{}), nonfood_custs: top8c(mData.nonfood_custs||{}), nonfood_detail: topDetail(mData.nonfood_detail||{}),
            ext_reasons: top10(mData.ext_reasons||{}), ext_custs: top8c(mData.ext_custs||{}), ext_detail: topDetail(mData.ext_detail||{}),
            int_reasons: top10(mData.int_reasons||{}), int_custs: top8c(mData.int_custs||{}), int_detail: topDetail(mData.int_detail||{})
          };
        });
      }
      orgsOut[org]={tDel:v.tDel,tRej:v.tRej,val:fmtVal(v.val),food_rej:v.food_rej||0,food_del:v.food_del||0,nonfood_rej:v.nonfood_rej||0,nonfood_del:v.nonfood_del||0,ext_rej:v.ext_rej||0,int_rej:v.int_rej||0,food_val:fmtVal(v.food_val||0),nonfood_val:fmtVal(v.nonfood_val||0),del:v.del,rej:v.rej,reasons:top10(v.reasons),custs:top8c(v.custs),areas:top6a(v.areas),detail:topDetail(v.detail||{}),food_detail:topDetail(v.food_detail||{}),nonfood_detail:topDetail(v.nonfood_detail||{}),ext_detail:topDetail(v.ext_detail||{}),int_detail:topDetail(v.int_detail||{}),food_reasons:top10(v.food_reasons||{}),food_custs:top8c(v.food_custs||{}),nonfood_reasons:top10(v.nonfood_reasons||{}),nonfood_custs:top8c(v.nonfood_custs||{}),ext_reasons:top10(v.ext_reasons||{}),ext_custs:top8c(v.ext_custs||{}),int_reasons:top10(v.int_reasons||{}),int_custs:top8c(v.int_custs||{}),byMonth:byMonthOut};
    });

  return { orgsOut: orgsOut, monthsOut: monthsOut };
}

// ── Rebuilds the entire rejection aggregate from EVERY row currently stored
// in rejection_rows (all months combined). Called after an upload persists
// new rows, and after a manual correction edits one row — both paths must
// produce identical output, so both just call this. ──
async function rebuildRejectionAggregateFromDB() {
  var orgMap = {}, monthMap = {};
  var totalRej = 0, totalDel = 0, totalVal = 0;
  var seenOrderVals = {};
  var allDbRows = [];
  try {
    var allRes = await pool.query('SELECT * FROM rejection_rows');
    allDbRows = allRes.rows;
  } catch (e) { console.error('Rejection: could not load rejection_rows to rebuild aggregate:', e.message); }

  allDbRows.forEach(function(hr) {
    var rej = hr.status === 'rej', del = hr.status === 'del';
    if (!rej && !del) return;
    var d = hr.entry_date ? new Date(hr.entry_date) : null;
    var mo = d ? d.getMonth() + 1 : null, day = d ? d.getDate() : null;
    var rawVal = parseFloat(hr.value) || 0;
    var val = (rej && hr.order_no && seenOrderVals[hr.order_no]) ? 0 : rawVal;
    if (rej && hr.order_no && !seenOrderVals[hr.order_no]) seenOrderVals[hr.order_no] = rawVal;
    if (del) totalDel++;
    if (rej) { totalRej++; totalVal += val; }
    // Self-heal: any row stored with a blank source (from before this default
    // existed) falls back to Internal here too, same reasoning as at upload
    // time — never leave a rejected row invisible to both the External and
    // Internal breakdowns.
    var srcStr = (hr.source || '').toUpperCase();
    if (!srcStr && rej) srcStr = 'INTERNAL';
    accumulateRejectionRow(orgMap, monthMap, {
      org: hr.org, root: hr.root_cause, cust: hr.customer_name, addr: hr.address, area: hr.area,
      mo: mo, day: day, isFood: hr.is_food, isNF: hr.is_nonfood, srcStr: srcStr,
      val: val, rej: rej, del: del
    });
  });

  var out = buildDashboardOutputs(orgMap, monthMap, totalRej, totalDel, totalVal);
  return { orgMap: orgMap, monthMap: monthMap, totalRej: totalRej, totalDel: totalDel, totalVal: totalVal, orgsOut: out.orgsOut, monthsOut: out.monthsOut };
}

function saveJSON(fp, data) {
  try { fs.writeFileSync(fp, JSON.stringify(data)); } catch(e) { console.error('File save error:', e.message); }
}
function loadJSON(fp) {
  try { if (fs.existsSync(fp)) return JSON.parse(fs.readFileSync(fp, 'utf8')); } catch(e) {}
  return null;
}

//  Init DB tables 
async function initDB() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS dispatch_data (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL UNIQUE,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT DEFAULT 'Admin',
      summary JSONB NOT NULL,
      csv_text TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    // Some days need more than one dispatch file — a Main dispatch report and
    // a separate one for another BU (e.g. Salon, dispatched from a different
    // warehouse). Each source's raw rows are kept here under its own label,
    // so uploading one never overwrites the other; dispatch_data.summary
    // above is always the MERGED result across every source for that date.
    await pool.query(`CREATE TABLE IF NOT EXISTS dispatch_data_sources (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL,
      source_label TEXT NOT NULL,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT DEFAULT 'Admin',
      file_name TEXT,
      rows_json JSONB NOT NULL,
      row_count INT NOT NULL DEFAULT 0,
      UNIQUE(date_key, source_label)
    )`);
    // Transport-team-reported drop counts, uploaded from their own raw drop
    // file (Distinct DROP ID, by date + Bulk/Multi classification). Compared
    // against the app's own dispatch_data.summary.total_drops for the same
    // dates to catch under/over-counting on either side before the monthly
    // invoice arrives.
    //
    // reported_drops = ALL dispatched drops for that date/class, regardless of
    // the transport team's internal TASK STATUS (Completed/Ongoing/Waiting) —
    // a drop that was dispatched already incurred the trip cost whether or not
    // it's been marked done in their system yet. completed_drops is kept only
    // as supplementary context (how much of that day has been closed out),
    // never used to filter the primary count.
    await pool.query(`CREATE TABLE IF NOT EXISTS transport_drop_reconciliation (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL,
      drop_class TEXT NOT NULL,
      reported_drops INT NOT NULL DEFAULT 0,
      completed_drops INT NOT NULL DEFAULT 0,
      task_status_rule TEXT,
      upload_batch_id TEXT,
      uploaded_by TEXT,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      UNIQUE(date_key, drop_class)
    )`);
    await pool.query(`ALTER TABLE transport_drop_reconciliation ADD COLUMN IF NOT EXISTS completed_drops INT NOT NULL DEFAULT 0`);
    // Individual order identifiers from the transport team's file (TASK ID,
    // suffix stripped), per date — lets us diff against this app's own
    // order_tracking table (already populated from daily dispatch uploads)
    // to find the SPECIFIC orders that don't match on either side, not just
    // a day-level count difference.
    await pool.query(`CREATE TABLE IF NOT EXISTS transport_order_ids (
      id SERIAL PRIMARY KEY,
      date_key DATE NOT NULL,
      order_id TEXT NOT NULL,
      upload_batch_id TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      UNIQUE(date_key, order_id)
    )`);
    await pool.query(`ALTER TABLE transport_order_ids ADD COLUMN IF NOT EXISTS customer TEXT`);
    await pool.query(`ALTER TABLE transport_order_ids ADD COLUMN IF NOT EXISTS location_name TEXT`);
    await pool.query(`ALTER TABLE transport_order_ids ADD COLUMN IF NOT EXISTS address TEXT`);
    await pool.query(`CREATE TABLE IF NOT EXISTS rejection_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT DEFAULT 'Admin',
      file_name TEXT,
      total_orders INT,
      orgs JSONB,
      months JSONB,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    // Persist the auto-classification stats + manual-review queue from the last upload,
    // so any superadmin/subadmin can see them later — not just the person who uploaded
    // (previously this only ever lived in the uploader's own browser session).
    await pool.query(`ALTER TABLE rejection_data ADD COLUMN IF NOT EXISTS needs_review JSONB`);
    await pool.query(`ALTER TABLE rejection_data ADD COLUMN IF NOT EXISTS auto_classify JSONB`);
    // ── Individual rejection/delivery rows, kept so uploading one month's
    // transport file (e.g. August) only replaces THAT month — Jan-Jul stays
    // untouched. month_key ('2026-08') scopes the delete-then-insert on
    // upload; order_no is the unique key a manual correction targets. ──
    await pool.query(`CREATE TABLE IF NOT EXISTS rejection_rows (
      id SERIAL PRIMARY KEY,
      month_key TEXT NOT NULL,
      entry_date DATE,
      status TEXT NOT NULL,
      org TEXT,
      root_cause TEXT,
      root_cause_source TEXT,
      customer_name TEXT,
      address TEXT,
      area TEXT,
      order_no TEXT,
      value NUMERIC DEFAULT 0,
      is_food BOOLEAN,
      is_nonfood BOOLEAN,
      source TEXT,
      file_name TEXT,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      corrected_by TEXT,
      corrected_at TIMESTAMPTZ
    )`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_rejection_rows_month ON rejection_rows (month_key)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_rejection_rows_orderno ON rejection_rows (order_no)`);
    // ── Sub-admin upload guardrails: up to 2 uploads per rolling 10-hour
    // window, per (user, endpoint). Superadmin is never gated by this table. ──
    await pool.query(`CREATE TABLE IF NOT EXISTS upload_cooldowns (
      user_id INT NOT NULL,
      endpoint TEXT NOT NULL,
      last_upload_at TIMESTAMPTZ DEFAULT NOW(),
      window_started_at TIMESTAMPTZ DEFAULT NOW(),
      upload_count INT DEFAULT 1,
      PRIMARY KEY (user_id, endpoint)
    )`);
    await pool.query(`ALTER TABLE upload_cooldowns ADD COLUMN IF NOT EXISTS window_started_at TIMESTAMPTZ DEFAULT NOW()`);
    await pool.query(`ALTER TABLE upload_cooldowns ADD COLUMN IF NOT EXISTS upload_count INT DEFAULT 1`);
    // ── Sub-admin requests to overwrite/re-upload dates that already exist on
    // the dashboard get parked here instead of applied — only a super admin can
    // approve them. file_data holds the exact original upload so approval can
    // replay it with zero re-typing; it's cleared out once decided. ──
    await pool.query(`CREATE TABLE IF NOT EXISTS upload_approval_requests (
      id SERIAL PRIMARY KEY,
      user_id INT,
      username TEXT,
      endpoint TEXT NOT NULL,
      file_name TEXT,
      file_data BYTEA,
      blocked_dates JSONB,
      reason TEXT NOT NULL,
      meta JSONB,
      status TEXT DEFAULT 'pending',
      requested_at TIMESTAMPTZ DEFAULT NOW(),
      decided_by TEXT,
      decision_comment TEXT,
      decided_at TIMESTAMPTZ
    )`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_upload_approvals_status ON upload_approval_requests (status)`);
    await pool.query(`CREATE TABLE IF NOT EXISTS backlog_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_orders INT,
      summary JSONB
    )`);
    // Daily tracking history for WH Backlog Risk Command Center. One row per
    // calendar day (Asia/Dubai). Upserted on every backlog upload + a periodic
    // safeguard capture, so the trend stays populated even on days with no upload.
    await pool.query(`CREATE TABLE IF NOT EXISTS wh_backlog_snapshots (
      id SERIAL PRIMARY KEY,
      date_key TEXT UNIQUE NOT NULL,
      total_orders INT DEFAULT 0,
      total_val NUMERIC DEFAULT 0,
      advance JSONB,
      credit_hold JSONB,
      wh_backlog JSONB,
      cut_off_frozen JSONB,
      unplanned_frozen JSONB,
      org_breakdown JSONB,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    // Daily evening snapshot of orders still stuck in credit hold (not yet released,
    // not yet reached the warehouse) — a separate feed from the morning WH Backlog
    // upload, so we can track releases day-over-day: what was on hold yesterday
    // evening that's gone today (released), vs what's still stuck (still_on_hold),
    // vs what's newly appeared (new_holds).
    await pool.query(`CREATE TABLE IF NOT EXISTS credit_hold_tracking (
      id SERIAL PRIMARY KEY,
      date_key TEXT UNIQUE NOT NULL,
      captured_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_count INT DEFAULT 0,
      total_val NUMERIC DEFAULT 0,
      orders JSONB,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      updated_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS returns_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_orders INT,
      summary JSONB
    )`);
    // Row-level returns data, kept alongside the returns_data summary so the
    // Market Returns "Full Analyzed Report" export can use live SUMIF/COUNTIF
    // formulas (same pattern as Customer Visit's rejection_rows / Order Master
    // tab) instead of frozen numbers. Cleared and reloaded on every upload,
    // same lifecycle as returns_data.
    await pool.query(`CREATE TABLE IF NOT EXISTS returns_rows (
      id SERIAL PRIMARY KEY,
      warehouse TEXT,
      category TEXT,
      bu TEXT,
      value NUMERIC DEFAULT 0,
      month INT,
      day INT,
      date_str TEXT,
      customer TEXT,
      area TEXT,
      order_no TEXT,
      reason TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_returns_rows_month ON returns_rows(month)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_returns_rows_bu ON returns_rows(bu)`);
    // Tracks every order code seen per dispatch date, so the same order code appearing
    // again on a LATER date can be detected as a re-delivery (failed first attempt,
    // re-attempted later) — not just a same-day duplicate.
    await pool.query(`CREATE TABLE IF NOT EXISTS order_tracking (
      id SERIAL PRIMARY KEY,
      order_code TEXT NOT NULL,
      date_key DATE NOT NULL,
      customer TEXT,
      value NUMERIC DEFAULT 0,
      route TEXT,
      org TEXT,
      drop_type TEXT,
      temperature TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    await pool.query(`ALTER TABLE order_tracking ADD COLUMN IF NOT EXISTS temperature TEXT`);
    // Tracks which dispatch upload source a row came from (Main, Salon, etc.) so the
    // Transport Cost Reconciliation order-level diff can exclude Salon the same way
    // isCostExcludedRow already does for routes/drops — Salon (DIP warehouse) deliveries
    // run on an unconfirmed transport billing model and structurally never appear in
    // transport's own file, so they'd otherwise always show as a false "missing" order.
    await pool.query(`ALTER TABLE order_tracking ADD COLUMN IF NOT EXISTS source_label TEXT`);
    // Needed so the DCF exclusion can be Dubai-specific: DCF self-delivers on its own van
    // only in Dubai, so DCF orders elsewhere must still be compared against transport's
    // file normally — city is what tells the two cases apart.
    await pool.query(`ALTER TABLE order_tracking ADD COLUMN IF NOT EXISTS city TEXT`);
    // Needed for Transport Cost Control: drop-level (not just order-level) counts, and a
    // per-order rate-card lookup for re-delivery cost estimates.
    await pool.query(`ALTER TABLE order_tracking ADD COLUMN IF NOT EXISTS location_id TEXT`);
    await pool.query(`ALTER TABLE order_tracking ADD COLUMN IF NOT EXISTS truck_type TEXT`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_order_tracking_code ON order_tracking(order_code)`);
    await pool.query(`CREATE INDEX IF NOT EXISTS idx_order_tracking_date ON order_tracking(date_key)`);
    // AUTH TABLES
    await pool.query(`CREATE TABLE IF NOT EXISTS users (
      id SERIAL PRIMARY KEY,
      username TEXT UNIQUE NOT NULL,
      password_hash TEXT NOT NULL,
      full_name TEXT,
      role TEXT NOT NULL DEFAULT 'viewer',
      dashboards JSONB DEFAULT '["dispatch","rejection","summary","email","invoice","backlog","returns","sales","automation"]'::jsonb,
      active BOOLEAN DEFAULT true,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      created_by TEXT DEFAULT 'system',
      last_login TIMESTAMPTZ,
      must_change_password BOOLEAN DEFAULT false
    )`);
    // Lets a superadmin temporarily lift the sub-admin 2-per-10-hour upload cooldown for
    // one user — e.g. someone catching up on several missing days needs to upload more
    // than twice in a stretch. NULL/past = no exemption (normal cooldown applies).
    await pool.query(`ALTER TABLE users ADD COLUMN IF NOT EXISTS upload_exempt_until TIMESTAMPTZ`);
    await pool.query(`CREATE TABLE IF NOT EXISTS sessions (
      id SERIAL PRIMARY KEY,
      token TEXT UNIQUE NOT NULL,
      user_id INT REFERENCES users(id) ON DELETE CASCADE,
      created_at TIMESTAMPTZ DEFAULT NOW(),
      expires_at TIMESTAMPTZ NOT NULL,
      ip_address TEXT
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS audit_log (
      id SERIAL PRIMARY KEY,
      user_id INT,
      username TEXT,
      action TEXT NOT NULL,
      details TEXT,
      ip_address TEXT,
      created_at TIMESTAMPTZ DEFAULT NOW()
    )`);
    // Create or update default super admin
    var adminCheck = await pool.query("SELECT id FROM users WHERE username = 'azhar'");
    var hash = await bcrypt.hash('YAmaha100@', 10);
    if (adminCheck.rows.length === 0) {
      await pool.query(
        "INSERT INTO users (username, password_hash, full_name, role) VALUES ($1,$2,$3,$4)",
        ['azhar', hash, 'Mohammed Azharuddin', 'superadmin']
      );
      console.log('Default super admin created: azhar / YAmaha100@');
    } else {
      // Always sync password with code on server start
      await pool.query("UPDATE users SET password_hash=$1, active=true WHERE username='azhar'", [hash]);
      console.log('Super admin password synced: azhar / YAmaha100@');
    }
    console.log('DB tables ready');

    // ── One-time data fix: "Declined — LPO Date Expired" was learned as
    // Customer Related (External) from historical data, but it's actually
    // caused by our own transport running late and the LPO expiring before
    // delivery could happen — an AKI-side (Internal) issue, not the
    // customer's fault. Idempotent: only touches rows still marked
    // External for this cause, so it's a no-op on every run after the
    // first. Also fixed going forward in rejection_lookups.json. ──
    try {
      var lpoFix = await pool.query(
        `UPDATE rejection_rows SET source='INTERNAL', root_cause_source='manual', corrected_by='system-reclassify', corrected_at=NOW()
         WHERE UPPER(TRIM(root_cause)) = 'DECLINED — LPO DATE EXPIRED' AND UPPER(source) = 'EXTERNAL'
         RETURNING id`
      );
      if (lpoFix.rows.length) {
        console.log('Reclassified', lpoFix.rows.length, '"Declined — LPO Date Expired" rows from External to Internal');
        await rebuildRejectionAggregateFromDB();
      }
    } catch (e) {
      console.error('LPO Date Expired reclassification error:', e.message);
    }

    // ── One-time data fix: any "Driver Spend N Hours In <location> Due To
    // Late Receiving" root cause describes a delay at a DIFFERENT stop on
    // the driver's route (a different customer or hub entirely) — that's
    // an internal routing/scheduling problem, never something the customer
    // whose delivery it ultimately was did. These were showing up as
    // External for whichever customer's order happened to be late.
    // Idempotent — a no-op after the first successful run. ──
    try {
      var driverSpendFix = await pool.query(
        `UPDATE rejection_rows SET source='INTERNAL', root_cause_source='manual', corrected_by='system-reclassify', corrected_at=NOW()
         WHERE root_cause ILIKE 'Driver Spend%' AND UPPER(source) = 'EXTERNAL'
         RETURNING id`
      );
      if (driverSpendFix.rows.length) {
        console.log('Reclassified', driverSpendFix.rows.length, '"Driver Spend..." rows from External to Internal');
        await rebuildRejectionAggregateFromDB();
      }
    } catch (e) {
      console.error('Driver Spend reclassification error:', e.message);
    }
  } catch(e) {
    console.error('DB init error:', e.message);
  }
}
initDB();

//  DB helpers 
async function dbSaveDispatch(dateKey, uploadedBy, summary, csvText) {
  try {
    await pool.query(`
      INSERT INTO dispatch_data (date_key, uploaded_by, summary, csv_text)
      VALUES ($1, $2, $3, $4)
      ON CONFLICT (date_key) DO UPDATE
      SET uploaded_by=$2, summary=$3, csv_text=$4, uploaded_at=NOW()
    `, [dateKey, uploadedBy, JSON.stringify(summary), csvText?.substring(0, 500000)]);
    return true;
  } catch(e) {
    console.error('DB save dispatch error:', e.message);
    return false;
  }
}

async function dbLoadDispatch() {
  try {
    // Keep 2 years of daily dispatch history loaded into memory — plenty
    // for any lookback the dashboard needs, without growing unbounded
    // forever. Older rows stay in the database untouched; they just don't
    // load into memory. Filtered by date rather than row count so it's a
    // stable 2-year window regardless of how many days end up on file.
    const res = await pool.query(`SELECT date_key::text, uploaded_at, uploaded_by, summary, csv_text FROM dispatch_data WHERE date_key >= (CURRENT_DATE - INTERVAL '2 years') ORDER BY date_key DESC`);
    return res.rows;
  } catch(e) {
    console.error('DB load dispatch error:', e.message);
    return [];
  }
}

async function dbSaveRejection(uploadedBy, fileName, totalOrders, orgs, months, needsReview, autoClassify) {
  try {
    await pool.query(`DELETE FROM rejection_data`);
    await pool.query(`
      INSERT INTO rejection_data (uploaded_by, file_name, total_orders, orgs, months, needs_review, auto_classify)
      VALUES ($1, $2, $3, $4, $5, $6, $7)
    `, [uploadedBy, fileName, totalOrders, orgs, months, JSON.stringify(needsReview||[]), JSON.stringify(autoClassify||null)]);
    return true;
  } catch(e) {
    console.error('DB save rejection error:', e.message);
    return false;
  }
}

async function dbLoadRejection() {
  try {
    const res = await pool.query(`SELECT * FROM rejection_data ORDER BY created_at DESC LIMIT 1`);
    return res.rows[0] || null;
  } catch(e) {
    console.error('DB load rejection error:', e.message);
    return null;
  }
}


var backlogData = null;
var BACKLOG_FILE = path.join(DATA_DIR, 'backlog.json');

async function dbSaveBacklog(uploadedBy, fileName, totalOrders, summary) {
  try {
    await pool.query('DELETE FROM backlog_data');
    await pool.query(
      'INSERT INTO backlog_data (uploaded_by, file_name, total_orders, summary) VALUES ($1, $2, $3, $4)',
      [uploadedBy, fileName, totalOrders, JSON.stringify(summary)]
    );
    return true;
  } catch(e) {
    console.error('DB save backlog error:', e.message);
    return false;
  }
}

async function loadBacklogFromDB() {
  try {
    var res = await pool.query('SELECT * FROM backlog_data ORDER BY uploaded_at DESC LIMIT 1');
    if (res.rows[0]) {
      backlogData = { uploadedAt: res.rows[0].uploaded_at, uploadedBy: res.rows[0].uploaded_by, fileName: res.rows[0].file_name, totalOrders: res.rows[0].total_orders, summary: res.rows[0].summary };
      console.log('Loaded backlog from DB');
      return true;
    }
  } catch(e) { console.error('DB load backlog:', e.message); }
  var saved = loadJSON(BACKLOG_FILE);
  if (saved) { backlogData = saved; console.log('Loaded backlog from file'); }
  return false;
}
loadBacklogFromDB();

// Asia/Dubai is UTC+4 year-round (no DST) — used to key one snapshot row per day
function dubaiDateKey(d) {
  d = d || new Date();
  var utcMs = d.getTime() + (d.getTimezoneOffset() * 60000);
  var dubai = new Date(utcMs + 4 * 3600000);
  return dubai.toISOString().slice(0, 10);
}

// Builds the daily tracking snapshot from whatever backlog summary is currently
// live in memory, then upserts it into wh_backlog_snapshots keyed by today's date.
// Safe to call repeatedly (e.g. after every upload, and on a periodic safeguard
// timer) since ON CONFLICT just refreshes today's row.
async function captureBacklogSnapshot() {
  try {
    if (!backlogData || !backlogData.summary) return;
    var s = backlogData.summary;
    var rows = s.actionRows || [];
    var dateKey = dubaiDateKey();

    function bucket(filterFn) {
      var count = 0, val = 0, orgMap = {};
      rows.forEach(function(r) {
        if (!filterFn(r)) return;
        count++; val += (r.val || 0);
        var o = r.org || 'Unknown';
        if (!orgMap[o]) orgMap[o] = { count: 0, val: 0 };
        orgMap[o].count++; orgMap[o].val += (r.val || 0);
      });
      return { count: count, val: val, orgMap: orgMap };
    }

    var advance = bucket(function(r) { return r.cat === 'RSD >30 Advance'; });
    var whBacklogBucket = bucket(function(r) { return r.cat === 'Before Cut-Off'; });
    var creditHold = bucket(function(r) { return /CREDIT HOLD|AUTO BOOKED|LATE RELEAS/i.test(r.rawStatus || ''); });
    var cutOffFrozen = bucket(function(r) { return r.cat === 'Cut-Off Frozen' && !/CREDIT HOLD/i.test(r.rawStatus || ''); });
    var unplannedFrozen = bucket(function(r) { return r.cat === 'Unplanned Frozen'; });

    var orgBreakdown = {};
    rows.forEach(function(r) {
      var o = r.org || 'Unknown';
      if (!orgBreakdown[o]) orgBreakdown[o] = { count: 0, val: 0 };
      orgBreakdown[o].count++; orgBreakdown[o].val += (r.val || 0);
    });

    await pool.query(`
      INSERT INTO wh_backlog_snapshots (date_key, total_orders, total_val, advance, credit_hold, wh_backlog, cut_off_frozen, unplanned_frozen, org_breakdown, updated_at)
      VALUES ($1,$2,$3,$4,$5,$6,$7,$8,$9,NOW())
      ON CONFLICT (date_key) DO UPDATE SET
        total_orders=$2, total_val=$3, advance=$4, credit_hold=$5, wh_backlog=$6, cut_off_frozen=$7, unplanned_frozen=$8, org_breakdown=$9, updated_at=NOW()
    `, [dateKey, s.totalOrders || 0, s.totalVal || 0, JSON.stringify(advance), JSON.stringify(creditHold), JSON.stringify(whBacklogBucket), JSON.stringify(cutOffFrozen), JSON.stringify(unplannedFrozen), JSON.stringify(orgBreakdown)]);
  } catch (e) {
    console.error('captureBacklogSnapshot error:', e.message);
  }
}
// Periodic safeguard: keeps today's row populated even if no file gets uploaded
// on a given day. Idempotent upsert, so this never creates duplicate rows.
setInterval(function() { captureBacklogSnapshot(); }, 30 * 60 * 1000);
setTimeout(function() { captureBacklogSnapshot(); }, 10000);

app.get('/api/backlog/history', requireAuth, async function(req, res) {
  try {
    var days = parseInt(req.query.days) || 30;
    if (days > 365) days = 365;
    var r = await pool.query(
      'SELECT date_key, total_orders, total_val, advance, credit_hold, wh_backlog, cut_off_frozen, unplanned_frozen, org_breakdown, updated_at FROM wh_backlog_snapshots ORDER BY date_key DESC LIMIT $1',
      [days]
    );
    res.json({ success: true, rows: r.rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Full-color Executive Summary + Advance Orders / Credit Hold / WH Backlog workbook.
// "Credit Hold" = orders auto-booked after cut-off that are later released from credit
// hold (matches on raw status text: CREDIT HOLD / AUTO BOOKED / LATE RELEASE), not the
// generic "Cut-Off Frozen" bucket (which also covers temperature/stock freezes).
app.get('/api/backlog/export', noCache, requireAuth, async function(req, res) {
  try {
    if (!backlogData || !backlogData.summary) return res.status(400).json({ error: 'No backlog data loaded yet.' });
    var ExcelJS = require('exceljs');
    var s = backlogData.summary;
    var rows = s.actionRows || [];

    var NAVY = 'FF1B2338', GOLD = 'FFC9A84C', RED = 'FFE84B4B', BLUE = 'FF5B8DEE';
    var RISK_LABELS = { DCV: 'High Volume', DGC: 'Medium', DSN: 'Low', DCF: 'Low' };
    var RISK_FILL = { 'High Volume': 'FFF8D7DA', 'Medium': 'FFFCF3CF', 'Low': 'FFD4EDDA' };
    var RISK_FONT = { 'High Volume': 'FF842029', 'Medium': 'FF7D6608', 'Low': 'FF155724' };

    function headerRow(ws, cells) {
      var r = ws.addRow(cells);
      r.eachCell(function(cell) {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } };
      });
      return r;
    }

    function bucket(filterFn) {
      var list = rows.filter(filterFn);
      var orgMap = {}, totalVal = 0;
      list.forEach(function(rw) {
        var o = rw.org || 'Unknown';
        if (!orgMap[o]) orgMap[o] = { count: 0, val: 0 };
        orgMap[o].count++; orgMap[o].val += (rw.val || 0); totalVal += (rw.val || 0);
      });
      return { list: list, orgMap: orgMap, totalVal: totalVal };
    }

    var advance = bucket(function(r) { return r.cat === 'RSD >30 Advance'; });
    var whBacklogBucket = bucket(function(r) { return r.cat === 'Before Cut-Off'; });
    var creditHold = bucket(function(r) { return /CREDIT HOLD|AUTO BOOKED|LATE RELEAS/i.test(r.rawStatus || ''); });

    var cm = s.catMap || {};
    var whBL = cm['Before Cut-Off'] || { count: 0, val: 0 };
    var advOrders = cm['RSD >30 Advance'] || { count: 0, val: 0 };
    var cutFrozen = cm['Cut-Off Frozen'] || { count: 0, val: 0 };
    var unplanFrozen = cm['Unplanned Frozen'] || { count: 0, val: 0 };
    // "Cut-Off Frozen" also includes non-credit-hold frozen orders (plain stock/temperature
    // holds) — this is that leftover so the KPI section visibly reconciles to the total pool.
    var otherFrozenCount = cutFrozen.count - creditHold.list.length;
    var otherFrozenVal = cutFrozen.val - creditHold.totalVal;

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();

    // ── SHEET 1: EXECUTIVE SUMMARY ──
    var exec = wb.addWorksheet('Executive Summary');
    exec.columns = [{ width: 42 }, { width: 16 }, { width: 18 }, { width: 12 }, { width: 16 }];
    var t = exec.addRow(['WH BACKLOG RISK COMMAND CENTER — EXECUTIVE SUMMARY']);
    exec.mergeCells('A' + t.number + ':E' + t.number);
    t.font = { bold: true, size: 14, color: { argb: NAVY } };
    exec.addRow(['Generated', new Date().toLocaleString('en-AE'), '', 'Source file', backlogData.fileName || '']);
    exec.addRow([]);

    headerRow(exec, ['METRIC', 'ORDERS', 'VALUE (AED)', '', '']);
    function kpiRow(label, count, val, fillArgb, textArgb) {
      var r = exec.addRow([label, count, Math.round(val)]);
      r.getCell(1).font = { bold: true, color: textArgb ? { argb: textArgb } : undefined };
      r.getCell(2).numFmt = '#,##0';
      r.getCell(3).numFmt = '#,##0';
      if (fillArgb) { [1, 2, 3].forEach(function(ci) { r.getCell(ci).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } }; }); }
      return r;
    }
    kpiRow('Total Pool Orders', s.totalOrders || 0, s.totalVal || 0, 'FFFCF3CF', 'FF7D6608');
    kpiRow('WH Backlog Pure Risk (Before Cut-Off)', whBL.count, whBL.val, 'FFF8D7DA', 'FF842029');
    kpiRow('Advance Orders (RSD > 30 Days)', advOrders.count, advOrders.val);
    kpiRow('Credit Hold (Auto-Booked After Cutoff \u2192 Late Release)', creditHold.list.length, creditHold.totalVal, 'FFFCF3CF', 'FF7D6608');
    kpiRow('Cut-Off Frozen \u2014 Other (Non Credit-Hold: stock/temp hold)', otherFrozenCount, otherFrozenVal);
    kpiRow('Unplanned Frozen', unplanFrozen.count, unplanFrozen.val);
    var checkRow = kpiRow('= Reconciles to Total Pool Orders \u2713', whBL.count + advOrders.count + creditHold.list.length + otherFrozenCount + unplanFrozen.count, whBL.val + advOrders.val + creditHold.totalVal + otherFrozenVal + unplanFrozen.val);
    checkRow.font = { italic: true, color: { argb: 'FF8FA0B0' } };
    kpiRow('Action Required Today', whBL.count, whBL.val, 'FFF8D7DA', 'FF842029');
    exec.addRow([]);

    headerRow(exec, ['ORG', 'ORDERS', 'VALUE (AED)', 'SHARE %', 'RISK']);
    var orgKeys = Object.keys(s.orgMap || {}).sort(function(a, b) { return (s.orgMap[b].val) - (s.orgMap[a].val); });
    var grandVal = s.totalVal || 0;
    orgKeys.forEach(function(o) {
      var v = s.orgMap[o];
      var share = grandVal ? Math.round(v.val / grandVal * 1000) / 10 : 0;
      var rLabel = RISK_LABELS[o] || 'Medium';
      var r = exec.addRow([o, v.count, Math.round(v.val), share, rLabel]);
      r.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RISK_FILL[rLabel] } };
      r.getCell(5).font = { color: { argb: RISK_FONT[rLabel] }, bold: true };
    });
    var totRow = exec.addRow(['TOTAL', s.totalOrders || 0, Math.round(s.totalVal || 0), 100, '']);
    totRow.font = { bold: true };
    exec.addRow([]);

    headerRow(exec, ['CATEGORY', 'ORDERS', 'VALUE (AED)']);
    ['Before Cut-Off', 'RSD >30 Advance', 'Cut-Off Frozen', 'Unplanned Frozen'].forEach(function(catName) {
      var c = cm[catName] || { count: 0, val: 0 };
      exec.addRow([catName, c.count, Math.round(c.val)]);
    });

    // ── CATEGORY SHEETS ──
    function buildCategorySheet(name, data, accentArgb) {
      var ws = wb.addWorksheet(name);
      ws.columns = [{ width: 10 }, { width: 28 }, { width: 16 }, { width: 12 }, { width: 14 }, { width: 24 }];
      var title = ws.addRow([name.toUpperCase() + ' — ORG WISE BREAKDOWN']);
      ws.mergeCells('A' + title.number + ':F' + title.number);
      title.font = { bold: true, size: 13, color: { argb: NAVY } };
      ws.addRow(['Generated', new Date().toLocaleString('en-AE')]);
      ws.addRow([]);
      headerRow(ws, ['ORG', 'ORDERS', 'VALUE (AED)', 'SHARE %']);
      var orgKeys2 = Object.keys(data.orgMap).sort(function(a, b) { return data.orgMap[b].val - data.orgMap[a].val; });
      orgKeys2.forEach(function(o) {
        var v = data.orgMap[o];
        var share = data.totalVal ? Math.round(v.val / data.totalVal * 1000) / 10 : 0;
        var r = ws.addRow([o, v.count, Math.round(v.val), share]);
        r.getCell(1).font = { bold: true, color: { argb: accentArgb } };
      });
      var trow = ws.addRow(['TOTAL', data.list.length, Math.round(data.totalVal), 100]);
      trow.font = { bold: true };
      trow.eachCell(function(c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } }; });
      ws.addRow([]);
      headerRow(ws, ['ORG', 'CUSTOMER', 'SHIPMENT', 'RSD', 'VALUE (AED)', 'STATUS']);
      data.list.forEach(function(r, idx) {
        var row = ws.addRow([r.org || '', r.cust || '', r.ship || '', r.rsd || '', Math.round(r.val || 0), r.rawStatus || r.cat || '']);
        if (idx % 2 === 1) row.eachCell(function(c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F7F7' } }; });
      });
    }

    buildCategorySheet('Advance Orders', advance, BLUE);
    buildCategorySheet('Credit Hold', creditHold, GOLD);
    buildCategorySheet('WH Backlog', whBacklogBucket, RED);

    var buf = await wb.xlsx.writeBuffer();
    var stamp = new Date().toISOString().slice(0, 10);
    res.setHeader('Content-Disposition', 'attachment; filename="WH_Backlog_Risk_' + stamp + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error('backlog export error:', e.message);
    res.status(500).json({ error: 'Export failed: ' + e.message });
  }
});

// ── CREDIT HOLD NOT RELEASED TRACKER ──
// Separate daily feed (typically uploaded ~5:30 PM) listing orders still stuck in
// credit hold — not yet released, not yet reached the warehouse. Matched by
// shipment/order number day-over-day against the tracker itself.
app.post('/api/backlog/credit-hold/upload', requireAuth, async function (req, res) {
  try {
    var b = req.body || {};
    var rows = Array.isArray(b.rows) ? b.rows : [];
    if (!rows.length) return res.status(400).json({ error: 'No rows found in the file.' });
    var uploadedBy = b.uploadedBy || (req.user && req.user.username) || 'Unknown';
    var fileName = b.fileName || 'credit_hold.xlsx';
    var totalCount = rows.length;
    var totalVal = rows.reduce(function (s, r) { return s + (+r.val || 0); }, 0);
    var dateKey = dubaiDateKey();

    await pool.query(`
      INSERT INTO credit_hold_tracking (date_key, uploaded_by, file_name, total_count, total_val, orders, captured_at, updated_at)
      VALUES ($1,$2,$3,$4,$5,$6,NOW(),NOW())
      ON CONFLICT (date_key) DO UPDATE SET
        uploaded_by=$2, file_name=$3, total_count=$4, total_val=$5, orders=$6, captured_at=NOW(), updated_at=NOW()
    `, [dateKey, uploadedBy, fileName, totalCount, totalVal, JSON.stringify(rows)]);

    auditLog(null, uploadedBy, 'UPLOAD', 'Credit Hold Not Released: ' + fileName + ' \u2014 ' + totalCount + ' orders', req.headers['x-forwarded-for'] || req.ip || '');
    res.json({ success: true, date_key: dateKey, total_count: totalCount, total_val: totalVal });
  } catch (e) {
    console.error('credit-hold upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/backlog/credit-hold/status', requireAuth, async function (req, res) {
  try {
    var r = await pool.query('SELECT date_key, captured_at, uploaded_by, file_name, total_count, total_val, orders FROM credit_hold_tracking ORDER BY date_key DESC LIMIT 2');
    if (!r.rows.length) return res.json({ available: false });
    var latest = r.rows[0];
    var previous = r.rows[1] || null;

    function keyOf(o) { return (o.shipment || o.ship || o.invoice || o.order || '') + '|' + (o.org || ''); }
    var latestOrders = latest.orders || [];
    var prevOrders = previous ? (previous.orders || []) : [];
    var latestKeys = {}; latestOrders.forEach(function (o) { latestKeys[keyOf(o)] = o; });
    var prevKeys = {}; prevOrders.forEach(function (o) { prevKeys[keyOf(o)] = o; });

    var released = [], stillOnHold = [], newHolds = [];
    prevOrders.forEach(function (o) { if (!latestKeys[keyOf(o)]) released.push(o); });
    latestOrders.forEach(function (o) {
      if (prevKeys[keyOf(o)]) stillOnHold.push(o); else newHolds.push(o);
    });
    function sumVal(list) { return list.reduce(function (s, o) { return s + (+o.val || 0); }, 0); }

    var orgBreakdown = {};
    latestOrders.forEach(function (o) {
      var org = o.org || 'Unknown';
      if (!orgBreakdown[org]) orgBreakdown[org] = { count: 0, val: 0 };
      orgBreakdown[org].count++; orgBreakdown[org].val += (+o.val || 0);
    });
    var orgList = Object.keys(orgBreakdown).sort(function (a, b) { return orgBreakdown[b].val - orgBreakdown[a].val; }).map(function (o) { return { org: o, count: orgBreakdown[o].count, val: orgBreakdown[o].val }; });

    res.json({
      available: true,
      latest: { date_key: latest.date_key, captured_at: latest.captured_at, uploaded_by: latest.uploaded_by, total_count: +latest.total_count, total_val: +latest.total_val },
      previous: previous ? { date_key: previous.date_key, captured_at: previous.captured_at } : null,
      org_breakdown: orgList,
      released: { count: released.length, val: sumVal(released) },
      still_on_hold: { count: stillOnHold.length, val: sumVal(stillOnHold) },
      new_holds: { count: newHolds.length, val: sumVal(newHolds) }
    });
  } catch (e) {
    console.error('credit-hold status error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Full-color export: org-wise Summary sheet (with released/still-on-hold/new-holds
// context) + a Raw Data sheet listing every order in the latest snapshot.
app.get('/api/backlog/credit-hold/export', requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var r = await pool.query('SELECT date_key, captured_at, uploaded_by, file_name, total_count, total_val, orders FROM credit_hold_tracking ORDER BY date_key DESC LIMIT 2');
    if (!r.rows.length) return res.status(400).json({ error: 'No credit-hold file uploaded yet.' });
    var latest = r.rows[0];
    var previous = r.rows[1] || null;
    var latestOrders = latest.orders || [];
    var prevOrders = previous ? (previous.orders || []) : [];

    function keyOf(o) { return (o.shipment || o.ship || o.invoice || o.order || '') + '|' + (o.org || ''); }
    var latestKeys = {}; latestOrders.forEach(function (o) { latestKeys[keyOf(o)] = o; });
    var prevKeys = {}; prevOrders.forEach(function (o) { prevKeys[keyOf(o)] = o; });
    var released = [], stillOnHold = [], newHolds = [];
    prevOrders.forEach(function (o) { if (!latestKeys[keyOf(o)]) released.push(o); });
    latestOrders.forEach(function (o) { if (prevKeys[keyOf(o)]) stillOnHold.push(o); else newHolds.push(o); });
    function sumVal(list) { return list.reduce(function (s, o) { return s + (+o.val || 0); }, 0); }

    var orgBreakdown = {};
    latestOrders.forEach(function (o) {
      var org = o.org || 'Unknown';
      if (!orgBreakdown[org]) orgBreakdown[org] = { count: 0, val: 0 };
      orgBreakdown[org].count++; orgBreakdown[org].val += (+o.val || 0);
    });
    var orgKeys = Object.keys(orgBreakdown).sort(function (a, b) { return orgBreakdown[b].val - orgBreakdown[a].val; });
    var grandVal = latestOrders.reduce(function (s, o) { return s + (+o.val || 0); }, 0);

    var NAVY = 'FF1B2338', GOLD = 'FFC9A84C', RED = 'FFE84B4B', GREEN = 'FF4ECB8D', BLUE = 'FF5B8DEE';
    function headerRow(ws, cells, fillArgb) {
      var row = ws.addRow(cells);
      row.eachCell(function (cell) {
        cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb || NAVY } };
      });
      return row;
    }

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();

    // ── SUMMARY SHEET ──
    var sm = wb.addWorksheet('Summary');
    sm.columns = [{ width: 30 }, { width: 14 }, { width: 18 }];
    var t = sm.addRow(['CREDIT HOLD NOT RELEASED \u2014 SUMMARY']);
    sm.mergeCells('A' + t.number + ':C' + t.number);
    t.font = { bold: true, size: 14, color: { argb: NAVY } };
    sm.addRow(['Snapshot date', latest.date_key]);
    sm.addRow(['Captured at', new Date(latest.captured_at).toLocaleString('en-AE')]);
    sm.addRow(['Uploaded by', latest.uploaded_by || 'Unknown']);
    sm.addRow([]);

    headerRow(sm, ['METRIC', 'ORDERS', 'VALUE (AED)']);
    function kpiRow(label, count, val, fillArgb) {
      var row = sm.addRow([label, count, Math.round(val)]);
      row.getCell(1).font = { bold: true };
      if (fillArgb) [1, 2, 3].forEach(function (ci) { row.getCell(ci).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillArgb } }; });
      return row;
    }
    kpiRow('Current Credit Hold (Not Released)', latest.total_count, latest.total_val, 'FFFCF3CF');
    if (previous) {
      kpiRow('Released Since Last Update', released.length, sumVal(released), 'FFD4EDDA');
      kpiRow('Still On Hold', stillOnHold.length, sumVal(stillOnHold), 'FFF8D7DA');
      kpiRow('New Holds Since Last Update', newHolds.length, sumVal(newHolds), 'FFD6E4FA');
    }
    sm.addRow([]);

    headerRow(sm, ['ORG', 'ORDERS', 'VALUE (AED)']);
    orgKeys.forEach(function (o) {
      sm.addRow([o, orgBreakdown[o].count, Math.round(orgBreakdown[o].val)]);
    });
    var totRow = sm.addRow(['TOTAL', latestOrders.length, Math.round(grandVal)]);
    totRow.font = { bold: true };
    totRow.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } }; });

    // ── RAW DATA SHEET ──
    var rd = wb.addWorksheet('Raw Data');
    rd.columns = [{ width: 10 }, { width: 32 }, { width: 20 }, { width: 16 }];
    headerRow(rd, ['ORG', 'CUSTOMER', 'SHIPMENT / ORDER #', 'VALUE (AED)'], NAVY);
    latestOrders.forEach(function (o, idx) {
      var row = rd.addRow([o.org || '', o.cust || '', o.shipment || '', Math.round(+o.val || 0)]);
      if (idx % 2 === 1) row.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF7F7F7' } }; });
    });

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Credit_Hold_Not_Released_' + latest.date_key + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error('credit-hold export error:', e.message);
    res.status(500).json({ error: 'Export failed: ' + e.message });
  }
});

app.post('/api/backlog/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    // ── Sub-admin guardrail: 10-hour cooldown only. WH Backlog is a live
    // "current state" snapshot, not per-date historical data, so there's no
    // "existing date" to protect the way Rejection/Dispatch have — the
    // cooldown is what stops rapid-fire re-uploads here. Superadmin exempt. ──
    if (req.user.role === 'subadmin') {
      try {
        var blCooldown = await checkSubadminCooldown(req.user.uid, 'backlog');
        if (blCooldown) {
          return res.status(429).json({
            error: '\u23f3 Sub-admins can upload twice every 10 hours \u2014 you have used both uploads, you cannot re-upload for ' + blCooldown.hoursLeft + 'h ' + blCooldown.minsLeft + 'm.',
            cooldownActive: true, hoursLeft: blCooldown.hoursLeft, minsLeft: blCooldown.minsLeft
          });
        }
      } catch (e) { console.error('Backlog: cooldown check failed:', e.message); }
    }
    var summary = (typeof req.body.summary === 'object') ? req.body.summary : JSON.parse(req.body.summary || '{}');
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var fileName = req.file.originalname || 'backlog.xlsx';
    var totalOrders = parseInt(req.body.totalOrders) || 0;
    backlogData = { uploadedAt: new Date().toISOString(), uploadedBy: uploadedBy, fileName: fileName, totalOrders: totalOrders, summary: summary };
    var dbOk = await dbSaveBacklog(uploadedBy, fileName, totalOrders, summary);
    saveJSON(BACKLOG_FILE, backlogData);
    captureBacklogSnapshot(); // refresh today's daily-tracking row (fire-and-forget)
    auditLog(null, uploadedBy, 'UPLOAD', 'WH Backlog: ' + fileName + ' \u2014 ' + totalOrders + ' orders', req.headers['x-forwarded-for'] || req.ip || '');
    console.log('Backlog saved:', totalOrders, 'orders', dbOk ? '(DB+file)' : '(file only)');
    if (req.user.role === 'subadmin') {
      try { await recordSubadminUpload(req.user.uid, 'backlog'); }
      catch (e) { console.error('Backlog: could not record upload cooldown:', e.message); }
    }
    res.json({ success: true });
  } catch(e) {
    console.error('Backlog upload error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.delete('/api/backlog/clear', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM backlog_data');
    backlogData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

var BACKLOG_SUMMARY_VERSION = 'v3'; // Increment when shortCat logic changes
app.get('/api/backlog/status', function(req, res) {
  if (!backlogData) return res.json({ hasData: false });
  // If summary version doesn't match, force re-upload
  if (!backlogData.summary || backlogData.summary.version !== BACKLOG_SUMMARY_VERSION) {
    return res.json({ hasData: false, reason: 'version_mismatch' });
  }
  res.json({ hasData: true, uploadedAt: backlogData.uploadedAt, uploadedBy: backlogData.uploadedBy, fileName: backlogData.fileName, totalOrders: backlogData.totalOrders, summary: backlogData.summary });
});


app.get('/health', function(req, res) {
  res.json({ status: 'ok', time: new Date().toISOString(), db: !!process.env.DATABASE_URL });
});

//  HELPERS 
function toStr(v) { return String(v == null ? '' : v).trim(); }

// ── Shared sub-admin upload guardrails (used by Rejection, Dispatch, WH Backlog
// uploads). Superadmin is NEVER subject to any of this — always call these only
// after checking req.user.role === 'subadmin'. ──
async function checkSubadminCooldown(userId, endpoint) {
  // Allows up to 2 uploads per rolling 10-hour window (window starts at the
  // first upload). Returns null if clear to upload, or {hoursLeft, minsLeft}
  // once both uploads in the window are used, until 10 hours have passed
  // since the window started.
  // A superadmin-granted exemption (upload_exempt_until in the future) bypasses this
  // entirely, regardless of endpoint — used for someone catching up on several missing
  // days at once, who'd otherwise be blocked after 2 uploads.
  var userR = await pool.query(`SELECT upload_exempt_until FROM users WHERE id=$1`, [userId]);
  if (userR.rows[0] && userR.rows[0].upload_exempt_until && new Date(userR.rows[0].upload_exempt_until) > new Date()) {
    return null;
  }
  var r = await pool.query(`SELECT window_started_at, upload_count FROM upload_cooldowns WHERE user_id=$1 AND endpoint=$2`, [userId, endpoint]);
  if (!r.rows[0]) return null;
  var hoursSince = (Date.now() - new Date(r.rows[0].window_started_at).getTime()) / 3600000;
  if (hoursSince >= 10) return null; // window expired — fresh allowance of 2
  if ((r.rows[0].upload_count || 0) < 2) return null; // still have an upload left in this window
  var minsLeft = Math.ceil((10 - hoursSince) * 60);
  return { hoursLeft: Math.floor(minsLeft / 60), minsLeft: minsLeft % 60 };
}
async function recordSubadminUpload(userId, endpoint) {
  var r = await pool.query(`SELECT window_started_at, upload_count FROM upload_cooldowns WHERE user_id=$1 AND endpoint=$2`, [userId, endpoint]);
  if (!r.rows[0]) {
    await pool.query(
      `INSERT INTO upload_cooldowns (user_id, endpoint, last_upload_at, window_started_at, upload_count) VALUES ($1,$2,NOW(),NOW(),1)`,
      [userId, endpoint]
    );
    return;
  }
  var hoursSince = (Date.now() - new Date(r.rows[0].window_started_at).getTime()) / 3600000;
  if (hoursSince >= 10) {
    // Previous window expired — this upload starts a brand new window.
    await pool.query(
      `UPDATE upload_cooldowns SET last_upload_at=NOW(), window_started_at=NOW(), upload_count=1 WHERE user_id=$1 AND endpoint=$2`,
      [userId, endpoint]
    );
  } else {
    await pool.query(
      `UPDATE upload_cooldowns SET last_upload_at=NOW(), upload_count=upload_count+1 WHERE user_id=$1 AND endpoint=$2`,
      [userId, endpoint]
    );
  }
}
// Stores a blocked upload as a pending approval request for the super admin.
// fileBuffer/fileName let the super admin's later "Approve" click re-run the
// exact same file through the exact same upload logic — nothing is re-typed.
async function submitUploadApproval(opts) {
  await pool.query(
    `INSERT INTO upload_approval_requests (user_id, username, endpoint, file_name, file_data, blocked_dates, reason, meta)
     VALUES ($1,$2,$3,$4,$5,$6,$7,$8)`,
    [opts.userId, opts.username, opts.endpoint, opts.fileName, opts.fileBuffer,
     JSON.stringify(opts.blockedDates || []), opts.reason, JSON.stringify(opts.meta || {})]
  );
}

// Some dispatch sources (confirmed via Salon's file: TASK_ID like "4240053DGC260626SO")
// have no separate ORG column at all — the business-unit code is only ever embedded
// inside the order code/TASK ID itself. Falling back to blank in that case wrongly
// excludes otherwise-legitimate DCV/DCF/DGC/DGS/DSN orders from every org-scoped count
// (App Captured, App Orders, the reconciliation diff). This reads the embedded code the
// same way a human would glancing at the code, and only trusts it if it's one of the
// known business-unit codes — anything else stays blank rather than guessing.
var KNOWN_ORG_CODES = { DCV: true, DCF: true, DGC: true, DGS: true, DSN: true, DPS: true, DPB: true };
function deriveOrgFromCode(code) {
  var m = /^\d+([A-Z]{2,4})\d{6}/.exec(toStr(code).toUpperCase());
  return (m && KNOWN_ORG_CODES[m[1]]) ? m[1] : '';
}

// Same DD/MM/YYYY-safe date parser used in the pallet module — JS's bare
// `new Date(string)` assumes MM/DD/YYYY and silently mis-parses UAE-format
// text dates like "14/07/2026". Needed here too since transport-team files
// mix real Excel date cells with text-formatted ones.
function toDateStrGeneric(val) {
  if (val === undefined || val === null || val === '') return null;
  if (val instanceof Date) { return isNaN(val.getTime()) ? null : val.toISOString().slice(0, 10); }
  if (typeof val === 'number') {
    var d = new Date(Math.round((val - 25569) * 86400 * 1000));
    return isNaN(d.getTime()) ? null : d.toISOString().slice(0, 10);
  }
  var str = String(val).trim();
  var dmy = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (dmy) {
    var day = +dmy[1], month = +dmy[2], year = +dmy[3];
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      var dd = new Date(Date.UTC(year, month - 1, day));
      if (!isNaN(dd.getTime())) return dd.toISOString().slice(0, 10);
    }
  }
  var parsed = new Date(str);
  return isNaN(parsed.getTime()) ? null : parsed.toISOString().slice(0, 10);
}

function normaliseType(raw) {
  var t = toStr(raw).toUpperCase().replace(/\s+/g, ' ').trim();
  if (t === 'FOOD' || t.startsWith('FOOD')) return 'food';
  if (t.includes('NON-FOOD') || t.includes('NON FOOD')) return 'nonfood';
  // Compact (letters/digits only) comparison so stray spaces/punctuation in the
  // source file — "Shark Ninja", "GSEB " — still match; a literal brand name landing
  // in the TYPE column is itself Non-Food, so it's mapped rather than left as raw text.
  var tCompact = t.replace(/[^A-Z0-9]/g, '');
  if (tCompact === 'GSEB' || tCompact === 'SHARKNINJA') return 'nonfood';
  if (t === '3PL' || t === '3 PL' || t === 'HCP') return '3pl';
  if (t === 'VAN') return 'van';
  return t.toLowerCase();
}

function normaliseCity(raw) {
  var c = toStr(raw).toLowerCase();
  if (c.includes('abu dhabi')) return 'Abu Dhabi';
  if (c.includes('hatta')) return 'Hatta';
  if (c.includes('dubai')) return 'Dubai';
  if (c.includes('sharjah')) return 'Sharjah';
  if (c.includes('ajman')) return 'Ajman';
  if (c.includes('fujairah')) return 'Fujairah';
  if (c.includes('al ain') || c.includes('al-ain')) return 'Al Ain';
  if (c.includes('ras al') || c === 'rak') return 'Ras Al Khaimah';
  if (c.includes('umm')) return 'Umm Al Quwain';
  var s = toStr(raw);
  return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
}

// The CITY column is sometimes wrong (e.g. an internal transfer address that literally
// contains "Sharjah" in the text gets logged under CITY="Dubai"). Scan the full address
// text for a city name and prefer that over a mismatched CITY column value.
function detectCityFromAddress(addressText, cityColumnValue) {
  var fallback = normaliseCity(cityColumnValue);
  var addr = toStr(addressText).toLowerCase();
  if (!addr) return fallback;
  var found = null;
  if (addr.includes('abu dhabi')) found = 'Abu Dhabi';
  else if (addr.includes('hatta')) found = 'Hatta';
  else if (addr.includes('sharjah')) found = 'Sharjah';
  else if (addr.includes('ajman')) found = 'Ajman';
  else if (addr.includes('fujairah')) found = 'Fujairah';
  else if (addr.includes('al ain') || addr.includes('al-ain')) found = 'Al Ain';
  else if (addr.includes('ras al khaimah') || addr.includes('rak,') || addr.endsWith('rak')) found = 'Ras Al Khaimah';
  else if (addr.includes('umm al quwain')) found = 'Umm Al Quwain';
  else if (addr.includes('dubai')) found = 'Dubai';
  // If the address text clearly names a different city than the CITY column, trust the address text.
  return found || fallback;
}

// Transport team's FY26 rate card (AED per vehicle, per day/trip) — provided directly by
// transport, not estimated. Matches on distinctive keywords so variations in how the truck-type
// column gets typed ("Ambient-Multi", "AMBIENT MULTI", etc.) still resolve to the right rate.
var TRUCK_RATE_CARD = [
  { keywords: ['FROZEN', 'MULTI'], label: 'Frozen - Multi', rate: 120 },
  { keywords: ['FROZEN', '4 TON'], label: 'Frozen - Bulk 4 Ton', rate: 850 },
  { keywords: ['FROZEN', '10 TON'], label: 'Frozen - Bulk 10 Ton', rate: 1350 },
  { keywords: ['AMBIENT', 'MULTI'], label: 'Ambient - Multi', rate: 104 },
  { keywords: ['AMBIENT', '4 TON'], label: 'Ambient - Bulk 4 Ton', rate: 750 },
  { keywords: ['AMBIENT', '10 TON'], label: 'Ambient - Bulk 10 Ton', rate: 950 },
  { keywords: ['AMBIENT', '40'], label: 'Ambient - Bulk 40 FT', rate: 1200 },
  { keywords: ['E-COMMERCE'], label: 'E-commerce', rate: 20 },
  { keywords: ['ECOMMERCE'], label: 'E-commerce', rate: 20 },
  { keywords: ['EXCLUSIVE', '1 TON'], label: 'Exclusive 1 Ton', rate: 550 },
  { keywords: ['EXCLUSIVE', '4 TON'], label: 'Exclusive 4 Ton', rate: 750 }
];
function matchTruckRate(rawTruckType) {
  var u = toStr(rawTruckType).toUpperCase();
  if (!u) return null;
  for (var i = 0; i < TRUCK_RATE_CARD.length; i++) {
    var entry = TRUCK_RATE_CARD[i];
    var allMatch = entry.keywords.every(function(kw) { return u.indexOf(kw) !== -1; });
    if (allMatch) return entry;
  }
  return null;
}

// ── Vehicle Master fallback: when a drop has no truck-type text but does have a vehicle
// plate number, look up that vehicle's tonnage + Chiller/Frozen/Ambient from the uploaded
// Vehicle Master, combine with that vehicle's drop-count today (1 drop = Bulk, 2+ = Multi,
// per transport team's rule), and reuse the same rate card via a synthetic keyword string.
function normalizeVehicleNoForLookup(raw) {
  return String(raw || '').toUpperCase().replace(/\s+/g, '');
}
function vehicleTempBucket(vehicleTypeRaw) {
  var u = toStr(vehicleTypeRaw).toUpperCase();
  if (!u) return null;
  if (u.indexOf('FROZEN') !== -1 || u.indexOf('FREEZER') !== -1) return 'FROZEN';
  if (u.indexOf('CHILL') !== -1 || u.indexOf('AMBIENT') !== -1) return 'AMBIENT'; // Chiller priced as Ambient per transport team
  return null; // Dry / Open Pick-up / Car / Bus / Other — no rate applies
}
function vehicleTonnageBucket(vehTypeRaw) {
  var u = toStr(vehTypeRaw).toUpperCase().replace(/\s+/g, '');
  if (u.indexOf('4.2T') !== -1 || u === '4T') return '4 TON';
  if (u.indexOf('10T') !== -1) return '10 TON';
  return null; // 3T / 12T — genuinely no Bulk tier for these on the FY26 rate card
}
function lookupVehicleMasterByAnyId(vehicleId) {
  var norm = normalizeVehicleNoForLookup(vehicleId);
  var v = VEHICLE_MASTER_MAP[norm];
  if (!v) {
    var digitGroups = toStr(vehicleId).match(/\d+/g);
    if (digitGroups && digitGroups.length) {
      v = VEHICLE_MASTER_MAP['DIGITS:' + digitGroups[digitGroups.length - 1]];
    }
  }
  return v || null;
}
function matchRateViaVehicleMaster(vehicleId, dropCountForVehicle) {
  var v = lookupVehicleMasterByAnyId(vehicleId);
  if (!v) return null;
  var tempBucket = vehicleTempBucket(v.vehicle_type_raw);
  if (!tempBucket) return null;
  var isOneTon = toStr(v.veh_type).toUpperCase().replace(/\s+/g, '') === '1T';
  // 1 Ton has no dedicated Multi/Bulk-4-Ton style tier on the FY26 rate card — "Exclusive 1 Ton"
  // (550 AED) is the only rate that exists for this tonnage at all, so it's used regardless of
  // drop count for 1-ton vehicles specifically.
  if (isOneTon) return matchTruckRate('EXCLUSIVE 1 TON');
  var tier = dropCountForVehicle >= 2 ? 'MULTI' : vehicleTonnageBucket(v.veh_type);
  if (!tier) return null;
  return matchTruckRate(tempBucket + ' ' + tier);
}

function extractDriverName(contact) {
  var s = toStr(contact);
  if (!s) return '';
  // Transport staff sometimes already label these explicitly, e.g. "Hired Driver-Ayaz 566" —
  // that's a complete, meaningful identifier already; keep it whole instead of running it
  // through the phone-number-stripping logic below, which would otherwise discard the name
  // and collapse every distinct hired driver into one generic "Hired Driver" entry.
  if (/^hired driver/i.test(s)) return s.trim();
  if (/^[A-Za-z][A-Za-z\s]{2,}$/.test(s)) return s.trim();
  var m = s.match(/^([A-Za-z][A-Za-z\s]{2,29})(?:\s*[-+\d])/);
  if (m) return m[1].trim();
  var parts = s.split(/[-+\d]/);
  var name = (parts[0] || '').trim();
  if (name.length > 2) return name;
  // No name at all in the source data — just a bare phone number/ID. Label it clearly
  // instead of silently displaying the number, so it reads as "name missing", not a bug.
  if (/^\d+$/.test(s)) return 'Hired Driver (ID: ' + s + ')';
  return s.trim();
}

function stripBranch(name) {
  var base = toStr(name);
  var kws = [',Branch', ', Branch', ',Br.', ', Br.', ' -Branch', ',CPD', ' CPD', '- Branch', '-Branch'];
  for (var i = 0; i < kws.length; i++) {
    var idx = base.toLowerCase().indexOf(kws[i].toLowerCase());
    if (idx > 3) { base = base.substring(0, idx).trim(); break; }
  }
  return base.replace(/,\s*(LLC|L\.L\.C|llc).*$/i, '').trim();
}

function findDataSheet(wb) {
  var bestSheet = wb.SheetNames[0], bestRows = 0;
  for (var i = 0; i < wb.SheetNames.length; i++) {
    var name = wb.SheetNames[i];
    var ws = wb.Sheets[name];
    if (!ws['!ref']) continue;
    var range = XLSX.utils.decode_range(ws['!ref']);
    var r = range.e.r - range.s.r;
    if (r > bestRows) { bestRows = r; bestSheet = name; }
  }
  console.log('Using sheet:', bestSheet, 'rows:', bestRows);
  return bestSheet;
}

//  DISPATCH PARSER 
function parseDispatch(buffer) {
  var wb = XLSX.read(buffer, { type: 'buffer', dense: true, cellDates: false, cellNF: false, cellHTML: false, cellFormula: false });
  var sheetName = findDataSheet(wb);
  var ws = wb.Sheets[sheetName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  return parseDispatchFromRows(rows);
}

// Same aggregation this always ran, just usable on rows straight from
// storage (e.g. merged Main + Salon rows for one date) as well as a
// freshly-read file, without duplicating the ~450 lines of logic below.
function parseDispatchFromRows(rows) {
  if (!rows.length) return null;

  // Sources whose orders/sales count normally, but whose routes, vehicles, and drivers
  // never feed transport-cost or driver-cost tracking — set explicitly here rather than
  // relying on their file format happening to lack a Route column, so this stays true
  // even if that file's structure changes later. Currently: Salon (mixed/DIP-warehouse
  // deliveries with an unconfirmed transport billing model).
  var COST_TRACKING_EXCLUDED_SOURCES = ['Salon'];
  function isCostExcludedRow(row) {
    return row.__dispatch_source && COST_TRACKING_EXCLUDED_SOURCES.indexOf(row.__dispatch_source) !== -1;
  }

  function findCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
    }) || null;
  }
  // Exact (not fuzzy-substring) header match, kept as a fallback only. Confirmed directly
  // against real files: "ORDER CODE" (e.g. "4256142DCV130726SO") stays IDENTICAL across
  // days for the same re-delivered order, so it's the primary identifier (see findCol call
  // below). The plain "ORDER" column isn't always present in every day's export, so it's
  // only used when "ORDER CODE" is missing entirely.
  function findExactCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.trim().toUpperCase() === n.toUpperCase(); });
    }) || null;
  }

  var C = {
    route:    findCol('ROUTE'),
    city:     findCol('CITY', 'AREA'),
    customer: findCol('CUSTOMER NAME', 'CUSTOMER'),
    amount:   findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE'),
    driver:   findCol('DRIVER CONTACT DETAILS', 'DRIVERS NAME', 'DRIVER NAME', 'DRIVER CONTACT', 'DRIVER_CONTACT', 'DRIVER_ID'),
    location: findCol('LOCATION_ID', 'LOCATION'),
    address:  findCol('CUSTOMER ADDRESS', 'ADDRESS'),
    keep:     findCol('KEEP TOGETHER', 'KEEP_TOGETHER', 'KEEPTOGETHER', 'KEEP'),
    type:     findCol('TYPE'),
    temperature: findCol('TEMPERATURE', 'TEMP'),
    vehicleId: findCol('VEHICLE_ID', 'VEHICLE ID', 'VEHICLE'),
    truckType: findCol('TRUCK TYPE', 'TRUCK_TYPE', 'VEHICLE TYPE', 'VEHICLE_TYPE', 'DROP TYPE', 'DROP_TYPE'),
    orderCode: findCol('ORDER CODE', 'ORDER_CODE') || findExactCol('ORDER', 'ORDER NUMBER', 'ORDER_NUMBER') || findCol('TASK_ID', 'TASK ID') || findCol('ORDER '),
    org:      findCol('ORG') || findCol('BU') || findCol('ORGANIZATION') || findCol('ORG-BU')
  };
  console.log('Dispatch cols:', JSON.stringify(C));

  var totalOrders=0, totalValue=0, foodOrders=0, foodValue=0, nonFoodOrders=0, nonFoodValue=0, plOrders=0, vanOrders=0;
  var frozenOrders=0, frozenValue=0, ambientOrders=0, ambientValue=0;
  var cities={}, customers={}, routes={}, driverSet={};
  var dropsByCity = {}; // one increment per unique (route, location) drop — cannot exceed total_drops by construction
  var dropRecords = {}; // route::loc -> { city, truckType, types:{}, vehicleId } — accumulated across ALL rows for that drop, so region/food-type breakdowns are based on complete data, not just whichever row happened to create the drop first
  var orgStats={ DCV:{o:0,v:0}, DCF:{o:0,v:0}, DGC:{o:0,v:0}, DGS:{o:0,v:0}, DSN:{o:0,v:0}, DPS:{o:0,v:0}, DPB:{o:0,v:0}, HCP:{o:0,v:0} };
  var cityTypeCross = {}; // city -> {food, nonfood, pl, van, other}
  var locationVisits = {}; // locationId -> { address, customer, routes:Set(all), ownRoutes:Set(non-3PL) }
  // Transport Cost Reconciliation only ever compares against the 4 business-unit org
  // codes that the transport team's own file could ever contain (DCV, DGC, DGS, DSN).
  // 3PL/HCP drops are fulfilled by a third party, and DCF (food service) self-delivers
  // on its own van — neither ever goes through the shared transport team's truck fleet
  // — so counting either toward "App Captured" would always show as a false variance/
  // "missing from transport file". This is a SEPARATE counter from total_drops below
  // (which intentionally stays all-inclusive for the Executive Summary / cost-estimate
  // features) so nothing else on the dashboard changes.
  var RECON_ELIGIBLE_ORGS = { DCV: true, DGC: true, DGS: true, DSN: true };
  // DCF self-delivers on its own van ONLY in Dubai — DCF orders in other emirates still
  // go through the shared transport team and should be compared normally. So DCF isn't
  // blanket-excluded; it's excluded only when the drop's own city is Dubai.
  function isReconEligibleOrg(orgCode, cityName) {
    if (RECON_ELIGIBLE_ORGS[orgCode]) return true;
    if (orgCode === 'DCF' && String(cityName || '').toUpperCase() !== 'DUBAI') return true;
    return false;
  }
  var reconDropKeysSeen = {}; // route::loc -> true, only for RECON_ELIGIBLE_ORGS rows
  var totalDropsReconcile = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    totalOrders++;
    var amt = parseFloat(row[C.amount]) || 0;
    totalValue += amt;
    var type = normaliseType(C.type ? row[C.type] : '');
    var org = C.org ? toStr(row[C.org]).toUpperCase() : '';
    if (!org && C.orderCode) org = deriveOrgFromCode(row[C.orderCode]);
    var tempForRow = C.temperature ? toStr(row[C.temperature]).toUpperCase() : '';
    var isFrozenRow = tempForRow.indexOf('FROZEN') !== -1;
    // Frozen/Ambient is a DCV-only breakdown — other BUs (Salon/DGC, DSN, etc.) carry
    // their own temperature values in the same column, and mixing them in here made this
    // card show a number that didn't reflect DCV's own frozen volume.
    if (tempForRow && org === 'DCV') {
      if (isFrozenRow) { frozenOrders++; frozenValue += amt; }
      else { ambientOrders++; ambientValue += amt; }
    }
    var rawTruckTypeForRow = C.truckType ? toStr(row[C.truckType]) : '';
    if (type === 'food')   { foodOrders++;    foodValue    += amt; }
    else if (type === 'nonfood') { nonFoodOrders++; nonFoodValue += amt; }
    else if (type === '3pl')     { plOrders++; }
    else if (type === 'van')     { vanOrders++; }
    if (type === '3pl') { orgStats.HCP.o++; orgStats.HCP.v += amt; }
    else if (orgStats[org]) { orgStats[org].o++; orgStats[org].v += amt; }
    else if (org === '3 PL' || org === 'HCP' || org === '3PL') { orgStats.HCP.o++; orgStats.HCP.v += amt; }
    if (C.city && row[C.city]) {
      var city = normaliseCity(row[C.city]);
      if (!cities[city]) cities[city] = { orders:0, value:0 };
      cities[city].orders++; cities[city].value += amt;
      if (!cityTypeCross[city]) cityTypeCross[city] = { food:0, nonfood:0, pl:0, van:0, other:0 };
      if (type === 'food') cityTypeCross[city].food++;
      else if (type === 'nonfood') cityTypeCross[city].nonfood++;
      else if (type === '3pl') cityTypeCross[city].pl++;
      else if (type === 'van') cityTypeCross[city].van++;
      else cityTypeCross[city].other++;
    }
    if (C.customer && row[C.customer]) {
      var cust = toStr(row[C.customer]);
      if (!customers[cust]) customers[cust] = { orders:0, value:0 };
      customers[cust].orders++; customers[cust].value += amt;
    }
    if (C.route && row[C.route] && !isCostExcludedRow(row)) {
      var route = toStr(row[C.route]);
      if (!routes[route]) routes[route] = { locs:{}, drivers:{}, orders:0, value:0, types:{}, vehicleIds:{} };
      var vehicleIdForRoute = C.vehicleId ? toStr(row[C.vehicleId]) : '';
      if (vehicleIdForRoute) routes[route].vehicleIds[vehicleIdForRoute] = (routes[route].vehicleIds[vehicleIdForRoute] || 0) + 1;
      var routeTypeLabel = (type||'other').charAt(0).toUpperCase()+(type||'other').slice(1);
      if (type === 'nonfood') routeTypeLabel = 'Non-Food';
      else if (type === '3pl') routeTypeLabel = '3PL';
      if (tempForRow) routeTypeLabel += ' (' + (isFrozenRow ? 'Frozen' : 'Ambient') + ')';
      routes[route].types[routeTypeLabel] = (routes[route].types[routeTypeLabel] || 0) + 1;
      var rawLoc = C.location ? toStr(row[C.location]) : '';
      // Internal cash-van transfers all drop at a fixed hub per city (not a real unique customer
      // address per transaction) — collapse them to one drop per city instead of counting every
      // internal order's own LOCATION_ID as a separate physical stop.
      var custForInternal = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
      var isInternalVan = (type === 'van') && custForInternal.indexOf('INTERNAL') !== -1;
      var addrTextForCity = C.address ? toStr(row[C.address]) : '';
      var cityForLoc = detectCityFromAddress(addrTextForCity, C.city ? row[C.city] : '');
      var loc = isInternalVan ? ('INTERNAL-HUB::' + (cityForLoc || 'Unknown')) : rawLoc;
      if (loc) {
        if (!routes[route].locs[loc]) {
          // First time this exact (route, location) pair is seen — this is a genuinely new drop.
          dropsByCity[cityForLoc || 'Unknown'] = (dropsByCity[cityForLoc || 'Unknown'] || 0) + 1;
        }
        routes[route].locs[loc] = 1;
        // Reconciliation-scoped count: DCV/DGC/DGS/DSN always eligible; DCF eligible only
        // outside Dubai (see isReconEligibleOrg above). Excludes 3PL/HCP and any other org.
        if (isReconEligibleOrg(org, cityForLoc)) {
          var reconDropKey = route + '::' + loc;
          if (!reconDropKeysSeen[reconDropKey]) {
            reconDropKeysSeen[reconDropKey] = true;
            totalDropsReconcile++;
          }
        }
        // Accumulate this drop's full record across every row that contributes to it, so the
        // region + Food/Non-Food breakdown (computed after the loop) sees ALL types on the
        // drop, not just whichever row happened to create it first.
        var dropKey = route + '::' + loc;
        if (!dropRecords[dropKey]) {
          dropRecords[dropKey] = { city: cityForLoc || 'Unknown', truckType: rawTruckTypeForRow, types: {}, vehicleId: '' };
        }
        if (rawTruckTypeForRow && !dropRecords[dropKey].truckType) dropRecords[dropKey].truckType = rawTruckTypeForRow;
        dropRecords[dropKey].types[type || 'other'] = true;
        var vehicleIdValForDrop = C.vehicleId ? toStr(row[C.vehicleId]) : '';
        if (vehicleIdValForDrop) dropRecords[dropKey].vehicleId = vehicleIdValForDrop;
      }
      routes[route].value += amt;
      routes[route].orders++;
      if (C.driver && row[C.driver]) {
        var drvName = extractDriverName(row[C.driver]);
        if (drvName) routes[route].drivers[drvName] = 1;
      }
      // Track which routes visit each physical location, to catch the same address being
      // driven to twice by two different routes on the same day (double-charged drop).
      // Cash/walk-in orders ("**** Cash **** (Dxb)") share a generic placeholder location,
      // not a real fixed address, so they're flagged and excluded from the repeat-visit REPORT
      // further below (but still counted normally for the own-fleet/3PL drop-cost split).
      // Also track the order TYPE per route-visit — Food and Non-Food often can't share a
      // truck, so a "repeat visit" that's actually Food+Non-Food is a legitimate separate
      // trip, not a duplicate/avoidable one. Frozen vs Ambient of the SAME product type also
      // needs a separate truck for temperature control, so that's folded into the same check.
      var truckType = type + (isFrozenRow ? '-frozen' : (tempForRow ? '-ambient' : ''));
      if (loc) {
        var custNameForCash = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
        if (!locationVisits[loc]) {
          locationVisits[loc] = {
            address: C.address ? toStr(row[C.address]) : '',
            customer: C.customer ? toStr(row[C.customer]) : '',
            isCashOrder: custNameForCash.indexOf('CASH') !== -1,
            isInternalHub: loc.indexOf('INTERNAL-HUB::') === 0,
            routes: {}, ownRoutes: {}, typesByRoute: {}, valueByRoute: {}, orderCountByRoute: {}, ordersByRoute: {}
          };
        }
        locationVisits[loc].routes[route] = 1;
        if (type !== '3pl') locationVisits[loc].ownRoutes[route] = 1;
        if (!locationVisits[loc].typesByRoute[route]) locationVisits[loc].typesByRoute[route] = {};
        locationVisits[loc].typesByRoute[route][truckType || 'other'] = true;
        locationVisits[loc].valueByRoute[route] = (locationVisits[loc].valueByRoute[route] || 0) + amt;
        locationVisits[loc].orderCountByRoute[route] = (locationVisits[loc].orderCountByRoute[route] || 0) + 1;
        if (!locationVisits[loc].ordersByRoute[route]) locationVisits[loc].ordersByRoute[route] = [];
        locationVisits[loc].ordersByRoute[route].push({
          order_code: C.orderCode ? toStr(row[C.orderCode]) : '',
          type: type || 'other',
          temperature: isFrozenRow ? 'Frozen' : (tempForRow ? 'Ambient' : ''),
          value: Math.round(amt)
        });
      }
    }
    if (C.driver && row[C.driver]) {
      var keepVal2 = C.keep ? toStr(row[C.keep]) : (C.location ? toStr(row[C.location]) : '');
      var drv = extractDriverName(row[C.driver]) || toStr(row[C.driver]);
      if (drv) driverSet[drv] = 1;
    }
  }

  console.log('TYPE counts food:'+foodOrders+' nonfood:'+nonFoodOrders+' 3pl:'+plOrders+' van:'+vanOrders);

  var byCity = Object.keys(cities).map(function(c) {
    return { city:c, orders:cities[c].orders, value:Math.round(cities[c].value) };
  }).sort(function(a,b) { return b.orders-a.orders; });

  var baseCust = {};
  Object.keys(customers).forEach(function(name) {
    var base = stripBranch(name);
    if (!baseCust[base]) baseCust[base] = { orders:0, value:0 };
    baseCust[base].orders += customers[name].orders;
    baseCust[base].value  += customers[name].value;
  });
  var topCustomers = Object.keys(baseCust).map(function(name) {
    return { name:name, orders:baseCust[name].orders, value:Math.round(baseCust[name].value) };
  }).sort(function(a,b) { return b.value-a.value; }).slice(0,6);

  var topRoutes = Object.keys(routes).map(function(route) {
    var typeMap = routes[route].types || {};
    var topTypesForRoute = Object.keys(typeMap).sort(function(a,b){ return typeMap[b]-typeMap[a]; });
    var hasFood = topTypesForRoute.some(function(t){ return t.indexOf('Food') === 0; });
    var hasNonFood = topTypesForRoute.some(function(t){ return t.indexOf('Non-Food') === 0; });
    var hasFrozen = topTypesForRoute.some(function(t){ return t.indexOf('Frozen') !== -1; });
    var hasAmbient = topTypesForRoute.some(function(t){ return t.indexOf('Ambient') !== -1; });
    var guessedPartition = (hasFood && hasNonFood) || (hasFrozen && hasAmbient);

    // Prefer the actual Vehicle Master partition flag over the route-content guess, whenever
    // we can identify which vehicle ran this route (majority vehicle ID seen on it).
    var isPartitionVehicle = guessedPartition;
    var partitionSource = 'guessed';
    var vids = Object.keys(routes[route].vehicleIds || {});
    if (vids.length) {
      var majorityVid = vids.sort(function(a,b){ return routes[route].vehicleIds[b] - routes[route].vehicleIds[a]; })[0];
      var vm = lookupVehicleMasterByAnyId(majorityVid);
      if (vm && vm.partition_flag) {
        var pf = toStr(vm.partition_flag).trim().toUpperCase();
        if (pf === 'YES' || pf === 'Y') { isPartitionVehicle = true; partitionSource = 'vehicle_master'; }
        else if (pf === 'NO' || pf === 'N') { isPartitionVehicle = false; partitionSource = 'vehicle_master'; }
      }
    }
    return { route:route, orders:routes[route].orders, drops:Object.keys(routes[route].locs).length, driverCount:Object.keys(routes[route].drivers).length, value:Math.round(routes[route].value), types:topTypesForRoute, isPartitionVehicle:isPartitionVehicle, partitionSource:partitionSource };
  }).sort(function(a,b) { return b.drops-a.drops; });

  // Count actual orders per driver (not route drops)
  var driverOrders = {};
  rows.forEach(function(row) {
    if (isCostExcludedRow(row)) return;
    var drv = C.driver && row[C.driver] ? extractDriverName(row[C.driver]) : '';
    if (!drv) return;
    var amt = C.amount ? parseFloat(row[C.amount]) || 0 : 0;
    var rawLocId = C.location ? toStr(row[C.location]) : '';
    // Same internal-van hub collapse + address-based city correction used for Route Summary,
    // so the driver leaderboard's drop count matches what's shown everywhere else.
    var typeForDrv = C.type ? normaliseType(row[C.type]) : '';
    var custForDrv = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
    var isInternalVanDrv = (typeForDrv === 'van') && custForDrv.indexOf('INTERNAL') !== -1;
    var cityForDrv = detectCityFromAddress(C.address ? toStr(row[C.address]) : '', C.city ? row[C.city] : '');
    var locId = isInternalVanDrv ? ('INTERNAL-HUB::' + (cityForDrv || 'Unknown')) : rawLocId;
    var custNameForDrv = C.customer ? toStr(row[C.customer]) : '';
    var addressForDrv = C.address ? toStr(row[C.address]) : '';
    var orderCodeForDrv = C.orderCode ? toStr(row[C.orderCode]) : '';
    if (!driverOrders[drv]) driverOrders[drv] = {orders:0, drops:{}, value:0, customers:{}, types:{}, orderDetails:[]};
    driverOrders[drv].orders++;
    driverOrders[drv].value += amt;
    if (locId) driverOrders[drv].drops[locId] = 1;
    if (custNameForDrv) driverOrders[drv].customers[custNameForDrv] = (driverOrders[drv].customers[custNameForDrv] || 0) + 1;
    if (typeForDrv) driverOrders[drv].types[typeForDrv] = (driverOrders[drv].types[typeForDrv] || 0) + 1;
    // Order-level record — one per dispatch row — so the driver's own drops can be
    // inspected order by order (which customer, what value, which final drop) rather
    // than only as a rolled-up count. "Final drop" is that order's actual delivery
    // point: the internal-van-hub-collapsed location used everywhere else on this
    // dashboard, falling back to the raw address when no location code exists.
    driverOrders[drv].orderDetails.push({
      customer: custNameForDrv || '(no customer name)',
      value: Math.round(amt),
      drop: locId || addressForDrv || '(no location)',
      address: addressForDrv,
      order_code: orderCodeForDrv,
      type: typeForDrv || ''
    });
  });
  var driverList = Object.keys(driverOrders).map(function(name) {
    var custMap = driverOrders[name].customers || {};
    var typeMap = driverOrders[name].types || {};
    var topCustomers = Object.keys(custMap).sort(function(a,b){ return custMap[b]-custMap[a]; }).slice(0,3);
    var topTypes = Object.keys(typeMap).sort(function(a,b){ return typeMap[b]-typeMap[a]; });
    var drops = Object.keys(driverOrders[name].drops).length;
    var value = Math.round(driverOrders[name].value);
    return { name:name, orders:driverOrders[name].orders, drops:drops, value:value, isHired: /^hired driver/i.test(name), customers:topCustomers, types:topTypes,
      value_per_drop: drops > 0 ? Math.round(value / drops) : 0,
      order_details: driverOrders[name].orderDetails.sort(function(a,b){ return b.value-a.value; }) };
  });
  var topDrivers = driverList.slice().sort(function(a,b) { return b.orders-a.orders; }).slice(0,5);
  // Order-count ranking hides drivers who carry only a few, very high-value deliveries
  // (e.g. a single route to a major supermarket) — surface those separately.
  var topDriversByValue = driverList.slice().sort(function(a,b) { return b.value-a.value; }).slice(0,5);

  // How much of today's dispatch relied on hired/agency drivers (no name in source data)
  // vs named in-house drivers — a bare phone number is the signal of a hired driver.
  var hiredDrivers = driverList.filter(function(d) { return d.isHired; });
  var inhouseDrivers = driverList.filter(function(d) { return !d.isHired; });
  var driverSourceSplit = {
    hired: {
      driver_count: hiredDrivers.length,
      orders: hiredDrivers.reduce(function(s,d){ return s+d.orders; }, 0),
      value: hiredDrivers.reduce(function(s,d){ return s+d.value; }, 0),
      drops: hiredDrivers.reduce(function(s,d){ return s+d.drops; }, 0)
    },
    inhouse: {
      driver_count: inhouseDrivers.length,
      orders: inhouseDrivers.reduce(function(s,d){ return s+d.orders; }, 0),
      value: inhouseDrivers.reduce(function(s,d){ return s+d.value; }, 0),
      drops: inhouseDrivers.reduce(function(s,d){ return s+d.drops; }, 0)
    },
    hired_driver_details: hiredDrivers.sort(function(a,b){ return b.value-a.value; }),
    // In-house driver detail, same shape as hired_driver_details above (incl. per-order
    // customer/value/drop) — stored so the Driver Order Details export/analysis below
    // can cover the whole fleet, not just hired drivers.
    inhouse_driver_details: inhouseDrivers.sort(function(a,b){ return b.value-a.value; })
  };

  // ── Own fleet vs 3PL drop split, and repeat-visit detection ──
  // 3PL orders are fulfilled by a third party (not our own fleet), so they're tracked
  // separately from our own-fleet drop count.
  var ownFleetDrops = 0, plDrops = 0;
  Object.keys(routes).forEach(function(r) {
    Object.keys(routes[r].locs).forEach(function(loc) {
      // A location counts as a "3PL drop" for this route only if EVERY visit to it
      // on this route was 3PL; otherwise it's counted as an own-fleet drop.
      if (locationVisits[loc] && locationVisits[loc].ownRoutes[r]) ownFleetDrops++;
      else plDrops++;
    });
  });

  var HIGH_VALUE_EXCEPTION_THRESHOLD = 100000;

  var repeatLocations = Object.keys(locationVisits)
    .map(function(loc) {
      var lv = locationVisits[loc];
      var ownRouteList = Object.keys(lv.ownRoutes);
      // For each own-fleet route that visited this location, what order type(s), value, order count,
      // and the actual order codes did it carry? (for full traceability back to source rows)
      var routeDetails = ownRouteList.map(function(r) {
        var typesHere = Object.keys(lv.typesByRoute[r] || {});
        // Which driver(s)/truck(s) ran this route — so a duplicate-drop location can name
        // exactly who to consolidate, not just which anonymous "route" visited twice.
        var driversHere = (routes[r] && routes[r].drivers) ? Object.keys(routes[r].drivers) : [];
        return {
          route: r, types: typesHere, drivers: driversHere,
          value: Math.round(lv.valueByRoute[r] || 0),
          order_count: lv.orderCountByRoute[r] || 0,
          orders: lv.ordersByRoute[r] || []
        };
      });
      var totalValue = routeDetails.reduce(function(s, rd) { return s + rd.value; }, 0);
      // All distinct types seen across all routes at this location
      var allTypes = {};
      routeDetails.forEach(function(rd) { rd.types.forEach(function(t) { allTypes[t] = true; }); });
      var distinctTypeCount = Object.keys(allTypes).length;
      // Legitimate split = different routes carried genuinely different order types
      // (e.g. one route Food, another Non-Food) — those must use separate trucks.
      // A real avoidable duplicate = multiple routes carrying the SAME type to the same address.
      var isLegitimateSplit = distinctTypeCount > 1;
      // High-value exception: an "avoidable" duplicate over AED 100k is more likely a genuinely
      // large order that needed splitting across trucks for capacity — flag for manual review
      // rather than assuming it's a routing mistake.
      var isHighValueException = !isLegitimateSplit && totalValue > HIGH_VALUE_EXCEPTION_THRESHOLD;
      return {
        location_id: loc,
        address: lv.address,
        customer: lv.customer,
        isCashOrder: lv.isCashOrder,
        isInternalHub: lv.isInternalHub,
        routes: ownRouteList,
        route_types: routeDetails,
        total_value: totalValue,
        visit_count: ownRouteList.length,
        is_legitimate_split: isLegitimateSplit,
        is_high_value_exception: isHighValueException,
        reason: isLegitimateSplit
          ? 'Different order types (' + Object.keys(allTypes).join(' + ') + ') — separate trucks required'
          : (isHighValueException
              ? 'EXCEPTION: AED ' + totalValue.toLocaleString() + ' — likely a genuinely large order needing capacity split, verify before flagging as routing error'
              : 'Same order type visited ' + ownRouteList.length + 'x — likely avoidable')
      };
    })
    .filter(function(l) { return l.visit_count > 1 && !l.isCashOrder && !l.isInternalHub; })
    .sort(function(a, b) { return (b.total_value - a.total_value) || (b.visit_count - a.visit_count); });

  var repeatLocationAvoidableCount = repeatLocations.filter(function(l) { return !l.is_legitimate_split; }).length;

  var cityTypeCrossOut = {};
  Object.keys(cityTypeCross).forEach(function(c) {
    cityTypeCrossOut[c] = cityTypeCross[c];
  });

  // Food/Non-Food headline totals must always equal the sum of the org codes shown
  // directly underneath them (DCV+DCF for Food, DGC+DGS+DSN for Non-Food) — anything
  // else lets the big number disagree with its own breakdown, which is what was
  // happening on files with no explicit TYPE column (the real dispatch report has
  // none; classification there only ever came from ORG, so the TYPE-based totals
  // were structurally always 0 while the org breakdown was correct).
  if (foodOrders === 0 && (orgStats.DCV.o > 0 || orgStats.DCF.o > 0)) {
    foodOrders = orgStats.DCV.o + orgStats.DCF.o;
    foodValue = orgStats.DCV.v + orgStats.DCF.v;
  }
  if (nonFoodOrders === 0 && (orgStats.DGC.o > 0 || orgStats.DGS.o > 0 || orgStats.DSN.o > 0)) {
    nonFoodOrders = orgStats.DGC.o + orgStats.DGS.o + orgStats.DSN.o;
    nonFoodValue = orgStats.DGC.v + orgStats.DGS.v + orgStats.DSN.v;
  }

  return {
    total_orders: totalOrders, total_value: Math.round(totalValue),
    total_routes: Object.keys(routes).length,
    total_drivers: Object.keys(driverOrders).length || Object.keys(driverSet).length,
    total_drops: Object.keys(routes).reduce(function(s,r){ return s + Object.keys(routes[r].locs).length; }, 0),
    // Reconciliation-scoped drop count — DCV/DCF/DGC/DGS/DSN only, excludes 3PL/HCP and
    // any other org. This is what "App Captured" on the Transport Cost Reconciliation
    // card should use, since transport's own file only ever covers these 5 org codes.
    total_drops_reconcile: totalDropsReconcile,
    food_orders: foodOrders, food_value: Math.round(foodValue),
    non_food_orders: nonFoodOrders, non_food_value: Math.round(nonFoodValue),
    pl_orders: plOrders, van_orders: vanOrders,
    frozen_orders: frozenOrders, frozen_value: Math.round(frozenValue),
    ambient_orders: ambientOrders, ambient_value: Math.round(ambientValue),
    type_breakdown: {
      DCV: { orders:orgStats.DCV.o, value:Math.round(orgStats.DCV.v) },
      DCF: { orders:orgStats.DCF.o, value:Math.round(orgStats.DCF.v) },
      DGC: { orders:orgStats.DGC.o, value:Math.round(orgStats.DGC.v) },
      DGS: { orders:orgStats.DGS.o, value:Math.round(orgStats.DGS.v) },
      DSN: { orders:orgStats.DSN.o, value:Math.round(orgStats.DSN.v) },
      DPS: { orders:orgStats.DPS.o, value:Math.round(orgStats.DPS.v) },
      DPB: { orders:orgStats.DPB.o, value:Math.round(orgStats.DPB.v) },
      HCP: { orders:orgStats.HCP.o, value:Math.round(orgStats.HCP.v) }
    },
    by_city: byCity, top_customers: topCustomers,
    top_drivers: topDrivers, top_drivers_by_value: topDriversByValue, driver_source_split: driverSourceSplit, top_routes: topRoutes,
    city_type_cross: cityTypeCrossOut,
    drops_by_city: dropsByCity,
    truck_cost_estimate: (function(){
      var unmatchedTruckTypes = {};
      var byType = {};       // label -> { rate, drop_count, vehicles:{} }
      var byRegion = {};     // city -> { label -> { rate, drop_count } }
      var byFoodType = {};   // 'Food' | 'Non-Food' | 'Mixed (Partition)' | 'Other' -> { drop_count, estimated_cost }
      var byTempFood = {};   // 'Frozen · Food' | 'Ambient · Non-Food' etc -> { drop_count, estimated_cost }

      // Pre-pass: how many drops does each vehicle make today? Needed to decide Multi vs Bulk
      // for the Vehicle Master fallback (transport team's rule: 1 drop = Bulk, 2+ = Multi).
      var dropCountByVehicle = {};
      Object.keys(dropRecords).forEach(function(key){
        var vid = dropRecords[key].vehicleId;
        if (vid) dropCountByVehicle[vid] = (dropCountByVehicle[vid] || 0) + 1;
      });

      Object.keys(dropRecords).forEach(function(key){
        var d = dropRecords[key];
        var rateEntry = null;
        if (d.truckType) {
          rateEntry = matchTruckRate(d.truckType);
        } else if (d.vehicleId) {
          // No truck-type text on this row — fall back to the Vehicle Master lookup by plate number.
          rateEntry = matchRateViaVehicleMaster(d.vehicleId, dropCountByVehicle[d.vehicleId] || 1);
        } else {
          return; // neither truck-type nor vehicle plate available — genuinely no info, not guessed
        }
        if (!rateEntry) { var uk = d.truckType || ('Vehicle ' + d.vehicleId); unmatchedTruckTypes[uk] = (unmatchedTruckTypes[uk] || 0) + 1; return; }

        if (!byType[rateEntry.label]) byType[rateEntry.label] = { rate: rateEntry.rate, drop_count: 0, vehicles: {} };
        byType[rateEntry.label].drop_count++;
        if (d.vehicleId) byType[rateEntry.label].vehicles[d.vehicleId] = 1;

        var city = d.city || 'Unknown';
        if (!byRegion[city]) byRegion[city] = {};
        if (!byRegion[city][rateEntry.label]) byRegion[city][rateEntry.label] = { rate: rateEntry.rate, drop_count: 0 };
        byRegion[city][rateEntry.label].drop_count++;

        var hasFood = !!d.types['food'];
        var hasNonFood = !!d.types['nonfood'];
        var foodCategory = (hasFood && hasNonFood) ? 'Mixed (Partition)' : (hasFood ? 'Food' : (hasNonFood ? 'Non-Food' : 'Other (3PL/Van)'));
        if (!byFoodType[foodCategory]) byFoodType[foodCategory] = { drop_count: 0, estimated_cost: 0 };
        byFoodType[foodCategory].drop_count++;
        byFoodType[foodCategory].estimated_cost += rateEntry.rate;

        var tempCategory = rateEntry.label.indexOf('Frozen') === 0 ? 'Frozen' : (rateEntry.label.indexOf('Ambient') === 0 ? 'Ambient' : 'Other');
        var tempFoodKey = tempCategory + ' · ' + foodCategory;
        if (!byTempFood[tempFoodKey]) byTempFood[tempFoodKey] = { drop_count: 0, estimated_cost: 0 };
        byTempFood[tempFoodKey].drop_count++;
        byTempFood[tempFoodKey].estimated_cost += rateEntry.rate;
      });

      var byTypeArr = Object.keys(byType).map(function(label){
        var d = byType[label];
        return { label: label, rate: d.rate, drop_count: d.drop_count, vehicle_count: Object.keys(d.vehicles).length, estimated_cost: d.drop_count * d.rate };
      }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });

      var byRegionArr = Object.keys(byRegion).map(function(city){
        var types = Object.keys(byRegion[city]).map(function(label){
          var d = byRegion[city][label];
          return { label: label, rate: d.rate, drop_count: d.drop_count, estimated_cost: d.drop_count * d.rate };
        }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });
        var regionCost = types.reduce(function(s,t){ return s + t.estimated_cost; }, 0);
        var regionDrops = types.reduce(function(s,t){ return s + t.drop_count; }, 0);
        return { city: city, types: types, total_cost: regionCost, total_drops: regionDrops };
      }).sort(function(a,b){ return b.total_cost - a.total_cost; });

      var byFoodTypeArr = Object.keys(byFoodType).map(function(cat){
        return { category: cat, drop_count: byFoodType[cat].drop_count, estimated_cost: byFoodType[cat].estimated_cost };
      }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });

      var byTempFoodArr = Object.keys(byTempFood).map(function(cat){
        return { category: cat, drop_count: byTempFood[cat].drop_count, estimated_cost: byTempFood[cat].estimated_cost };
      }).sort(function(a,b){ return b.estimated_cost - a.estimated_cost; });

      var totalCost = byTypeArr.reduce(function(s,t){ return s + t.estimated_cost; }, 0);
      var totalVehicles = byTypeArr.reduce(function(s,t){ return s + t.vehicle_count; }, 0);
      var totalDropsBilled = byTypeArr.reduce(function(s,t){ return s + t.drop_count; }, 0);

      return {
        available: byTypeArr.length > 0,
        by_type: byTypeArr,
        by_region: byRegionArr,
        by_food_type: byFoodTypeArr,
        by_temp_food_type: byTempFoodArr,
        total_estimated_cost: totalCost,
        total_vehicles: totalVehicles,
        total_drops_billed: totalDropsBilled,
        unmatched_truck_types: unmatchedTruckTypes
      };
    })(),
    cost_analysis: {
      own_fleet_drops: ownFleetDrops,
      pl_drops: plDrops,
      repeat_location_count: repeatLocations.length,
      repeat_location_avoidable_count: repeatLocationAvoidableCount
    },
    repeat_locations: repeatLocations,
    // Salon — per Azhar's explicit correction: DGC in the Main file is mixed with
    // Pharma W/H, NOT reliably Salon, so it must never be used here. This is built
    // ONLY from the standalone Salon (DIP W/H) upload's own rows: order count, a
    // drop count (distinct location), and a sales value IF that file ever includes
    // one (it currently doesn't, so this is 0 until they add it). Per
    // COST_TRACKING_EXCLUDED_SOURCES above, these rows never feed vehicle, driver,
    // or transport-cost tracking, and this summary itself carries no driver/vehicle
    // fields either — orders, drops, and order references only.
    salon_summary: (function () {
      var salonRows = rows.filter(function (r) { return r.__dispatch_source === 'Salon'; });
      var dropKeys = {}, value = 0, orderRefs = [];
      salonRows.forEach(function (r) {
        var locKey = r['LOCATION_NAME'] || r['LOCATION_ID'] || r['Location Name'] || r['ADDRESS'] || '';
        if (locKey) dropKeys[toStr(locKey)] = true;
        var amt = parseFloat(r['TOTAL_AMOUNT'] || r['AMOUNT'] || r['VALUE'] || 0) || 0;
        value += amt;
        var ref = r['TASK_ID'] || r['ORDER CODE'] || r['SOURCE_ORDER_ID'] || '';
        if (ref) orderRefs.push(toStr(ref));
      });
      return {
        orders: salonRows.length,
        drops: Object.keys(dropKeys).length,
        value: Math.round(value),
        order_references: orderRefs
      };
    })()
  };
}

//  ORDER-LEVEL EXTRACTION FOR RE-DELIVERY TRACKING 
// Extracts a lightweight per-row record (order code, customer, value, route, org, type)
// from a dispatch file, used to detect the SAME order code appearing again on a LATER
// dispatch date (a true re-delivery — failed first attempt, re-attempted later), as
// opposed to a same-day duplicate.
function extractOrderRows(buffer) {
  var wb = XLSX.read(buffer, { type: 'buffer', dense: true, cellDates: false, cellNF: false, cellHTML: false, cellFormula: false });
  var sheetName = findDataSheet(wb);
  var ws = wb.Sheets[sheetName];
  var rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  return extractOrderRowsFromRows(rows);
}

function extractOrderRowsFromRows(rows, sourceLabel) {
  if (!rows.length) return [];

  function findCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
    }) || null;
  }
  // Exact header match, same fallback-only role as in parseDispatch: "ORDER CODE" is
  // confirmed stable across days for the same re-delivered order (verified against real
  // files), so it's preferred. Plain "ORDER" is only used if "ORDER CODE" is missing.
  function findExactCol() {
    var names = Array.prototype.slice.call(arguments);
    return Object.keys(rows[0]).find(function(k) {
      return names.some(function(n) { return k.trim().toUpperCase() === n.toUpperCase(); });
    }) || null;
  }
  var C = {
    orderCode: findCol('ORDER CODE', 'ORDER_CODE') || findExactCol('ORDER', 'ORDER NUMBER', 'ORDER_NUMBER') || findCol('TASK_ID', 'TASK ID') || findCol('ORDER '),
    customer: findCol('CUSTOMER NAME', 'CUSTOMER'),
    amount: findCol('TOTAL_AMOUNT', 'AMOUNT', 'VALUE'),
    route: findCol('ROUTE'),
    org: findCol('ORG') || findCol('BU') || findCol('ORGANIZATION') || findCol('ORG-BU'),
    type: findCol('TYPE'),
    temperature: findCol('TEMPERATURE'),
    city: findCol('CITY', 'AREA'),
    address: findCol('ADDRESS'),
    location: findCol('LOCATION_ID', 'LOCATION'),
    truckType: findCol('TRUCK TYPE', 'TRUCK_TYPE', 'VEHICLE TYPE', 'VEHICLE_TYPE', 'DROP TYPE', 'DROP_TYPE')
  };
  console.log('Re-delivery tracking cols:', JSON.stringify(C));
  if (!C.orderCode) return []; // no order code column in this file — can't track re-delivery

  var out = [];
  rows.forEach(function(row) {
    var code = toStr(row[C.orderCode]);
    if (!code) return;
    var typeHere = C.type ? normaliseType(row[C.type]) : '';
    var custHere = C.customer ? toStr(row[C.customer]).toUpperCase() : '';
    var isInternalVanHere = (typeHere === 'van') && custHere.indexOf('INTERNAL') !== -1;
    var cityHere = detectCityFromAddress(C.address ? row[C.address] : '', C.city ? row[C.city] : '');
    var rawLocHere = C.location ? toStr(row[C.location]) : '';
    // Same internal-hub collapse as the main dispatch parser, so a drop count built from
    // this table lines up with the one on the Daily Dispatch / Duplicate Drops panels.
    var locHere = isInternalVanHere ? ('INTERNAL-HUB::' + (cityHere || 'Unknown')) : rawLocHere;
    out.push({
      order_code: code,
      customer: C.customer ? toStr(row[C.customer]) : '',
      value: C.amount ? (parseFloat(row[C.amount]) || 0) : 0,
      route: C.route ? toStr(row[C.route]) : '',
      org: (C.org ? toStr(row[C.org]).toUpperCase() : '') || deriveOrgFromCode(code),
      city: cityHere,
      drop_type: typeHere,
      temperature: C.temperature ? toStr(row[C.temperature]).trim() : '',
      source_label: sourceLabel || '',
      location_id: locHere,
      truck_type: C.truckType ? toStr(row[C.truckType]) : ''
    });
  });
  return out;
}

async function saveOrderTracking(dateKey, orderRows) {
  try {
    await pool.query('DELETE FROM order_tracking WHERE date_key=$1', [dateKey]);
    if (!orderRows.length) return true;
    var CHUNK = 500;
    for (var i = 0; i < orderRows.length; i += CHUNK) {
      var chunk = orderRows.slice(i, i + CHUNK);
      var vals = [];
      var phs = [];
      var idx = 1;
      chunk.forEach(function(r) {
        phs.push('($' + idx + ',$' + (idx+1) + ',$' + (idx+2) + ',$' + (idx+3) + ',$' + (idx+4) + ',$' + (idx+5) + ',$' + (idx+6) + ',$' + (idx+7) + ',$' + (idx+8) + ',$' + (idx+9) + ',$' + (idx+10) + ',$' + (idx+11) + ')');
        vals.push(r.order_code, dateKey, r.customer, r.value, r.route, r.org, r.drop_type, r.temperature, r.source_label || '', r.city || '', r.location_id || '', r.truck_type || '');
        idx += 12;
      });
      await pool.query('INSERT INTO order_tracking (order_code, date_key, customer, value, route, org, drop_type, temperature, source_label, city, location_id, truck_type) VALUES ' + phs.join(','), vals);
    }
    return true;
  } catch(e) {
    console.error('saveOrderTracking error:', e.message);
    return false;
  }
}

//  DISPATCH MEMORY (+ DB) 
var dispatchHistory = {};
var currentDispatch = null;

async function loadDispatchFromDB() {
  try {
    var rows = await dbLoadDispatch();
    if (rows.length > 0) {
      rows.forEach(function(r) {
        dispatchHistory[r.date_key] = {
          uploadedAt: r.uploaded_at, uploadedBy: r.uploaded_by,
          summary: r.summary, csvText: r.csv_text, date: r.date_key
        };
      });
      var latest = rows[0];
      currentDispatch = dispatchHistory[latest.date_key];
      console.log('Loaded', rows.length, 'dispatch dates from DB');
      return true;
    }
  } catch(e) { console.error('loadDispatchFromDB error:', e.message); }
  // Fallback to file
  try {
    var saved = loadJSON(DISPATCH_FILE);
    if (saved) {
      dispatchHistory = saved.history || {};
      var keys = Object.keys(dispatchHistory).sort().reverse();
      if (keys.length) currentDispatch = dispatchHistory[keys[0]];
      console.log('Loaded dispatch from file:', keys.length, 'dates');
    }
  } catch(e) { console.error('loadDispatchFromDB file fallback error:', e.message); }
  return false;
}
loadDispatchFromDB();

async function dispatchUploadHandler(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file received' });
    var dateKey    = req.body.dateKey    || new Date().toISOString().split('T')[0];
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var sourceLabel = (req.body.source || 'Main').trim() || 'Main';

    // ── Sub-admin upload guardrails. Superadmin is fully exempt: unlimited
    // uploads, any date/source, no cooldown. Dispatch stores each (date,
    // source) pair separately — e.g. "Main" and "Salon" can both exist for
    // the same day — so the overwrite check here is per exact (date, source),
    // not "any date before today" like Rejection. That still blocks the exact
    // failure mode (silently re-uploading over an already-saved day/source)
    // without breaking the legitimate multi-source-per-day workflow. ──
    if (req.user.role === 'subadmin') {
      try {
        var dsCooldown = await checkSubadminCooldown(req.user.uid, 'dispatch');
        if (dsCooldown) {
          return res.status(429).json({
            error: '\u23f3 Sub-admins can upload twice every 10 hours \u2014 you have used both uploads, you cannot re-upload for ' + dsCooldown.hoursLeft + 'h ' + dsCooldown.minsLeft + 'm.',
            cooldownActive: true, hoursLeft: dsCooldown.hoursLeft, minsLeft: dsCooldown.minsLeft
          });
        }
      } catch (e) { console.error('Dispatch: cooldown check failed:', e.message); }

      try {
        var dsExisting = await pool.query('SELECT 1 FROM dispatch_data_sources WHERE date_key=$1 AND source_label=$2', [dateKey, sourceLabel]);
        if (dsExisting.rows.length) {
          // Previously this blocked sub-admins and forced a superadmin
          // approval step for any re-upload of a date/source already on
          // the dashboard — even routine corrections. That meant the
          // corrected numbers sat invisible until manually approved,
          // which looked like data had "disappeared." Sub-admin re-uploads
          // now apply immediately (the cooldown above is still the rate
          // limit), and are simply logged here so there's still a record
          // of who overwrote what and when.
          auditLog(null, req.user.username, 'DISPATCH_REUPLOAD_OVERWRITE', 'Daily Dispatch (' + sourceLabel + ', ' + dateKey + ') overwritten by sub-admin: ' + req.file.originalname, req.headers['x-forwarded-for'] || req.ip || '');
        }
      } catch (e) { console.error('Dispatch: date-cutoff check failed:', e.message); }
    }

    var wbThis = XLSX.read(req.file.buffer, { type: 'buffer', dense: true, cellDates: false, cellNF: false, cellHTML: false });
    var sheetThis = wbThis.Sheets[findDataSheet(wbThis)];
    var rowsThis = XLSX.utils.sheet_to_json(sheetThis, { defval: '', raw: true });
    if (!rowsThis.length) return res.status(400).json({ error: 'Could not parse file — no data rows found' });
    var csvThis = XLSX.utils.sheet_to_csv(sheetThis);

    // Store this source's rows under its own label — never overwrites a different
    // source's rows for the same date, only replaces this exact source on re-upload.
    await pool.query(
      `INSERT INTO dispatch_data_sources (date_key, source_label, uploaded_by, file_name, rows_json, row_count)
       VALUES ($1,$2,$3,$4,$5,$6)
       ON CONFLICT (date_key, source_label) DO UPDATE SET uploaded_by=$3, file_name=$4, rows_json=$5, row_count=$6, uploaded_at=NOW()`,
      [dateKey, sourceLabel, uploadedBy, req.file.originalname, JSON.stringify(rowsThis), rowsThis.length]
    );

    // Safety net for dates uploaded BEFORE this multi-source feature existed: that data
    // lives only as a computed summary in dispatch_data, with no raw rows in
    // dispatch_data_sources — so it would be invisible to the merge below and get wiped
    // out by an upload of a different source. Recover it once, from its stored CSV, and
    // register it as a "Main" source so it merges normally from here on.
    if (sourceLabel !== 'Main') {
      var existingSourcesCheck = await pool.query('SELECT 1 FROM dispatch_data_sources WHERE date_key=$1 AND source_label=$2', [dateKey, 'Main']);
      if (!existingSourcesCheck.rows.length) {
        var legacyRes = await pool.query('SELECT csv_text FROM dispatch_data WHERE date_key=$1', [dateKey]);
        if (legacyRes.rows.length && legacyRes.rows[0].csv_text) {
          try {
            var legacyWb = XLSX.read(legacyRes.rows[0].csv_text, { type: 'string', raw: true });
            var legacyRows = XLSX.utils.sheet_to_json(legacyWb.Sheets[legacyWb.SheetNames[0]], { defval: '', raw: true });
            if (legacyRows.length) {
              await pool.query(
                `INSERT INTO dispatch_data_sources (date_key, source_label, uploaded_by, file_name, rows_json, row_count)
                 VALUES ($1,'Main',$2,$3,$4,$5)
                 ON CONFLICT (date_key, source_label) DO NOTHING`,
                [dateKey, 'Recovered', 'Recovered from pre-existing data for ' + dateKey, JSON.stringify(legacyRows), legacyRows.length]
              );
              console.log('Backfilled legacy Main data for', dateKey, '—', legacyRows.length, 'rows recovered from stored CSV');
            }
          } catch (legacyErr) {
            console.error('Legacy backfill failed for', dateKey, ':', legacyErr.message, '— this date\'s pre-existing data may be at risk of being overwritten.');
          }
        }
      }
    }

    // Merge EVERY source's rows for this date (Main + Salon + whatever else) into one
    // combined dataset, then run the exact same aggregation the app has always used —
    // this is what every other dashboard feature (drops, transport cost, reconciliation)
    // reads, so they all pick up the merge automatically.
    var allSourcesRes = await pool.query('SELECT source_label, rows_json, file_name FROM dispatch_data_sources WHERE date_key=$1 ORDER BY source_label', [dateKey]);
    var mergedRows = [];
    var sourcesIncluded = [];
    allSourcesRes.rows.forEach(function (s) {
      var taggedRows = s.rows_json.map(function (r) {
        // Non-destructive: adds one extra key, doesn't touch the row's real columns.
        return Object.assign({}, r, { __dispatch_source: s.source_label });
      });
      mergedRows = mergedRows.concat(taggedRows);
      sourcesIncluded.push(s.source_label + ' (' + s.rows_json.length + ' rows)');
    });

    var summary = parseDispatchFromRows(mergedRows);
    if (!summary) return res.status(400).json({ error: 'Could not compute summary from merged rows' });

    // Combined CSV kept for the AI Q&A / raw-data features — each source's own CSV,
    // concatenated with its label, rather than trying to force mismatched headers
    // into one table.
    var combinedCsv = allSourcesRes.rows.map(function (s) {
      var label = '=== SOURCE: ' + s.source_label + ' (' + s.file_name + ') ===';
      return label + '\n' + (s.source_label === sourceLabel ? csvThis : XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(s.rows_json)));
    }).join('\n\n');

    var dbOk = await dbSaveDispatch(dateKey, uploadedBy, summary, combinedCsv);
    var entry = { uploadedAt:new Date().toISOString(), uploadedBy:uploadedBy, csvText:combinedCsv.substring(0,200000), summary:summary, date:dateKey };
    dispatchHistory[dateKey] = entry;
    currentDispatch = entry;
    var keys = Object.keys(dispatchHistory).sort();
    while (keys.length > 180) delete dispatchHistory[keys.shift()];
    saveJSON(DISPATCH_FILE, { history:dispatchHistory });
    console.log('Dispatch saved:', dateKey, dbOk ? '(DB+file)' : '(file only)', '— sources:', sourcesIncluded.join(', '));

    try {
      // Extracted per-source, not from the merged array — different sources can use
      // completely different headers (Main: "ORDER CODE", Salon: "TASK_ID"), and a single
      // column-detection pass over merged rows would only ever see whichever source's
      // header names happened to appear first.
      var orderRows = [];
      allSourcesRes.rows.forEach(function (s) {
        orderRows = orderRows.concat(extractOrderRowsFromRows(s.rows_json, s.source_label));
      });
      await saveOrderTracking(dateKey, orderRows);
      console.log('Order tracking saved:', dateKey, orderRows.length, 'order rows (merged across', sourcesIncluded.length, 'source(s))');
    } catch(trackErr) {
      console.error('Order tracking error (non-fatal):', trackErr.message);
    }
    auditLog(null, uploadedBy, 'UPLOAD', 'Daily Dispatch (' + sourceLabel + ', ' + dateKey + '): ' + (req.file.originalname || ''), req.headers['x-forwarded-for'] || req.ip || '');
    if (req.user.role === 'subadmin') {
      try { await recordSubadminUpload(req.user.uid, 'dispatch'); }
      catch (e) { console.error('Dispatch: could not record upload cooldown:', e.message); }
    }
    res.json({ success:true, summary:summary, uploadedAt:entry.uploadedAt, date:dateKey, source_uploaded: sourceLabel, sources_merged: sourcesIncluded });
  } catch(e) {
    console.error('Dispatch upload error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
}
app.post('/api/dispatch/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), dispatchUploadHandler);

// ── Upload Approvals: sub-admin requests to overwrite/re-upload dates that
// already exist on a dashboard. Super admin only sees & decides all requests;
// a sub-admin can see the status of their own. ──
app.get('/api/upload-approvals', requireAuth, requireRole('superadmin', 'subadmin'), async function(req, res) {
  try {
    var params = [];
    var q = 'SELECT id, user_id, username, endpoint, file_name, blocked_dates, reason, meta, status, requested_at, decided_by, decision_comment, decided_at FROM upload_approval_requests';
    var where = [];
    if (req.user.role !== 'superadmin') { where.push('user_id=$' + (params.length + 1)); params.push(req.user.uid); }
    if (req.query.status) { where.push('status=$' + (params.length + 1)); params.push(req.query.status); }
    if (where.length) q += ' WHERE ' + where.join(' AND ');
    q += ' ORDER BY requested_at DESC LIMIT 200';
    var r = await pool.query(q, params);
    res.json({ requests: r.rows });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/upload-approvals/:id/approve', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var id = parseInt(req.params.id);
    var reqRow = (await pool.query('SELECT * FROM upload_approval_requests WHERE id=$1', [id])).rows[0];
    if (!reqRow) return res.status(404).json({ error: 'Request not found' });
    if (reqRow.status !== 'pending') return res.status(409).json({ error: 'This request was already ' + reqRow.status });
    if (!reqRow.file_data) return res.status(410).json({ error: 'The original file is no longer available for this request' });

    var meta = reqRow.meta || {};
    var fakeReq = {
      file: { buffer: reqRow.file_data, originalname: reqRow.file_name },
      user: { role: 'superadmin', uid: req.user.uid, username: req.user.username },
      body: Object.assign({ uploadedBy: reqRow.username, force: 'true' }, meta),
      headers: req.headers
    };
    var captured = { statusCode: 200, body: null };
    var fakeRes = {
      status: function (c) { captured.statusCode = c; return this; },
      json: function (b) { captured.body = b; return this; },
      headersSent: false
    };

    if (reqRow.endpoint === 'rejection') await rejectionUploadHandler(fakeReq, fakeRes);
    else if (reqRow.endpoint === 'dispatch') await dispatchUploadHandler(fakeReq, fakeRes);
    else return res.status(400).json({ error: 'Unknown upload type: ' + reqRow.endpoint });

    var execOk = !!(captured.body && captured.body.success);
    var comment = toStr(req.body.comment);
    await pool.query(
      `UPDATE upload_approval_requests SET status=$1, decided_by=$2, decision_comment=$3, decided_at=NOW(), file_data=NULL WHERE id=$4`,
      [execOk ? 'approved' : 'approve_failed', req.user.username, comment, id]
    );
    auditLog(null, req.user.username, 'UPLOAD_APPROVAL_DECISION',
      (execOk ? 'Approved' : 'Approve FAILED for') + ' upload #' + id + ' (' + reqRow.endpoint + ', ' + reqRow.file_name + ') requested by ' + reqRow.username +
      (comment ? ' \u2014 ' + comment : '') + (execOk ? '' : ' \u2014 error: ' + (captured.body && captured.body.error)),
      req.headers['x-forwarded-for'] || req.ip || '');
    res.status(execOk ? 200 : 500).json({ success: execOk, executionResult: captured.body });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/upload-approvals/:id/reject', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var id = parseInt(req.params.id);
    var comment = toStr(req.body.comment);
    var r = await pool.query(
      `UPDATE upload_approval_requests SET status='rejected', decided_by=$1, decision_comment=$2, decided_at=NOW(), file_data=NULL WHERE id=$3 AND status='pending' RETURNING id, username, endpoint, file_name`,
      [req.user.username, comment, id]
    );
    if (!r.rows[0]) return res.status(404).json({ error: 'Request not found or already decided' });
    auditLog(null, req.user.username, 'UPLOAD_APPROVAL_DECISION',
      'Rejected upload #' + id + ' (' + r.rows[0].endpoint + ', ' + r.rows[0].file_name + ') requested by ' + r.rows[0].username + (comment ? ' \u2014 ' + comment : ''),
      req.headers['x-forwarded-for'] || req.ip || '');
    res.json({ success: true });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/dispatch/sources/:dateKey', requireAuth, async function (req, res) {
  try {
    var r = await pool.query('SELECT source_label, uploaded_at, uploaded_by, file_name, row_count FROM dispatch_data_sources WHERE date_key=$1 ORDER BY source_label', [req.params.dateKey]);
    res.json({ date: req.params.dateKey, sources: r.rows });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/dispatch/status', function(req, res) {
  var avail = Object.keys(dispatchHistory).sort().reverse();
  if (!currentDispatch) return res.json({ hasData:false, availableDates:avail });
  res.json({ hasData:true, uploadedAt:currentDispatch.uploadedAt, uploadedBy:currentDispatch.uploadedBy, summary:currentDispatch.summary, date:currentDispatch.date, availableDates:avail });
});

app.get('/api/dispatch/date/:dateKey', function(req, res) {
  var entry = dispatchHistory[req.params.dateKey];
  if (!entry) return res.json({ hasData:false });
  currentDispatch = entry;
  res.json({ hasData:true, uploadedAt:entry.uploadedAt, uploadedBy:entry.uploadedBy, summary:entry.summary, date:entry.date });
});

// Re-delivery tracking: finds order codes active on the given date that ALSO appear on
// any EARLIER dispatch date — i.e. the same order was dispatched before, presumably
// failed, and is being re-delivered now. Same-day duplicates are not counted here.
app.get('/api/dispatch/redelivery/:dateKey', async function(req, res) {
  try {
    var dateKey = req.params.dateKey;
    var todayRes = await pool.query('SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1', [dateKey]);
    var todayCodes = todayRes.rows.map(function(r){ return r.order_code; });
    if (!todayCodes.length) return res.json({ hasData:true, dateKey:dateKey, total_repeated_orders:0, total_value_at_risk:0, orders:[] });

    var histRes = await pool.query(
      'SELECT order_code, date_key, customer, value, route, org, drop_type, temperature FROM order_tracking WHERE order_code = ANY($1) AND date_key <= $2 ORDER BY date_key ASC',
      [todayCodes, dateKey]
    );

    var byCode = {};
    histRes.rows.forEach(function(r) {
      if (!byCode[r.order_code]) byCode[r.order_code] = [];
      byCode[r.order_code].push(r);
    });

    var repeated = [];
    Object.keys(byCode).forEach(function(code) {
      var occ = byCode[code];
      var distinctDates = Array.from(new Set(occ.map(function(o) {
        var dk = o.date_key;
        return (dk && dk.toISOString) ? dk.toISOString().split('T')[0] : String(dk).split('T')[0];
      })));
      if (distinctDates.length > 1) {
        var latest = occ[occ.length - 1];
        repeated.push({
          order_code: code,
          customer: latest.customer || '',
          value: parseFloat(latest.value) || 0,
          org: latest.org || '',
          route: latest.route || '',
          drop_type: latest.drop_type || '',
          temperature: latest.temperature || '',
          times_delivered: distinctDates.length,
          dates: distinctDates
        });
      }
    });
    repeated.sort(function(a, b) { return b.value - a.value; });

    var totalValue = repeated.reduce(function(s, r) { return s + r.value; }, 0);
    res.json({
      hasData: true,
      dateKey: dateKey,
      total_repeated_orders: repeated.length,
      total_value_at_risk: Math.round(totalValue),
      orders: repeated.slice(0, 200)
    });
  } catch(e) {
    console.error('redelivery endpoint error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// ============================================================
// TRANSPORT COST RECONCILIATION — cross-checks the transport team's own
// reported drop counts (their raw "Distinct DROP ID" export) against what
// this app captured from the daily dispatch uploads, so discrepancies show
// up before the monthly invoice arrives rather than after paying it.
// ============================================================

// The transport team's own file only ever covers these 4 business-unit org codes
// (DCV, DGC, DGS, DSN) by default. 3PL/HCP orders are delivered by a third party, and
// DCF food service orders are self-delivered on DCF's own van — but ONLY in Dubai; DCF
// orders in other emirates still go through the shared transport team and should be
// compared normally. Every query below therefore uses this list PLUS a city-aware OR
// clause for DCF, rather than a flat org exclusion.
var RECON_ELIGIBLE_ORG_LIST = ['DCV', 'DGC', 'DGS', 'DSN'];

// Every reconciliation read/export route below gets this — without it, browsers can
// (and, confirmed in practice, do) serve back a stale cached Excel/JSON response even
// after the underlying data has genuinely changed server-side, making a real fix look
// like it "didn't work." This forces a fresh request every single time.
function noCache(req, res, next) {
  res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
  res.setHeader('Pragma', 'no-cache');
  res.setHeader('Expires', '0');
  next();
}

function normalizeReconcileHeader(h) { return String(h || '').toLowerCase().replace(/[^a-z]/g, ''); }
function matchReconcileField(header) {
  var h = normalizeReconcileHeader(header);
  if (h.indexOf('dropid') !== -1) return 'drop_id';
  if (h.indexOf('taskstatus') !== -1) return 'task_status';
  if (h.indexOf('taskid') !== -1) return 'task_order_id';
  if (h.indexOf('finaldropsremark') !== -1) return 'drop_class';
  if (h.indexOf('operatingunit') !== -1) return 'operating_unit';
  if (h.indexOf('dispatchdate') !== -1 || h === 'date') return 'dispatch_date';
  // Optional — only used so "In transport file, missing from AKI dispatch" can show
  // WHO the order was for, instead of just a bare code with no way to investigate it.
  // Not every transport export will have these; when absent, those columns just stay
  // blank rather than breaking anything.
  if (h.indexOf('customername') !== -1 || h === 'customer') return 'customer';
  if (h.indexOf('locationname') !== -1 || h === 'location') return 'location_name';
  if (h.indexOf('address') !== -1) return 'address';
  return null;
}
// Transport's TASK ID sometimes carries a re-attempt suffix like
// "-2026-07-01-1" that the AKI dispatch file's ORDER CODE never has —
// stripping it is what lets the two systems' order identifiers line up
// (confirmed: 923 of 930 real orders matched exactly once stripped).
function stripTaskIdSuffix(v) {
  return String(v || '').trim().replace(/-\d{4}-\d{2}-\d{2}(-\d+)?$/, '');
}

app.post('/api/dispatch/reconcile/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), async function (req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    var wb;
    try { wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true }); }
    catch (e) { return res.status(400).json({ error: 'Could not read that file.' }); }

    // Find the sheet that actually has DROP ID + TASK STATUS + Final Drops Remarks +
    // a date column — the transport team's export has multiple sheets (Summary,
    // Bulk, Multi pivots) and only "Raw Data" has what's needed here.
    var bestSheetName = wb.SheetNames[0], bestScore = -1;
    wb.SheetNames.forEach(function (name) {
      var s = wb.Sheets[name];
      var firstRow = XLSX.utils.sheet_to_json(s, { header: 1, defval: '' })[0] || [];
      var score = 0, hasDropId = false, hasStatus = false, hasClass = false, hasDate = false;
      firstRow.forEach(function (cell) {
        var f = matchReconcileField(cell);
        if (f) { score++; if (f === 'drop_id') hasDropId = true; if (f === 'task_status') hasStatus = true; if (f === 'drop_class') hasClass = true; if (f === 'dispatch_date') hasDate = true; }
      });
      if (hasDropId && hasStatus && hasClass && hasDate && score > bestScore) { bestScore = score; bestSheetName = name; }
    });
    if (bestScore === -1) return res.status(400).json({ error: 'Could not find a sheet with DROP ID, TASK STATUS, Final Drops Remarks, and a date column. This should be the "Raw Data" sheet from the transport team\'s export.' });

    var sheet = wb.Sheets[bestSheetName];
    var rawRows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
    if (!rawRows.length) return res.status(400).json({ error: 'The file has no data rows.' });

    var fieldByHeader = {};
    Object.keys(rawRows[0]).forEach(function (h) { var f = matchReconcileField(h); if (f) fieldByHeader[h] = f; });

    // TASK_STATUS_RULE documents the assumption in effect — currently COMPLETED-only,
    // matching how a transport-team invoice should only bill for completed drops.
    // If Azhar confirms a different rule with the transport team, this is the one
    // line to change.
    // Per Azhar's direction: a dispatched drop is a dispatched drop, whatever
    // the transport team's internal TASK STATUS says — that status describes
    // THEIR workflow, not whether the trip happened. So the primary count
    // below includes every status. COMPLETED is tracked in parallel purely as
    // supplementary "how much of this day is closed out" context.
    var COMPLETED_STATUS = 'COMPLETED';

    var byDateClass = {}; // "date|class" -> Set of drop_ids (ALL statuses — this is the real count)
    var completedByDateClass = {}; // "date|class" -> Set of drop_ids where status === COMPLETED (context only)
    var orderIdsByDate = {}; // date -> { stripped_task_id -> {customer, location_name, address} }, for order-level diffing
    var totalRows = 0, skippedNonConsumer = 0;
    rawRows.forEach(function (raw) {
      var rec = {};
      Object.keys(raw).forEach(function (h) { var f = fieldByHeader[h]; if (f) rec[f] = raw[h]; });
      if (!rec.drop_id || !rec.dispatch_date) return;
      // Only filter by OPERATING UNIT when that column actually exists in this file — a
      // Consumer-only export (no such column) is trusted as-is. A full multi-department
      // export gets filtered down to Consumer here, so uploading either kind of file
      // produces the same Consumer-only reconciliation.
      if (rec.operating_unit !== undefined && String(rec.operating_unit).trim().toUpperCase() !== 'CONSUMER') {
        skippedNonConsumer++;
        return;
      }
      totalRows++;
      var dateStr = toDateStrGeneric(rec.dispatch_date);
      if (!dateStr) return;
      var cls = String(rec.drop_class || 'Other').trim() || 'Other';
      var key = dateStr + '|' + cls;
      if (!byDateClass[key]) byDateClass[key] = new Set();
      byDateClass[key].add(String(rec.drop_id));
      var status = String(rec.task_status || '').trim().toUpperCase();
      if (status === COMPLETED_STATUS) {
        if (!completedByDateClass[key]) completedByDateClass[key] = new Set();
        completedByDateClass[key].add(String(rec.drop_id));
      }
      if (rec.task_order_id) {
        var stripped = stripTaskIdSuffix(rec.task_order_id);
        if (stripped) {
          if (!orderIdsByDate[dateStr]) orderIdsByDate[dateStr] = {};
          // First row seen for this order ID wins — later duplicate rows (e.g. a
          // split/-1/-2 shipment) don't need to overwrite it, same customer either way.
          if (!orderIdsByDate[dateStr][stripped]) {
            orderIdsByDate[dateStr][stripped] = {
              customer: rec.customer ? String(rec.customer).trim() : '',
              location_name: rec.location_name ? String(rec.location_name).trim() : '',
              address: rec.address ? String(rec.address).trim() : ''
            };
          }
        }
      }
    });

    var batchId = 'RECON-' + Date.now();
    var upserted = 0;
    for (var key in byDateClass) {
      var parts = key.split('|');
      var dateStr = parts[0], cls = parts[1];
      var count = byDateClass[key].size;
      var completedCount = completedByDateClass[key] ? completedByDateClass[key].size : 0;
      await pool.query(
        `INSERT INTO transport_drop_reconciliation (date_key, drop_class, reported_drops, completed_drops, task_status_rule, upload_batch_id, uploaded_by)
         VALUES ($1,$2,$3,$4,'ALL',$5,$6)
         ON CONFLICT (date_key, drop_class) DO UPDATE SET reported_drops=$3, completed_drops=$4, task_status_rule='ALL', upload_batch_id=$5, uploaded_by=$6, uploaded_at=NOW()`,
        [dateStr, cls, count, completedCount, batchId, req.user ? req.user.username : 'Admin']
      );
      upserted++;
    }

    var orderIdsSaved = 0;
    for (var od in orderIdsByDate) {
      // Clear this date's previous order IDs first so a re-upload doesn't leave stale entries.
      await pool.query('DELETE FROM transport_order_ids WHERE date_key=$1', [od]);
      var ids = Object.keys(orderIdsByDate[od]);
      for (var i = 0; i < ids.length; i += 400) {
        var chunk = ids.slice(i, i + 400);
        var phs = [], vals = [od], idx = 2;
        chunk.forEach(function (id) {
          var meta = orderIdsByDate[od][id];
          phs.push('($1,$' + idx + ',$' + (idx+1) + ',$' + (idx+2) + ',$' + (idx+3) + ',$' + (idx+4) + ')');
          vals.push(id, meta.customer, meta.location_name, meta.address, batchId);
          idx += 5;
        });
        await pool.query(
          'INSERT INTO transport_order_ids (date_key, order_id, customer, location_name, address, upload_batch_id) VALUES ' + phs.join(',') + ' ON CONFLICT (date_key, order_id) DO NOTHING',
          vals
        );
      }
      orderIdsSaved += ids.length;
    }

    res.json({
      batch_id: batchId,
      total_rows: totalRows,
      task_status_rule: 'ALL (every dispatched drop, any status)',
      date_class_buckets_saved: upserted,
      order_ids_saved: orderIdsSaved,
      order_level_diff_available: orderIdsSaved > 0,
      skipped_non_consumer: skippedNonConsumer,
      sheet_used: bestSheetName
    });
  } catch (e) {
    console.error('reconcile upload error:', e.message);
    res.status(500).json({ error: 'Upload failed: ' + e.message });
  }
});

// ONE-TIME MIGRATION — recomputes total_drops_reconcile (and refreshes the whole
// summary — including, as of this update, per-driver order_details for Driver Order
// Details) for every date already in dispatch_data, WITHOUT requiring a re-upload.
// Reads back the raw rows this app already has stored (dispatch_data_sources per
// date; falls back to the stored combined CSV for dates saved before the
// multi-source feature existed) and re-runs the same parseDispatchFromRows used at
// upload time. Safe to call more than once — it only ever recomputes and overwrites
// the `summary` column, nothing else (uploaded_by, csv_text, uploaded_at untouched).
app.post('/api/dispatch/reconcile/migrate-historical', requireAuth, requireRole('superadmin'), async function (req, res) {
  try {
    var dateRows = await pool.query('SELECT date_key::text FROM dispatch_data ORDER BY date_key');
    var results = [];
    for (var i = 0; i < dateRows.rows.length; i++) {
      var dateKey = dateRows.rows[i].date_key;
      var summary = null;
      var orderRows = [];
      var orderTrackingRebuilt = false;
      try {
        var sourcesRes = await pool.query('SELECT source_label, rows_json FROM dispatch_data_sources WHERE date_key=$1 ORDER BY source_label', [dateKey]);
        if (sourcesRes.rows.length) {
          var mergedRows = [];
          sourcesRes.rows.forEach(function (s) {
            var tagged = s.rows_json.map(function (r) { return Object.assign({}, r, { __dispatch_source: s.source_label }); });
            mergedRows = mergedRows.concat(tagged);
            // Rebuild order_tracking too — from the SAME raw rows this app already has
            // stored, using the corrected column-detection logic (TASK_ID checked before
            // the risky loose 'ORDER ' match) and tagging each row's real source_label
            // (so Salon can be excluded from the reconciliation diff without ever
            // re-uploading anything).
            orderRows = orderRows.concat(extractOrderRowsFromRows(s.rows_json, s.source_label));
          });
          summary = parseDispatchFromRows(mergedRows);
          orderTrackingRebuilt = true;
        } else {
          // No raw rows stored for this date (saved before the multi-source feature
          // existed) — re-parse from the combined CSV that was already stored. No
          // per-source label available here, so order_tracking rows get source_label=''
          // (same as before this fix — these very old dates simply have no Salon/Main
          // distinction to make).
          var legacyRes = await pool.query('SELECT csv_text FROM dispatch_data WHERE date_key=$1', [dateKey]);
          if (legacyRes.rows.length && legacyRes.rows[0].csv_text) {
            var legacyWb = XLSX.read(legacyRes.rows[0].csv_text, { type: 'string', raw: true });
            var legacyRows = XLSX.utils.sheet_to_json(legacyWb.Sheets[legacyWb.SheetNames[0]], { defval: '', raw: true });
            summary = parseDispatchFromRows(legacyRows);
            orderRows = extractOrderRowsFromRows(legacyRows);
            orderTrackingRebuilt = true;
          }
        }
      } catch (parseErr) {
        results.push({ date: dateKey, status: 'failed: ' + parseErr.message });
        continue;
      }
      if (!summary) { results.push({ date: dateKey, status: 'no data available, skipped' }); continue; }
      await pool.query('UPDATE dispatch_data SET summary=$2 WHERE date_key=$1', [dateKey, JSON.stringify(summary)]);
      if (dispatchHistory[dateKey]) dispatchHistory[dateKey].summary = summary;
      if (orderTrackingRebuilt) {
        try { await saveOrderTracking(dateKey, orderRows); }
        catch (trackErr) { console.error('migrate-historical order_tracking rebuild failed for', dateKey, ':', trackErr.message); }
      }
      results.push({
        date: dateKey, status: 'updated',
        total_drops_old: summary.total_drops,
        total_drops_reconcile_new: summary.total_drops_reconcile,
        order_rows_rebuilt: orderRows.length
      });
    }
    res.json({ success: true, dates_processed: results.length, results: results });
  } catch (e) {
    console.error('migrate-historical error:', e.message);
    res.status(500).json({ error: 'Migration failed: ' + e.message });
  }
});

app.get('/api/dispatch/reconcile', noCache, requireAuth, async function (req, res) {
  try {
    var params = [], clauses = [];
    if (req.query.date_from) { params.push(req.query.date_from); clauses.push('date_key >= $' + params.length); }
    if (req.query.date_to) { params.push(req.query.date_to); clauses.push('date_key <= $' + params.length); }
    var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';

    var transportRows = await pool.query(
      `SELECT date_key::text, drop_class, reported_drops, completed_drops, task_status_rule, uploaded_at FROM transport_drop_reconciliation ${where} ORDER BY date_key`, params
    );

    var byDate = {};
    transportRows.rows.forEach(function (r) {
      if (!byDate[r.date_key]) byDate[r.date_key] = { date: r.date_key, transport_total: 0, completed_total: 0, by_class: {}, uploaded_at: r.uploaded_at };
      byDate[r.date_key].transport_total += r.reported_drops;
      byDate[r.date_key].completed_total += r.completed_drops;
      byDate[r.date_key].by_class[r.drop_class] = r.reported_drops;
    });

    // Transport doesn't run on Sundays — those dates never get a transport file at all,
    // so they'd otherwise vanish from this list entirely even when the app DID dispatch
    // something that day. Rather than hide them (looks like missing data) or flag them
    // as a mismatch (looks like an error), pull in any Sunday that has app data but no
    // transport upload and mark it as a known, expected exception day.
    var dispatchDateRows = await pool.query(
      `SELECT date_key::text FROM dispatch_data ${where} ORDER BY date_key`, params
    );
    dispatchDateRows.rows.forEach(function (r) {
      var d = r.date_key;
      if (byDate[d]) return; // already has transport data, nothing to add
      var dayOfWeek = new Date(d + 'T00:00:00Z').getUTCDay(); // 0 = Sunday
      if (dayOfWeek !== 0) return; // only auto-add the known Sunday exception, not any other unexplained gap
      byDate[d] = { date: d, transport_total: 0, completed_total: 0, by_class: {}, uploaded_at: null, is_exception_day: true, exception_reason: 'Sunday — transport does not run, no report expected' };
    });

    if (!Object.keys(byDate).length) return res.json({ has_data: false, days: [] });

    var dateKeys = Object.keys(byDate).sort();
    var appRows = await pool.query(
      `SELECT date_key::text, summary FROM dispatch_data WHERE date_key = ANY($1::date[])`, [dateKeys]
    );
    var appByDate = {};
    appRows.rows.forEach(function (r) { appByDate[r.date_key] = r.summary; });

    var days = dateKeys.map(function (d) {
      var t = byDate[d];
      var appSummary = appByDate[d] || null;
      var appDrops = appSummary && appSummary.total_drops_reconcile !== undefined ? appSummary.total_drops_reconcile : (appSummary && appSummary.total_drops !== undefined ? appSummary.total_drops : null);
      var variance = (appDrops !== null && !t.is_exception_day) ? (appDrops - t.transport_total) : null;
      var variancePct = (variance !== null && t.transport_total > 0) ? +((variance / t.transport_total) * 100).toFixed(1) : null;
      var completedPct = t.transport_total > 0 ? +((t.completed_total / t.transport_total) * 100).toFixed(1) : null;
      return {
        date: d,
        app_captured: appDrops,
        app_has_data: appSummary !== null,
        transport_reported: t.transport_total,
        transport_by_class: t.by_class,
        transport_completed: t.completed_total,
        transport_completed_pct: completedPct,
        variance: variance,
        variance_pct: variancePct,
        is_exception_day: !!t.is_exception_day,
        exception_reason: t.exception_reason || null
      };
    });

    // Exception days are a known, expected gap — never counted toward the variance
    // totals (that would misrepresent a real cost/billing number), and never counted
    // toward "days with both sides" or "days missing app data" either, since neither
    // label fits a day where no transport report was ever going to exist.
    var totalApp = days.reduce(function (s, d) { return s + (d.is_exception_day ? 0 : (d.app_captured || 0)); }, 0);
    var totalTransportAllDays = days.reduce(function (s, d) { return s + (d.is_exception_day ? 0 : d.transport_reported); }, 0);
    var daysMissingAppData = days.filter(function (d) { return !d.is_exception_day && !d.app_has_data; }).length;

    // "Total Variance" must compare like with like: App Captured only ever has data for
    // days where a dispatch file was uploaded, so summing Transport across EVERY day
    // (including ones with no app upload at all) makes the headline variance compare an
    // incomplete App total against a complete Transport total — the two numbers were
    // structurally never going to agree, which is exactly the "-117 vs -111" mismatch
    // reported. Restricting Transport's total to the same matched days the App total
    // already uses (mirrors what the Excel export's "TOTAL (days with BOTH App +
    // Transport data)" row has always done) makes this an honest apples-to-apples number.
    var totalTransportMatched = days.reduce(function (s, d) { return s + ((d.app_has_data && !d.is_exception_day) ? d.transport_reported : 0); }, 0);
    var daysMatched = days.filter(function (d) { return d.app_has_data && !d.is_exception_day; }).length;

    res.json({
      has_data: true,
      days: days,
      total_app_captured: totalApp,
      total_transport_reported: totalTransportMatched,
      total_transport_reported_all_days: totalTransportAllDays,
      total_variance: totalApp - totalTransportMatched,
      days_missing_app_data: daysMissingAppData,
      days_matched: daysMatched,
      days_total: days.length,
      days_exception: days.filter(function (d) { return d.is_exception_day; }).length
    });
  } catch (e) {
    console.error('reconcile compare error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

// Order-level diff: which specific orders does the app have that transport
// doesn't (and vice versa) for one date — the actionable detail behind a
// day-level count mismatch. Matches on order_tracking (this app's own
// per-day order log, already populated from daily dispatch uploads) against
// transport_order_ids (populated from the transport file's TASK ID column).
app.get('/api/dispatch/reconcile/order-diff', noCache, requireAuth, async function (req, res) {
  try {
    var date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date is required, e.g. ?date=2026-07-18' });

    // Sunday — transport doesn't run, so an empty transport-side result here is expected,
    // not a mismatch. Flagged so the frontend can label it correctly instead of showing
    // every app order that day as a false "missing from transport" error.
    var isExceptionDay = new Date(date + 'T00:00:00Z').getUTCDay() === 0;

    var appRows = await pool.query("SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1 AND (org = ANY($2::text[]) OR (org = 'DCF' AND UPPER(COALESCE(city,'')) <> 'DUBAI'))", [date, RECON_ELIGIBLE_ORG_LIST]);
    var transportRows = await pool.query('SELECT order_id, customer, location_name, address FROM transport_order_ids WHERE date_key=$1', [date]);

    // Full visibility into what's actually stored for this date, with zero console
    // commands needed — shows every source label uploaded (Main, Salon, etc.) and every
    // org code seen, each with a count of how many order_tracking rows exist BEFORE the
    // org/Salon filter above is applied, so it's immediately obvious whether Salon data
    // even exists for this date, and whether the filter is doing anything.
    var sourceBreakdownRes = await pool.query(
      "SELECT COALESCE(NULLIF(source_label,''),'(unlabeled / pre-fix upload)') AS source_label, COUNT(*)::int AS cnt FROM order_tracking WHERE date_key=$1 GROUP BY 1 ORDER BY 1", [date]
    );
    var orgBreakdownRes = await pool.query(
      "SELECT COALESCE(NULLIF(org,''),'(blank)') AS org, COUNT(*)::int AS cnt FROM order_tracking WHERE date_key=$1 GROUP BY 1 ORDER BY 1", [date]
    );

    if (!appRows.rows.length && !transportRows.rows.length) {
      return res.json({ has_data: false, date: date });
    }

    var appSet = new Set(appRows.rows.map(function (r) { return r.order_code; }));
    var transportDetailById = {};
    transportRows.rows.forEach(function (r) { transportDetailById[r.order_id] = r; });
    var transportSet = new Set(transportRows.rows.map(function (r) { return r.order_id; }));

    var appOnly = Array.from(appSet).filter(function (o) { return !transportSet.has(o); });
    // "Missing from AKI dispatch" is the one that actually needs investigating — this is
    // where I attach who the order was for, so it's not just a bare code with no way to
    // trace it back to a real customer.
    var transportOnly = Array.from(transportSet).filter(function (o) { return !appSet.has(o); }).sort().map(function (o) {
      var t = transportDetailById[o] || {};
      return { order_id: o, customer: t.customer || '', location_name: t.location_name || '', address: t.address || '' };
    });
    var matched = Array.from(appSet).filter(function (o) { return transportSet.has(o); });

    res.json({
      has_data: true,
      date: date,
      app_total: appSet.size,
      transport_total: transportSet.size,
      matched_count: matched.length,
      app_only: appOnly.sort(),
      transport_only: transportOnly,
      is_exception_day: isExceptionDay,
      source_breakdown: sourceBreakdownRes.rows,
      org_breakdown: orgBreakdownRes.rows
    });
  } catch (e) {
    console.error('order-diff error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/dispatch/reconcile/order-diff/export', noCache, requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var date = req.query.date;
    if (!date) return res.status(400).json({ error: 'date is required' });
    var appRows = await pool.query("SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1 AND (org = ANY($2::text[]) OR (org = 'DCF' AND UPPER(COALESCE(city,'')) <> 'DUBAI'))", [date, RECON_ELIGIBLE_ORG_LIST]);
    var transportRows = await pool.query('SELECT order_id, customer, location_name, address FROM transport_order_ids WHERE date_key=$1', [date]);
    var appSet = new Set(appRows.rows.map(function (r) { return r.order_code; }));
    var transportDetailById = {};
    transportRows.rows.forEach(function (r) { transportDetailById[r.order_id] = r; });
    var transportSet = new Set(transportRows.rows.map(function (r) { return r.order_id; }));
    var appOnly = Array.from(appSet).filter(function (o) { return !transportSet.has(o); }).sort();
    var transportOnly = Array.from(transportSet).filter(function (o) { return !appSet.has(o); }).sort();

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();
    var ws = wb.addWorksheet('Order Diff ' + date);
    ws.columns = [{ width: 30 }, { width: 30 }, { width: 34 }, { width: 34 }];
    var titleRow = ws.addRow(['ORDER-LEVEL RECONCILIATION — ' + date]);
    ws.mergeCells('A' + titleRow.number + ':D' + titleRow.number);
    titleRow.font = { bold: true, size: 14 };
    ws.addRow(['App orders: ' + appSet.size, 'Transport orders: ' + transportSet.size]);
    ws.addRow([]);
    var hdr = ws.addRow(['Dispatched by AKI, missing from Transport file', '', 'In Transport file, missing from AKI dispatch (order code)', 'Customer / Location']);
    hdr.font = { bold: true };
    var maxLen = Math.max(appOnly.length, transportOnly.length);
    for (var i = 0; i < maxLen; i++) {
      var t = transportOnly[i] ? (transportDetailById[transportOnly[i]] || {}) : {};
      var detail = [t.customer, t.location_name].filter(Boolean).join(' — ');
      ws.addRow([appOnly[i] || '', '', transportOnly[i] || '', detail]);
    }
    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Order_Diff_' + date + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
});

app.get('/api/dispatch/reconcile/export', noCache, requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var params = [], clauses = [];
    if (req.query.date_from) { params.push(req.query.date_from); clauses.push('date_key >= $' + params.length); }
    if (req.query.date_to) { params.push(req.query.date_to); clauses.push('date_key <= $' + params.length); }
    var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
    var transportRows = await pool.query(`SELECT date_key::text, drop_class, reported_drops, completed_drops FROM transport_drop_reconciliation ${where} ORDER BY date_key`, params);
    var byDate = {};
    transportRows.rows.forEach(function (r) {
      if (!byDate[r.date_key]) byDate[r.date_key] = { total: 0, completed: 0, classes: {} };
      byDate[r.date_key].total += r.reported_drops;
      byDate[r.date_key].completed += r.completed_drops;
      byDate[r.date_key].classes[r.drop_class] = r.reported_drops;
    });
    var dateKeys = Object.keys(byDate).sort();
    var appRows = await pool.query(`SELECT date_key::text, summary FROM dispatch_data WHERE date_key = ANY($1::date[])`, [dateKeys]);
    var appByDate = {};
    appRows.rows.forEach(function (r) { appByDate[r.date_key] = r.summary; });

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();
    var ws = wb.addWorksheet('Reconciliation');
    ws.columns = [{ width: 14 }, { width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }];
    var titleRow = ws.addRow(['TRANSPORT DROP RECONCILIATION — App Captured vs Transport Reported (every dispatched drop, any status)']);
    ws.mergeCells('A' + titleRow.number + ':H' + titleRow.number);
    titleRow.font = { bold: true, size: 14 };
    ws.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
    ws.addRow([]);
    var hdr = ws.addRow(['Date', 'App Captured', 'Transport Bulk', 'Transport Multi', 'Transport Total', 'Variance', 'Variance %', 'Completed So Far']);
    hdr.font = { bold: true };
    var totA = 0, totT = 0, totTBothSides = 0, daysBothSides = 0;
    dateKeys.forEach(function (d) {
      var appSummary = appByDate[d];
      var appDrops = appSummary && appSummary.total_drops_reconcile !== undefined ? appSummary.total_drops_reconcile : (appSummary && appSummary.total_drops !== undefined ? appSummary.total_drops : null);
      var t = byDate[d];
      var variance = appDrops !== null ? appDrops - t.total : null;
      var variancePct = (appDrops !== null && t.total > 0) ? +((variance / t.total) * 100).toFixed(1) : null;
      var completedPct = t.total > 0 ? Math.round((t.completed / t.total) * 100) : 0;
      totT += t.total;
      if (appDrops !== null) { totA += appDrops; totTBothSides += t.total; daysBothSides++; }
      ws.addRow([d, appDrops === null ? 'No app data' : appDrops, t.classes.Bulk || 0, t.classes.Multi || 0, t.total, variance === null ? '' : variance, variancePct === null ? '' : variancePct, completedPct + '% (' + t.completed + ' of ' + t.total + ')']);
    });
    ws.addRow([]);
    // A straight sum across ALL dates is misleading whenever some dates are missing an
    // app-side upload (their Transport total would count with nothing to offset it) — so
    // the meaningful total only covers dates where BOTH sides actually have data.
    var totalRow = ws.addRow(['TOTAL (days with BOTH App + Transport data, ' + daysBothSides + ' of ' + dateKeys.length + ' days)', totA, '', '', totTBothSides, totA - totTBothSides, '']);
    totalRow.font = { bold: true };
    var allDatesRow = ws.addRow(['All ' + dateKeys.length + ' days — Transport total only (App data incomplete)', '', '', '', totT, '', '']);
    allDatesRow.font = { italic: true, color: { argb: 'FF8FA0B0' } };

    // ---- Sheet 2: Order-Level Issues — ready to send straight to the transport team ----
    var orderIdRows = await pool.query('SELECT DISTINCT date_key::text FROM transport_order_ids ORDER BY date_key');
    var wsIssues = wb.addWorksheet('Order-Level Issues');
    wsIssues.columns = [{ width: 14 }, { width: 26 }, { width: 40 }, { width: 40 }];
    var issueTitle = wsIssues.addRow(['ORDER-LEVEL ISSUES — orders that don\'t match between AKI dispatch and the transport team\'s file']);
    wsIssues.mergeCells('A' + issueTitle.number + ':D' + issueTitle.number);
    issueTitle.font = { bold: true, size: 14 };
    wsIssues.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
    wsIssues.addRow([]);
    var issueHdr = wsIssues.addRow(['Date', 'Order Code', 'Issue', 'Customer / Location (if in transport file)']);
    issueHdr.font = { bold: true };
    var totalIssueRows = 0, datesSkippedNoAppData = [];
    for (var i = 0; i < orderIdRows.rows.length; i++) {
      var dk = orderIdRows.rows[i].date_key;
      var appOrderRows = await pool.query("SELECT DISTINCT order_code FROM order_tracking WHERE date_key=$1 AND (org = ANY($2::text[]) OR (org = 'DCF' AND UPPER(COALESCE(city,'')) <> 'DUBAI'))", [dk, RECON_ELIGIBLE_ORG_LIST]);
      if (!appOrderRows.rows.length) { datesSkippedNoAppData.push(dk); continue; } // can't diff a date with no app dispatch upload at all
      var transportOrderRows = await pool.query('SELECT order_id, customer, location_name, address FROM transport_order_ids WHERE date_key=$1', [dk]);
      var appSet = new Set(appOrderRows.rows.map(function (r) { return r.order_code; }));
      var transportDetailById = {};
      transportOrderRows.rows.forEach(function (r) { transportDetailById[r.order_id] = r; });
      var transportSet = new Set(transportOrderRows.rows.map(function (r) { return r.order_id; }));
      var appOnly = Array.from(appSet).filter(function (o) { return !transportSet.has(o); }).sort();
      var transportOnly = Array.from(transportSet).filter(function (o) { return !appSet.has(o); }).sort();
      appOnly.forEach(function (o) {
        var row = wsIssues.addRow([dk, o, 'Dispatched by AKI — missing from transport file', '']);
        row.getCell(3).font = { color: { argb: 'FFE05C5C' } };
        totalIssueRows++;
      });
      transportOnly.forEach(function (o) {
        var t = transportDetailById[o] || {};
        var detail = [t.customer, t.location_name].filter(Boolean).join(' — ');
        var row = wsIssues.addRow([dk, o, 'In transport file — missing from AKI dispatch (ask why this was added)', detail]);
        row.getCell(3).font = { color: { argb: 'FFB8860B' }, bold: true };
        totalIssueRows++;
      });
    }
    if (!totalIssueRows) {
      wsIssues.addRow(['No order-level issues found for any date with data on both sides.']);
    }
    if (datesSkippedNoAppData.length) {
      wsIssues.addRow([]);
      wsIssues.addRow(['Dates skipped (no App dispatch upload for that date, so no order-level comparison possible):']);
      wsIssues.addRow([datesSkippedNoAppData.join(', ')]);
    }

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Transport_Reconciliation_' + Date.now() + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
});

// ── Fleet Mix (Own vs Hired Drivers), day-by-day — powers the small table
// under "Drops by City" on Daily Dispatch, and its own separate Excel export.
// Pulls straight from dispatch_data.summary.driver_source_split, which is
// already computed and saved on every dispatch upload — nothing new to store.
async function getFleetSummaryRows(dateFrom, dateTo) {
  var params = [], clauses = [];
  if (dateFrom) { params.push(dateFrom); clauses.push('date_key >= $' + params.length); }
  if (dateTo) { params.push(dateTo); clauses.push('date_key <= $' + params.length); }
  var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
  var r = await pool.query(`SELECT date_key::text, summary FROM dispatch_data ${where} ORDER BY date_key`, params);
  return r.rows.map(function(row) {
    var s = row.summary || {};
    var dss = s.driver_source_split || {};
    var inhouse = dss.inhouse || {};
    var hired = dss.hired || {};
    var totalDrivers = (inhouse.driver_count || 0) + (hired.driver_count || 0);
    var hiredPct = totalDrivers > 0 ? +(((hired.driver_count || 0) / totalDrivers) * 100).toFixed(1) : 0;
    var totalDrops = (inhouse.drops || 0) + (hired.drops || 0);
    var hiredDropsPct = totalDrops > 0 ? +(((hired.drops || 0) / totalDrops) * 100).toFixed(1) : 0;
    return {
      date: row.date_key,
      inhouse_drivers: inhouse.driver_count || 0,
      hired_drivers: hired.driver_count || 0,
      hired_pct: hiredPct,
      inhouse_drops: inhouse.drops || 0,
      hired_drops: hired.drops || 0,
      hired_drops_pct: hiredDropsPct,
      total_orders: s.total_orders || 0,
      total_value: s.total_value || 0,
      hired_value: hired.value || 0
    };
  });
}
app.get('/api/dispatch/fleet-summary', requireAuth, async function (req, res) {
  try {
    var rows = await getFleetSummaryRows(req.query.date_from, req.query.date_to);
    res.json({ rows: rows });
  } catch (e) { res.status(500).json({ error: e.message }); }
});
app.get('/api/dispatch/fleet-summary/export', noCache, requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var rows = await getFleetSummaryRows(req.query.date_from, req.query.date_to);
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();
    var ws = wb.addWorksheet('Fleet Mix');
    ws.columns = [{ width: 14 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 14 }, { width: 16 }];
    var titleRow = ws.addRow(['FLEET MIX — OWN-FLEET vs HIRED DRIVERS, DAY BY DAY']);
    ws.mergeCells('A' + titleRow.number + ':I' + titleRow.number);
    titleRow.font = { bold: true, size: 14 };
    ws.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
    ws.addRow(['Range', (req.query.date_from || 'earliest') + ' to ' + (req.query.date_to || 'latest')]);
    ws.addRow([]);
    var hdr = ws.addRow(['Date', 'In-House Drivers', 'Hired Drivers', '% Hired (Drivers)', 'In-House Drops', 'Hired Drops', '% Hired (Drops)', 'Total Orders', 'Total Value (AED)']);
    hdr.font = { bold: true };
    var totInhouse = 0, totHired = 0, totInhouseDrops = 0, totHiredDrops = 0, totOrders = 0, totValue = 0;
    rows.forEach(function (r) {
      var row = ws.addRow([r.date, r.inhouse_drivers, r.hired_drivers, r.hired_pct, r.inhouse_drops, r.hired_drops, r.hired_drops_pct, r.total_orders, r.total_value]);
      row.getCell(4).numFmt = '0.0"%"';
      row.getCell(7).numFmt = '0.0"%"';
      row.getCell(9).numFmt = '#,##0';
      if (r.hired_pct >= 25) row.getCell(4).font = { color: { argb: 'FFE05C5C' }, bold: true };
      totInhouse += r.inhouse_drivers; totHired += r.hired_drivers; totInhouseDrops += r.inhouse_drops; totHiredDrops += r.hired_drops; totOrders += r.total_orders; totValue += r.total_value;
    });
    ws.addRow([]);
    var avgPct = rows.length ? +((rows.reduce(function(s,r){return s+r.hired_pct;},0) / rows.length).toFixed(1)) : 0;
    var avgDropsPct = rows.length ? +((rows.reduce(function(s,r){return s+r.hired_drops_pct;},0) / rows.length).toFixed(1)) : 0;
    var totalRow = ws.addRow(['TOTAL / AVERAGE (' + rows.length + ' days)', totInhouse, totHired, avgPct, totInhouseDrops, totHiredDrops, avgDropsPct, totOrders, totValue]);
    totalRow.font = { bold: true };
    totalRow.getCell(4).numFmt = '0.0"%"';
    totalRow.getCell(7).numFmt = '0.0"%"';
    totalRow.getCell(9).numFmt = '#,##0';

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Fleet_Mix_' + Date.now() + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
});
app.get('/api/dispatch/fleet-summary/analyze', requireAuth, async function (req, res) {
  try {
    var rows = await getFleetSummaryRows(req.query.date_from, req.query.date_to);
    if (!rows.length) return res.status(400).json({ error: 'No data in this range to analyze.' });

    var totInhouse = rows.reduce(function(s,r){return s+r.inhouse_drivers;},0);
    var totHired = rows.reduce(function(s,r){return s+r.hired_drivers;},0);
    var totOrders = rows.reduce(function(s,r){return s+r.total_orders;},0);
    var totValue = rows.reduce(function(s,r){return s+r.total_value;},0);
    var avgPct = +((rows.reduce(function(s,r){return s+r.hired_pct;},0) / rows.length).toFixed(1));
    var worstDay = rows.slice().sort(function(a,b){return b.hired_pct-a.hired_pct;})[0];

    var dataText = 'Date | In-House Drivers | Hired Drivers | %Hired(Drivers) | In-House Drops | Hired Drops | %Hired(Drops) | Total Orders | Total Value(AED)\n' +
      rows.map(function(r){ return r.date+' | '+r.inhouse_drivers+' | '+r.hired_drivers+' | '+r.hired_pct+'% | '+r.inhouse_drops+' | '+r.hired_drops+' | '+r.hired_drops_pct+'% | '+r.total_orders+' | '+r.total_value; }).join('\n');

    var prompt = 'You are a senior logistics operations analyst reviewing a UAE last-mile delivery fleet\'s daily own-fleet-vs-hired-driver mix. ' +
      'Range: ' + rows.length + ' days (' + rows[0].date + ' to ' + rows[rows.length-1].date + '). ' +
      'Totals: ' + totOrders.toLocaleString() + ' orders, AED ' + totValue.toLocaleString() + ' value, average ' + avgPct + '% of drivers hired. ' +
      'Worst single day for hired reliance: ' + worstDay.date + ' at ' + worstDay.hired_pct + '% hired drivers (' + worstDay.hired_drivers + ' hired vs ' + worstDay.inhouse_drivers + ' in-house), ' + worstDay.total_orders + ' orders that day.\n\n' +
      'Raw daily data:\n' + dataText + '\n\n' +
      'Do a deep analysis and answer, in clear sections with short headers (no markdown tables, plain text/bullets only):\n' +
      '1) WHERE ARE THE GAPS — which specific date(s) show the fleet leaning most heavily on hired drivers, and is that driven by order volume spikes, a drop in in-house driver availability, or something else visible in the numbers?\n' +
      '2) WHY THIS IS HAPPENING — the most likely root cause(s) based on the pattern across days (e.g. weekday vs weekend, volume correlation, a specific date standing out as an outlier).\n' +
      '3) HOW TO FIX IT — concrete, practical recommendations to reduce hired-driver dependency going forward (e.g. staffing/roster changes, route consolidation, forecasting a known volume spike) — 3 to 5 specific, actionable points.\n' +
      'Be direct and specific with dates and numbers from the data — don\'t give generic advice that could apply to any dataset.';

    var msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001', max_tokens: 1400,
      system: 'You are AZHAR-AI\'s logistics analyst for a UAE distribution company. Be concise, specific, and cite actual dates/numbers from the data given. No fluff, no generic filler.',
      messages: [{ role: 'user', content: prompt }]
    });
    res.json({ analysis: msg.content[0].text });
  } catch (e) { res.status(500).json({ error: 'Analysis failed: ' + e.message }); }
});

// ============================================================
// DRIVER ORDER DETAILS — the real cost driver here is that EACH DROP IS BILLED
// SEPARATELY (per the transport team's rate card): if the same customer/location is
// delivered to by 4 different trucks on the same day, that's 4 drop charges for one
// physical stop. A single truck carrying a lot of value into ONE drop is NOT a
// problem — that's one truck, one stop, one charge, working exactly as it should.
// So this report is built around DUPLICATE DROPS: locations that were hit by more
// than one truck/driver on the same day, reusing the same repeat-visit detection
// already computed for the Drop Analysis panel above (which separates a genuinely
// required split — e.g. Food + Non-Food needing two trucks — from an avoidable
// duplicate — the same order type sent twice). A driver/truck is a DAILY
// assignment (especially a hired one, identified only by phone number), so this
// always reflects a SINGLE day — the most recent date in whatever range is picked
// — never a multi-day blend.
// ============================================================
async function getDriverDetailSnapshot(dateFrom, dateTo) {
  var params = [], clauses = [];
  if (dateFrom) { params.push(dateFrom); clauses.push('date_key >= $' + params.length); }
  if (dateTo) { params.push(dateTo); clauses.push('date_key <= $' + params.length); }
  var where = clauses.length ? ('WHERE ' + clauses.join(' AND ')) : '';
  // Only the single most recent date in the picked range — see note above.
  var r = await pool.query(`SELECT date_key::text, summary FROM dispatch_data ${where} ORDER BY date_key DESC LIMIT 1`, params);
  if (!r.rows.length) return null;
  var s = r.rows[0].summary || {};
  var dss = s.driver_source_split || {};
  return {
    date: r.rows[0].date_key,
    hired: dss.hired_driver_details || [],
    inhouse: dss.inhouse_driver_details || [],
    repeatLocations: s.repeat_locations || [],
    perDropRate: (s.truck_cost_estimate && s.truck_cost_estimate.available && s.truck_cost_estimate.total_drops_billed)
      ? (s.truck_cost_estimate.total_estimated_cost / s.truck_cost_estimate.total_drops_billed) : null
  };
}
// One day's hired + in-house drivers, flattened into a per-order list (unaffected by
// the fix below — this is still the correct "which customer, order value, final
// drop, per driver" view). The duplicate-drop view is built separately from
// repeat_locations, which is already keyed by physical location, not by driver.
function buildDriverOrderRows(snapshot) {
  var orderRows = [];
  ['inhouse', 'hired'].forEach(function(kind) {
    (snapshot[kind] || []).forEach(function(drv) {
      var type = kind === 'hired' ? 'Hired Truck' : 'In-House';
      (drv.order_details || []).forEach(function(o) {
        orderRows.push({
          date: snapshot.date, driver: drv.name, type: type,
          customer: o.customer, value: o.value, drop: o.drop, address: o.address, order_code: o.order_code, order_type: o.type
        });
      });
    });
  });
  return orderRows;
}
// Avoidable duplicate-drop locations only — same order type, multiple trucks, same
// day. Excludes legitimate splits (genuinely different order types needing separate
// trucks) since those aren't a cost problem to fix. Each row names the driver(s)
// involved and the extra drop charges that stop is paying.
function buildDuplicateDropRows(snapshot) {
  return (snapshot.repeatLocations || [])
    .filter(function(l) { return !l.is_legitimate_split; })
    .map(function(l) {
      var drivers = [];
      (l.route_types || []).forEach(function(rt) { (rt.drivers || []).forEach(function(d) { if (drivers.indexOf(d) === -1) drivers.push(d); }); });
      var extraDrops = Math.max(0, l.visit_count - 1);
      return {
        customer: l.customer || '(no name)', address: l.address || '', location_id: l.location_id,
        drivers: drivers.length ? drivers.join(' | ') : '(no driver on file)',
        visit_count: l.visit_count, extra_drops: extraDrops, total_value: l.total_value,
        status: l.is_high_value_exception ? 'Review (high value)' : 'Avoidable',
        extra_cost: snapshot.perDropRate != null ? Math.round(extraDrops * snapshot.perDropRate) : null
      };
    })
    .sort(function(a, b) { return b.extra_drops - a.extra_drops || b.total_value - a.total_value; });
}
// The AI is asked for plain, short text — but strip any markdown that slips through
// anyway (#, ##, ** **) so it never shows up as literal symbols in an Excel cell.
function stripMarkdown(text) {
  return String(text || '')
    .replace(/^#{1,6}\s*/gm, '')
    .replace(/\*\*(.*?)\*\*/g, '$1')
    .replace(/^-{3,}$/gm, '');
}
app.get('/api/dispatch/driver-details/export', noCache, requireAuth, async function (req, res) {
  try {
    var ExcelJS = require('exceljs');
    var snapshot = await getDriverDetailSnapshot(req.query.date_from, req.query.date_to);
    if (!snapshot) return res.status(400).json({ error: 'No dispatch data in this range.' });
    var orderRows = buildDriverOrderRows(snapshot);
    var dupRows = buildDuplicateDropRows(snapshot);
    var totalExtraDrops = dupRows.reduce(function(s, d) { return s + d.extra_drops; }, 0);
    var totalExtraCost = snapshot.perDropRate != null ? Math.round(totalExtraDrops * snapshot.perDropRate) : null;

    var aiAdvice = '';
    try { aiAdvice = await generateDuplicateDropAdvice(snapshot.date, dupRows, totalExtraDrops, totalExtraCost); }
    catch (e) { aiAdvice = 'AI advice unavailable: ' + e.message; }

    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();

    // ---- Sheet 1: Executive Summary ----
    var es = wb.addWorksheet('Executive Summary');
    es.columns = [{ width: 34 }, { width: 22 }];
    var t1 = es.addRow(['DRIVER ORDER DETAILS — EXECUTIVE SUMMARY']); es.mergeCells('A' + t1.number + ':B' + t1.number); t1.font = { bold: true, size: 14 };
    es.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
    es.addRow(['Date', snapshot.date]);
    es.addRow([]);
    var noteRow = es.addRow(['One drop = one charge. Sending the same location on 4 separate trucks pays the drop cost 4 times — that duplicate cost is what this sheet tracks, not how much value a single truck carries.']);
    es.mergeCells('A' + noteRow.number + ':B' + noteRow.number); noteRow.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
    es.addRow([]);
    var h1 = es.addRow(['Metric', 'Value']); h1.font = { bold: true };
    es.addRow(['Locations Hit by 2+ Trucks Today (Avoidable)', dupRows.length]);
    es.addRow(['Extra Drops Paid For (Avoidable)', totalExtraDrops]);
    var extraCostRow = es.addRow(['Estimated Extra Cost (AED)', totalExtraCost != null ? totalExtraCost : 'n/a — no truck-type rate data today']);
    if (totalExtraCost != null && totalExtraCost > 0) extraCostRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } };
    es.addRow([]);
    var h2 = es.addRow(['Duplicate-Drop Locations (most extra drops first)']); es.mergeCells('A' + h2.number + ':B' + h2.number); h2.font = { bold: true };
    var h2b = es.addRow(['Customer / Location', 'Extra Drops']); h2b.font = { bold: true };
    dupRows.slice(0, 10).forEach(function(d) {
      var row = es.addRow([d.customer + ' — ' + d.drivers, d.extra_drops]);
      if (d.extra_drops > 0) row.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF8D7DA' } };
    });

    // ---- Sheet 2: AI Advice ----
    var aiSheet = wb.addWorksheet('AI Advice');
    aiSheet.columns = [{ width: 100 }];
    var aiTitle = aiSheet.addRow(['AI ADVICE — REDUCING DUPLICATE DROP COST (' + snapshot.date + ')']); aiTitle.font = { bold: true, size: 14 };
    aiSheet.addRow([]);
    stripMarkdown(aiAdvice).split('\n').forEach(function(line) { aiSheet.addRow([line]); });
    aiSheet.getColumn(1).alignment = { wrapText: true, vertical: 'top' };

    // ---- Sheet 3: Duplicate Drop Locations ----
    var dd = wb.addWorksheet('Duplicate Drops');
    dd.columns = [{ width: 30 }, { width: 34 }, { width: 34 }, { width: 10 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 18 }];
    var ddTitle = dd.addRow(['DUPLICATE DROP LOCATIONS — ' + snapshot.date]); dd.mergeCells('A' + ddTitle.number + ':H' + ddTitle.number); ddTitle.font = { bold: true, size: 13 };
    dd.addRow([]);
    var ddHdr = dd.addRow(['Customer', 'Address', 'Drivers / Trucks Involved', 'Visits', 'Extra Drops', 'Total Value (AED)', 'Extra Cost (AED)', 'Status']); ddHdr.font = { bold: true };
    dupRows.forEach(function(d) {
      var row = dd.addRow([d.customer, d.address, d.drivers, d.visit_count, d.extra_drops, d.total_value, d.extra_cost != null ? d.extra_cost : 'n/a', d.status]);
      row.getCell(6).numFmt = '#,##0'; if (d.extra_cost != null) row.getCell(7).numFmt = '#,##0';
      var fillColor = d.status === 'Avoidable' ? 'FFF8D7DA' : 'FFFDF2CE';
      row.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
    });
    if (!dupRows.length) dd.addRow(['No duplicate-drop locations found for ' + snapshot.date + ' — every location was served by a single truck.']);

    // ---- Sheet 4: Order Details (customer / value / final drop, per order) ----
    var od = wb.addWorksheet('Order Details');
    od.columns = [{ width: 12 }, { width: 24 }, { width: 12 }, { width: 30 }, { width: 14 }, { width: 24 }, { width: 30 }, { width: 16 }, { width: 12 }];
    var odTitle = od.addRow(['ORDER DETAILS — ' + snapshot.date]); od.mergeCells('A' + odTitle.number + ':I' + odTitle.number); odTitle.font = { bold: true, size: 13 };
    od.addRow([]);
    var odHdr = od.addRow(['Date', 'Driver / Truck', 'Type', 'Customer', 'Order Value (AED)', 'Final Drop', 'Address', 'Order Code', 'Order Type']); odHdr.font = { bold: true };
    orderRows.forEach(function(o) {
      var row = od.addRow([o.date, o.driver, o.type, o.customer, o.value, o.drop, o.address, o.order_code, o.order_type]);
      row.getCell(5).numFmt = '#,##0';
      row.getCell(3).font = { bold: true, color: { argb: o.type === 'Hired Truck' ? 'FFC0392B' : 'FF1E7E34' } };
    });
    od.autoFilter = { from: { row: odHdr.number, column: 1 }, to: { row: odHdr.number, column: 9 } };

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Driver_Order_Details_' + snapshot.date + '.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
});
// Kept short and plain on purpose — 3 tight sections, no markdown symbols, no
// "worst"/ranking language — just the facts: which locations got double-billed and
// how to stop it. "Hired" trucks are a same-day engagement (phone number only, no
// name on file), not permanent staff, so the fix is a dispatch-planning one, not a
// performance judgment on any one driver.
async function generateDuplicateDropAdvice(date, dupRows, totalExtraDrops, totalExtraCost) {
  if (!dupRows.length) {
    return 'No duplicate-drop locations on ' + date + '. Every customer/location was served by exactly one truck today — no avoidable extra drop charges to report.';
  }
  var dataText = dupRows.slice(0, 10).map(function(d) {
    return d.customer + (d.address ? ' (' + d.address + ')' : '') + ': ' + d.visit_count + ' trucks — ' + d.drivers + ' — ' + d.extra_drops + ' extra drop(s), AED ' + d.total_value.toLocaleString() + ' total value' + (d.extra_cost != null ? ', ~AED ' + d.extra_cost.toLocaleString() + ' avoidable cost' : '');
  }).join('\n');
  var prompt = 'UAE last-mile distribution fleet, single day: ' + date + '. Each drop is billed separately by the transport team — sending the same customer/location on multiple trucks pays the drop charge multiple times for one physical stop.\n\n' +
    'Today: ' + dupRows.length + ' location(s) were visited by more than one truck for the same order type (not a required Food/Non-Food or Frozen/Ambient split — those need separate trucks and are excluded already). Total avoidable extra drops: ' + totalExtraDrops + (totalExtraCost != null ? ', roughly AED ' + totalExtraCost.toLocaleString() + ' in avoidable drop charges.' : '.') + '\n\n' +
    'The locations, with the driver(s)/truck(s) that each visited:\n' + dataText + '\n\n' +
    'Reply in PLAIN TEXT only — no #, no **, no markdown, no words like "worst". Keep it SHORT: 3 short sections, each 1-2 sentences plus at most 2 bullet lines, using the actual customer names and driver/truck IDs above:\n' +
    'WHAT HAPPENED: which location(s) got hit by more than one truck today, and by how much.\n' +
    'LIKELY CAUSE: the most probable dispatch-planning reason (orders for the same stop assigned to different trucks/routes instead of being grouped).\n' +
    'FIX: 2-3 concrete dispatch-planning steps to group same-type orders to one location onto a single truck before assigning routes.';
  var msg = await anthropic.messages.create({
    model: 'claude-haiku-4-5-20251001', max_tokens: 500,
    system: 'You are AZHAR-AI\'s logistics analyst for a UAE distribution company. Be brief and concrete — plain text only, no markdown symbols, no "worst"/ranking language, cite real names/numbers, no filler.',
    messages: [{ role: 'user', content: prompt }]
  });
  return msg.content[0].text;
}
app.get('/api/dispatch/driver-details/analyze', requireAuth, async function (req, res) {
  try {
    var snapshot = await getDriverDetailSnapshot(req.query.date_from, req.query.date_to);
    if (!snapshot) return res.status(400).json({ error: 'No dispatch data in this range to analyze.' });
    var dupRows = buildDuplicateDropRows(snapshot);
    var totalExtraDrops = dupRows.reduce(function(s, d) { return s + d.extra_drops; }, 0);
    var totalExtraCost = snapshot.perDropRate != null ? Math.round(totalExtraDrops * snapshot.perDropRate) : null;
    var advice = await generateDuplicateDropAdvice(snapshot.date, dupRows, totalExtraDrops, totalExtraCost);
    res.json({ analysis: stripMarkdown(advice), date: snapshot.date });
  } catch (e) { res.status(500).json({ error: 'Analysis failed: ' + e.message }); }
});


app.post('/api/dispatch/ask', function(req, res) {
  try {
    if (!currentDispatch) return res.json({ result:'No dispatch data. Please upload first.' });
    var s = currentDispatch.summary;
    var context = 'Date: '+currentDispatch.date+'\nTotal Orders: '+s.total_orders+'\nTotal Value: AED '+s.total_value+'\nFood: '+s.food_orders+' orders AED '+s.food_value+'\nNon-Food: '+s.non_food_orders+' orders AED '+s.non_food_value+'\n3PL: '+s.pl_orders+'\n\nCSV:\n'+(currentDispatch.csvText||'').substring(0,8000);
    anthropic.messages.create({
      model:'claude-haiku-4-5-20251001', max_tokens:1500,
      messages:[{ role:'user', content:'You are AZHAR-AI Dispatch Intelligence for UAE logistics.\n\n'+context+'\n\nQuestion: '+req.body.question+'\n\nAnswer with exact numbers. Use AED for currency.' }]
    }).then(function(msg) { res.json({ result: msg.content[0].text }); })
      .catch(function(e) { res.status(500).json({ error: e.message }); });
  } catch(e) { if (!res.headersSent) res.status(500).json({ error: e.message }); }
});

//  REJECTION STORE (+ DB) 
var rejectionData = null;

async function loadRejectionFromDB() {
  try {
    var row = await dbLoadRejection();
    if (row) {
      // Safely parse orgs/months — may be object (JSONB) or string (old saves)
      var orgs = row.orgs;
      var months = row.months;
      if (typeof orgs === 'string') try { orgs = JSON.parse(orgs); } catch(e) { orgs = {}; }
      if (typeof months === 'string') try { months = JSON.parse(months); } catch(e) { months = {}; }
      // Version check: if 'all' org has no detail array, data is old — needs re-upload
      var hasDetail = orgs && orgs.all && Array.isArray(orgs.all.detail) && orgs.all.detail.length > 0;
      var needsReview = row.needs_review;
      if (typeof needsReview === 'string') try { needsReview = JSON.parse(needsReview); } catch(e) { needsReview = []; }
      var autoClassify = row.auto_classify;
      if (typeof autoClassify === 'string') try { autoClassify = JSON.parse(autoClassify); } catch(e) { autoClassify = null; }
      rejectionData = {
        uploadedAt: row.uploaded_at, uploadedBy: row.uploaded_by,
        fileName: row.file_name, totalOrders: row.total_orders,
        orgs: orgs, months: months,
        needsReview: needsReview || [], autoClassify: autoClassify || null,
        needsReupload: !hasDetail
      };
      console.log('Loaded rejection from DB. hasDetail:', hasDetail, 'orgs keys:', Object.keys(orgs||{}).length);
      return true;
    }
  } catch(e) { console.error('loadRejectionFromDB error:', e.message); }
  try {
    var saved = loadJSON(REJECTION_FILE);
    if (saved) { rejectionData = saved; console.log('Loaded rejection from file'); }
  } catch(e) { console.error('loadRejectionFromDB file fallback error:', e.message); }
  return false;
}
loadRejectionFromDB();

async function rejectionUploadHandler(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error:'No file received' });
    var ext = path.extname(req.file.originalname||'').toLowerCase();
    if (ext !== '.xlsx' && ext !== '.xls' && ext !== '.csv')
      return res.status(400).json({ error:'Please upload .xlsx, .xls or .csv' });

    console.log('Reading rejection file:', req.file.originalname, req.file.size, 'bytes');
    var rows = [];
    if (ext === '.csv') {
      // Try UTF-8 first, fall back to latin1 for special characters
      var csvText;
      try {
        csvText = req.file.buffer.toString('utf8');
        // Check for replacement characters indicating wrong encoding
        if (csvText.includes('\uFFFD')) {
          csvText = req.file.buffer.toString('latin1');
          console.log('CSV: switched to latin1 encoding');
        }
      } catch(e) {
        csvText = req.file.buffer.toString('latin1');
        console.log('CSV: using latin1 encoding');
      }
      var csvRows = csvText.split('\n').filter(function(l){return l.trim();});
      if (csvRows.length < 2) return res.status(400).json({ error:'CSV file is empty' });
      function parseCSVLine(line) {
        var result=[], cell='', inQ=false;
        for (var ci=0; ci<line.length; ci++) {
          var ch=line[ci];
          if(ch==='"'){inQ=!inQ;}
          else if(ch===','&&!inQ){result.push(cell.trim());cell='';}
          else{cell+=ch;}
        }
        result.push(cell.trim());
        return result;
      }
      var headers = parseCSVLine(csvRows[0]).map(function(h){return h.replace(/"/g,'').trim();});
      for (var ci=1; ci<csvRows.length; ci++) {
        if (!csvRows[ci].trim()) continue;
        var vals = parseCSVLine(csvRows[ci]);
        var rowObj = {};
        headers.forEach(function(h,hi){ rowObj[h] = (vals[hi]||'').replace(/"/g,'').trim(); });
        rows.push(rowObj);
      }
      console.log('CSV rows parsed:', rows.length);
    } else {
      var wb = XLSX.read(req.file.buffer, { type:'buffer', dense:true, cellDates:false, cellNF:false, cellHTML:false, cellFormula:false });
      var sheetName = findDataSheet(wb);
      var ws = wb.Sheets[sheetName];
      rows = XLSX.utils.sheet_to_json(ws, { defval:'', raw:true });
      console.log('Excel rows:', rows.length, 'sheet:', sheetName);
    }
    if (!rows.length) return res.status(400).json({ error:'No rows found' });

    var keys0 = Object.keys(rows[0]);
    function findC() {
      var names = Array.prototype.slice.call(arguments);
      return keys0.find(function(k) {
        return names.some(function(n) { return k.toUpperCase().includes(n.toUpperCase()); });
      }) || null;
    }
    var RC = {
      status:  findC('FINAL STATUS', 'STATUS'),
      org:     findC('ORGANIZATION') || findC('ORG-BU'),
      date:    findC('D DATE', 'DATE', 'DELIVERY DATE'),
      root:    findC('FINAL- ROOT', 'FINA- ROOT', 'ROOT CAUSE', 'ROOT_CAUSE', 'REASON-1'),
      cust:    findC('CUSTOMER NAME') || findC('CUSTOMER'),
      addr:    findC('FULL ADDRESS', 'ADDRESS MATCHING', 'ADDRESS'),
      area:    findC('AREA', 'CITY'),
      value:   findC('VALUE', 'AMOUNT'),
      type:    findC('TYPE'),
      orderType: findC('ORDER TYPE'),
      custGroup: findC('CUSTOMER GROUP'),
      source:  findC('REMAKE -3', 'REMAKE') || findC('INTERNAL/EXTERNAL'),
      orderNo: findC('ORDER NO', 'ORDER_NO', 'ORDERNO', 'SHIPMENT_ID')
    };
    // Capture this BEFORE the safe-fallback below makes RC.value always truthy —
    // needed to distinguish "file genuinely has no Value column" from "column
    // exists but this cell is blank" further down (zero-value classification).
    var FILE_HAS_VALUE_COLUMN = !!RC.value;

    // ── Wrong-file guard: a genuine rejection report always has a Root Cause /
    // reason column. A generic order or dispatch report never does. This is
    // exactly the mistake that once silently wiped real rejection history —
    // someone uploaded "Delivery Status Report...xlsx" into this slot, it had
    // no root-cause column, everything parsed as 0 rejections, and that got
    // published over real data with no warning. Block that upfront now. ──
    var forcePublish = toStr(req.body.force) === 'true';
    if (!RC.root && !forcePublish) {
      return res.status(409).json({
        error: '⚠️ This file doesn\u2019t look like a Rejection report \u2014 no Root Cause / Reason column was found (only saw: ' + keys0.slice(0,8).join(', ') + (keys0.length>8?', …':'') + '). Uploading it anyway will overwrite the real rejection data currently on the dashboard with this file\u2019s numbers. Double-check you picked the right file before proceeding.',
        wrongFileWarning: true, detectedColumns: keys0, requiresConfirmation: true
      });
    }
    // Guarantee every RC.xxx is a real (non-null) key — needed so historical
    // rows reconstructed from rejection_rows (see below) always read/write
    // the same property names, even on a future file that happens to be
    // missing one of these columns. Never changes behavior for the actual
    // uploaded file: a genuinely absent column still reads as undefined.
    ['status','org','date','root','cust','addr','area','value','type','source','orderNo'].forEach(function(k) {
      RC[k] = RC[k] || ('__' + k);
    });
    console.log('Rejection cols:', JSON.stringify(RC));

    function isRej(row) {
      var s1 = toStr(row[RC.status]).toUpperCase();
      var s2 = toStr(row['Status']||'').toUpperCase();
      return s1==='REJECTION'||s1==='REJECTED'||s2==='R/D'||s2==='HOLD'||s2==='RD';
    }
    function isDel(row) {
      var s1 = toStr(row[RC.status]).toUpperCase();
      var s2 = toStr(row['Status']||'').toUpperCase();
      return s1==='DELIVERED'||s1.includes('DELIVER')||s2.includes('DELIVER')||s2==='D';
    }
    function parseDate(v) {
      if (!v) return null;
      if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

      // Excel serial date number (e.g. 46000)
      if (typeof v === 'number') {
        try { var unix = Math.round((v - 25569) * 86400 * 1000); var dd = new Date(unix); if (!isNaN(dd.getTime())) return dd; } catch(e2) {}
        return null;
      }

      var s = String(v).trim();
      if (!s) return null;

      var MONTH_NAMES = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };

      // Format: "13-Apr-26" or "13-Apr-2026" or "13 Apr 2026" (day - month name - year)
      var m1 = s.match(/^(\d{1,2})[\s\-\/]([A-Za-z]{3,})[\s\-\/](\d{2,4})$/);
      if (m1) {
        var mon1 = MONTH_NAMES[m1[2].toLowerCase().substring(0,3)];
        if (mon1 !== undefined) {
          var yr1 = parseInt(m1[3], 10); if (yr1 < 100) yr1 += 2000;
          var dt1 = new Date(yr1, mon1, parseInt(m1[1], 10));
          if (!isNaN(dt1.getTime())) return dt1;
        }
      }

      // Format: "Apr-13-26" or "Apr 13 2026" (month name - day - year)
      var m2 = s.match(/^([A-Za-z]{3,})[\s\-\/](\d{1,2})[\s\-\/](\d{2,4})$/);
      if (m2) {
        var mon2 = MONTH_NAMES[m2[1].toLowerCase().substring(0,3)];
        if (mon2 !== undefined) {
          var yr2 = parseInt(m2[3], 10); if (yr2 < 100) yr2 += 2000;
          var dt2 = new Date(yr2, mon2, parseInt(m2[2], 10));
          if (!isNaN(dt2.getTime())) return dt2;
        }
      }

      // Format: "YYYY-MM-DD" (ISO)
      var m3 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (m3) {
        var dt3 = new Date(parseInt(m3[1],10), parseInt(m3[2],10)-1, parseInt(m3[3],10));
        if (!isNaN(dt3.getTime())) return dt3;
      }

      // Format: "M/D/YYYY" or "D/M/YYYY" (slash-separated, ambiguous — resolved below)
      var m4 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
      if (m4) {
        var a = parseInt(m4[1],10), b = parseInt(m4[2],10), yr4 = parseInt(m4[3],10);
        if (yr4 < 100) yr4 += 2000;
        var month4, day4;
        if (a > 12 && b <= 12) { day4 = a; month4 = b; }       // first number can't be a month -> D/M/Y
        else { month4 = a; day4 = b; }                         // default: M/D/Y (matches the real data seen)
        var dt4 = new Date(yr4, month4-1, day4);
        if (!isNaN(dt4.getTime())) return dt4;
      }

      // Format: "13.04.2026" (dot-separated, D.M.Y)
      var m5 = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{2,4})$/);
      if (m5) {
        var yr5 = parseInt(m5[3],10); if (yr5 < 100) yr5 += 2000;
        var dt5 = new Date(yr5, parseInt(m5[2],10)-1, parseInt(m5[1],10));
        if (!isNaN(dt5.getTime())) return dt5;
      }

      // Last resort: native parser (handles anything unusual we haven't explicitly covered)
      var d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    }

    // ── Compute the exact set of dates this file covers up front — needed
    // below both for the subadmin guardrails and (later, unchanged) for the
    // per-date replace in rejection_rows. ──
    var newDates = new Set();
    rows.forEach(function(r) {
      var d0 = parseDate(r[RC.date]);
      if (d0) newDates.add(d0.getFullYear() + '-' + String(d0.getMonth() + 1).padStart(2, '0') + '-' + String(d0.getDate()).padStart(2, '0'));
    });
    var newDatesArr = Array.from(newDates).sort();

    // ── Sub-admin upload guardrails. Superadmin is fully exempt: unlimited
    // uploads, any date, no cooldown. This is what would have caught the
    // Prakash incident — a subadmin re-uploading a file that overlaps dates
    // already on the dashboard, silently reverting manual root-cause
    // corrections made on those dates. ──
    if (req.user.role === 'subadmin') {
      try {
        var cooldown = await checkSubadminCooldown(req.user.uid, 'rejection');
        if (cooldown) {
          return res.status(429).json({
            error: '\u23f3 Sub-admins can upload twice every 10 hours \u2014 you have used both uploads, you cannot re-upload for ' + cooldown.hoursLeft + 'h ' + cooldown.minsLeft + 'm.',
            cooldownActive: true, hoursLeft: cooldown.hoursLeft, minsLeft: cooldown.minsLeft
          });
        }
      } catch (e) { console.error('Rejection: cooldown check failed:', e.message); }

      if (newDatesArr.length) {
        try {
          var maxRes = await pool.query('SELECT MAX(entry_date) AS max_date FROM rejection_rows');
          var maxDateRaw = maxRes.rows[0] && maxRes.rows[0].max_date;
          if (maxDateRaw) {
            var md = new Date(maxDateRaw);
            var maxDateStr = md.getFullYear() + '-' + String(md.getMonth() + 1).padStart(2, '0') + '-' + String(md.getDate()).padStart(2, '0');
            var blockedDates = newDatesArr.filter(function(ds) { return ds <= maxDateStr; });
            if (blockedDates.length) {
              var approvalReason = toStr(req.body.approvalReason);
              if (req.body.requestApproval === 'true' && approvalReason) {
                // Subadmin has explained why — park the file for the super admin
                // to review, instead of applying it or hard-blocking it forever.
                await submitUploadApproval({
                  userId: req.user.uid, username: req.user.username || req.user.full_name || 'Sub-admin',
                  endpoint: 'rejection', fileName: req.file.originalname, fileBuffer: req.file.buffer,
                  blockedDates: blockedDates, reason: approvalReason, meta: {}
                });
                await recordSubadminUpload(req.user.uid, 'rejection');
                auditLog(null, req.user.username, 'UPLOAD_APPROVAL_REQUESTED', 'Rejection YTD: ' + req.file.originalname + ' \u2014 dates ' + blockedDates.slice(0,5).join(', ') + (blockedDates.length>5?', \u2026':'') + ' \u2014 reason: ' + approvalReason, req.headers['x-forwarded-for'] || req.ip || '');
                return res.status(202).json({
                  success: false, submittedForApproval: true,
                  message: 'Sent to the super admin for approval \u2014 you\u2019ll see it reflected once reviewed.'
                });
              }
              return res.status(409).json({
                error: '\u26d4 This file includes date(s) already on the dashboard (data exists through ' + maxDateStr + '): ' + blockedDates.slice(0, 5).join(', ') + (blockedDates.length > 5 ? ', \u2026' : '') + '. Sub-admins can only upload NEW dates after ' + maxDateStr + ' \u2014 existing dates can\u2019t be overwritten directly. You can send this to the super admin for approval instead.',
                blockedDates: blockedDates, cutoffDate: maxDateStr, subadminDateBlock: true, requiresApprovalReason: true
              });
            }
          }
        } catch (e) { console.error('Rejection: date-cutoff check failed:', e.message); }
      }
    }

    var orgMap={}, monthMap={};
    var totalRej=0, totalDel=0, totalVal=0;

    // ── Pull authoritative order values from the Dispatch module (order_tracking,
    // populated by Dispatch uploads) instead of trusting whatever's in the
    // rejection file's own Value column — that column is what the user asked
    // to have "pulled from the dispatch report link" rather than typed/trusted
    // as-is. Falls back to the file's own Value column when an order isn't
    // found there (e.g. dispatch data for that date hasn't been uploaded yet). ──
    var orderTrackingValues = {};
    try {
      var otRes = await pool.query('SELECT order_code, value FROM order_tracking');
      otRes.rows.forEach(function(r) { if (r.order_code) orderTrackingValues[r.order_code] = parseFloat(r.value) || 0; });
      console.log('Rejection: loaded', Object.keys(orderTrackingValues).length, 'order values from Dispatch (order_tracking) for auto-linking');
    } catch (e) { console.error('Rejection: could not load order_tracking for value linking:', e.message); }

    var classifyStats = { rootTotal: 0, rootAutoMatched: 0, typeAutoMatched: 0, sourceAutoMatched: 0, valueFromDispatch: 0, zeroValueTransfer: 0 };
    var needsReview = []; // capped list of discrepancy notes — either "still needs a human to classify" or "auto-handled, shown so you can verify" (see .flag)

    // ── Per-date storage: whether transport sends a whole month or just one
    // day (e.g. "today's 28th"), figure out the EXACT set of dates THIS file
    // covers and clear only those dates in rejection_rows before inserting
    // fresh rows. Uploading just the 28th replaces only the 28th — every
    // other already-stored day, whether earlier in the same month or a
    // different month entirely, stays untouched. The full aggregate is
    // rebuilt afterward from ALL stored rows, so nothing is ever lost. ──
    try {
      if (newDatesArr.length) {
        await pool.query('DELETE FROM rejection_rows WHERE entry_date = ANY($1::date[])', [newDatesArr]);
      }
    } catch (e) { console.error('Rejection: could not clear rejection_rows for per-date storage:', e.message); }

    var rowsToStore = [];
    var seenOrderVals = {};

    for (var i=0; i<rows.length; i++) {
      var row=rows[i];
      var rej=isRej(row), del=isDel(row);
      if (!rej && !del) continue;
      var d=parseDate(row[RC.date]);
      var mo=d?d.getMonth()+1:null, day=d?d.getDate():null;
      var org=toStr(row[RC.org]).toUpperCase().replace('NON-FOOD','DGC');
      var rootRaw=toStr(row[RC.root]);
      var rootClassified=autoClassifyRootCause(rootRaw);
      var root=rootClassified.category;
      var cust=toStr(row[RC.cust]);
      var addr=RC.addr?toStr(row[RC.addr]):'';
      var area=toStr(row[RC.area]);
      var orderNo=RC.orderNo?toStr(row[RC.orderNo]):'';
      var trackedVal = (orderNo && orderTrackingValues.hasOwnProperty(orderNo)) ? orderTrackingValues[orderNo] : null;
      if (trackedVal !== null && rej) classifyStats.valueFromDispatch++;
      // hasValueSignal distinguishes "confirmed AED 0" from "we simply don't know the
      // value" — the raw transport file often has NO Value column at all, so without
      // this check, every order not found in Dispatch would silently default to 0 and
      // get wrongly tagged as an Internal Transfer.
      var hasValueSignal = (trackedVal !== null) || FILE_HAS_VALUE_COLUMN;
      var rawVal=(trackedVal !== null) ? trackedVal : (parseFloat(row[RC.value])||0);

      // ── Zero-value rows are internal stock transfers, not real customer
      // rejections — this OVERRIDES whatever the reason text said, since a
      // AED 0 order was never a genuine sale to begin with. Only fires when
      // the value is CONFIRMED zero, never when it's just unknown. ──
      var isZeroValueTransfer = rej && hasValueSignal && rawVal === 0;
      var forceInternalSource = false;
      if (isZeroValueTransfer) {
        root = 'Internal Transfer';
        forceInternalSource = true;
        classifyStats.zeroValueTransfer++;
        if (needsReview.length < 300) needsReview.push({ orderNo: orderNo || '', org: org, customer: cust, rawReason: rootRaw, flag: 'zero_value_transfer', note: 'AED 0 value — auto-classified as Internal Transfer, not a customer rejection' });
      } else if (rej && rootRaw) {
        classifyStats.rootTotal++;
        if (rootClassified.matched) classifyStats.rootAutoMatched++;
        else if (needsReview.length < 300) needsReview.push({ orderNo: orderNo || '', org: org, customer: cust, rawReason: rootRaw, flag: 'unmatched', note: 'New/ambiguous reason text — needs a category picked manually' });
      }

      var val=(rej&&orderNo&&seenOrderVals[orderNo])?0:rawVal;
      if(rej&&orderNo&&!seenOrderVals[orderNo])seenOrderVals[orderNo]=rawVal;
      var rawTypeStr=toStr(row[RC.type]||'').toUpperCase();
      var orderTypeVal=RC.orderType?toStr(row[RC.orderType]):'';
      var custGroupVal=RC.custGroup?toStr(row[RC.custGroup]):'';
      var orgBuResult=autoClassifyOrgBu(org, orderTypeVal, custGroupVal, rawTypeStr);
      var typeStr=rawTypeStr;
      var isFood, isNF;
      if (orgBuResult.matched) {
        isFood=orgBuResult.isFood; isNF=orgBuResult.isNonFood;
        if (rej) classifyStats.typeAutoMatched++;
      } else {
        isFood=typeStr==='FOOD'||typeStr.startsWith('FOOD,');
        isNF=typeStr.includes('NON FOOD')||typeStr.includes('NON-FOOD');
      }
      var rawSrcStr=toStr(row[RC.source]||'').toUpperCase();
      var srcStr=rawSrcStr;
      if (forceInternalSource) {
        srcStr='INTERNAL';
      } else if (!rawSrcStr && rej) {
        var ieResult=deriveInternalExternal(root);
        if (ieResult.matched) {
          srcStr=ieResult.binary.toUpperCase();
          classifyStats.sourceAutoMatched++;
        } else {
          // Genuinely unmatched root cause (raw fallback text, still needs a
          // human to pick a category) — default to Internal rather than
          // leaving this blank. A blank source silently drops the row from
          // BOTH the External and Internal breakdowns, so filtering by
          // either one no longer sums to the "All" total. Internal is the
          // conservative choice here: it doesn't guess the customer is at
          // fault for something not yet confidently classified. Gets
          // corrected automatically once the row's Root Cause is fixed via
          // the manual correction endpoint, which recomputes source too.
          srcStr='INTERNAL';
        }
      }
      if (d) {
        rowsToStore.push({
          month_key: mo ? (d.getFullYear() + '-' + String(mo).padStart(2,'0')) : null,
          entry_date: d, status: rej ? 'rej' : 'del', org: org, root_cause: root,
          root_cause_source: isZeroValueTransfer ? 'zero_value' : (rootClassified.matched ? 'auto' : 'unmatched'),
          customer_name: cust, address: addr, area: area, order_no: orderNo, value: rawVal,
          is_food: !!isFood, is_nonfood: !!isNF, source: srcStr, file_name: req.file.originalname
        });
      }
    }

    // ── Persist this upload's rows, then rebuild the FULL aggregate (this
    // month's new rows + every other month already in the DB) from scratch —
    // this is what keeps Jan-Jul intact when August is uploaded, and what
    // lets a manual correction (see /api/rejection/rows/correct below) use
    // the exact same rebuild path. ──
    try {
      if (rowsToStore.length) {
        var CHUNK = 500;
        for (var ci = 0; ci < rowsToStore.length; ci += CHUNK) {
          var chunk = rowsToStore.slice(ci, ci + CHUNK);
          var vals = [], phs = [], pi = 1;
          chunk.forEach(function(r) {
            var cols = [r.month_key, r.entry_date, r.status, r.org, r.root_cause, r.root_cause_source, r.customer_name, r.address, r.area, r.order_no, r.value, r.is_food, r.is_nonfood, r.source, r.file_name];
            var ph = cols.map(function() { return '$' + (pi++); });
            phs.push('(' + ph.join(',') + ')');
            vals.push.apply(vals, cols);
          });
          await pool.query(
            'INSERT INTO rejection_rows (month_key, entry_date, status, org, root_cause, root_cause_source, customer_name, address, area, order_no, value, is_food, is_nonfood, source, file_name) VALUES ' + phs.join(','),
            vals
          );
        }
        console.log('Rejection: persisted', rowsToStore.length, 'rows for date(s)', newDatesArr.join(', '));
      }
    } catch (e) { console.error('Rejection: could not persist rows to rejection_rows:', e.message); }

    var rebuilt = await rebuildRejectionAggregateFromDB();
    orgMap = rebuilt.orgMap; monthMap = rebuilt.monthMap;
    totalRej = rebuilt.totalRej; totalDel = rebuilt.totalDel; totalVal = rebuilt.totalVal;
    var orgsOut = rebuilt.orgsOut, monthsOut = rebuilt.monthsOut;

    // ── Safety net: if this upload would make the dashboard's total drop
    // drastically, don't silently publish it — that's exactly the signature
    // of older data missing from rejection_rows (e.g. before this per-date
    // storage system existed, or a stray delete), not a real data change.
    // The new rows ARE safely stored either way; this only gates whether we
    // overwrite what the dashboard currently shows. ──
    var previousTotal = (rejectionData && rejectionData.totalOrders) || 0;
    var newTotal = totalRej + totalDel;
    if (previousTotal > 500 && newTotal < previousTotal * 0.5 && !forcePublish) {
      return res.status(409).json({
        error: 'This upload would drop the dashboard total from ' + previousTotal.toLocaleString() + ' to ' + newTotal.toLocaleString() + ' orders \u2014 that usually means older months/days aren\u2019t in storage yet, not that they were actually removed. Your new rows ARE saved safely. Re-upload your full historical master file first to restore the missing dates, then re-try this upload \u2014 or if you\u2019re certain ' + newTotal.toLocaleString() + ' is correct, resubmit with force=true.',
        previousTotal: previousTotal, newTotal: newTotal, requiresConfirmation: true
      });
    }

    rejectionData={uploadedAt:new Date().toISOString(),uploadedBy:req.body.uploadedBy||'Admin',fileName:req.file.originalname,totalOrders:totalRej+totalDel,orgs:orgsOut,months:monthsOut,needsReview:needsReview,autoClassify:classifyStats};

    // Save to DB + file
    var dbOk = await dbSaveRejection(rejectionData.uploadedBy, rejectionData.fileName, rejectionData.totalOrders, orgsOut, monthsOut, needsReview, classifyStats);
    saveJSON(REJECTION_FILE, rejectionData);
    auditLog(null, rejectionData.uploadedBy, 'UPLOAD', 'Rejection YTD: ' + rejectionData.fileName + ' \u2014 ' + rejectionData.totalOrders + ' orders', req.headers['x-forwarded-for'] || req.ip || '');
    console.log('Rejection saved:', totalRej, 'rej', totalDel, 'del', dbOk ? '(DB+file)' : '(file only)');

    // ── Record this upload's timestamp so the subadmin 10-hour cooldown
    // above can check it next time. Superadmin never writes here since the
    // cooldown check itself is skipped for superadmin. ──
    if (req.user.role === 'subadmin') {
      try { await recordSubadminUpload(req.user.uid, 'rejection'); }
      catch (e) { console.error('Rejection: could not record upload cooldown:', e.message); }
    }

    res.json({ success:true, summary:{totalRej:totalRej,totalDel:totalDel,fileName:req.file.originalname}, autoClassify: classifyStats, needsReview: needsReview });
  } catch(e) {
    console.error('Rejection upload error:', e.message, e.stack);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
}
app.post('/api/rejection/upload', requireAuth, requireRole('superadmin', 'subadmin'), upload.single('file'), rejectionUploadHandler);

// ── Bulk correction: many "unmatched" rows share the exact same raw reason
// text (e.g. 200 orders all saying "CUSTOMER REFUSED - CHANGED MIND"). Rather
// than clicking Save 300 times, pick the category once for that reason text
// and apply it to every row that has it. ──
app.put('/api/rejection/rows/correct-bulk', requireAuth, requireRole('superadmin', 'subadmin'), async function(req, res) {
  try {
    var rawReason = toStr(req.body.rawReason);
    var newRootCause = toStr(req.body.rootCause);
    var correctedBy = toStr(req.body.correctedBy) || 'Admin';
    if (!rawReason) return res.status(400).json({ error: 'rawReason is required' });
    if (!newRootCause) return res.status(400).json({ error: 'rootCause is required' });

    // Rows for this reason text are stored under whatever polishRootCause() produced
    // for it at upload time — recompute that same value to find them all.
    var storedValue = polishRootCause(rawReason);

    var canonical = newRootCause;
    var matchExisting = null;
    for (var catUpper of KNOWN_CLEAN_CATEGORIES) {
      if (catUpper === newRootCause.toUpperCase()) { matchExisting = catUpper; break; }
    }
    Object.values(REJECTION_LOOKUPS.reasonToRootCause).forEach(function(v) {
      if (String(v).toUpperCase() === newRootCause.toUpperCase()) canonical = v;
    });

    var ie = deriveInternalExternal(canonical);
    var upd;
    if (ie.matched) {
      upd = await pool.query(
        `UPDATE rejection_rows SET root_cause=$1, root_cause_source='manual', source=$2, corrected_by=$3, corrected_at=NOW() WHERE root_cause_source='unmatched' AND root_cause=$4 RETURNING id`,
        [canonical, ie.binary.toUpperCase(), correctedBy, storedValue]
      );
    } else {
      upd = await pool.query(
        `UPDATE rejection_rows SET root_cause=$1, root_cause_source='manual', corrected_by=$2, corrected_at=NOW() WHERE root_cause_source='unmatched' AND root_cause=$3 RETURNING id`,
        [canonical, correctedBy, storedValue]
      );
    }
    if (!upd.rows.length) return res.status(404).json({ error: 'No matching unresolved rows found for that reason text \u2014 they may already be corrected.' });

    var rebuilt = await rebuildRejectionAggregateFromDB();
    // Drop every needsReview entry whose reason text matches this batch — not just one order.
    var preservedNeedsReview = (rejectionData && rejectionData.needsReview || []).filter(function(n) {
      return n.flag !== 'unmatched' || polishRootCause(toStr(n.rawReason)) !== storedValue;
    });
    var preservedAutoClassify = (rejectionData && rejectionData.autoClassify) || null;
    rejectionData = { uploadedAt: new Date().toISOString(), uploadedBy: (rejectionData && rejectionData.uploadedBy) || 'Admin', fileName: (rejectionData && rejectionData.fileName) || '', totalOrders: rebuilt.totalRej + rebuilt.totalDel, orgs: rebuilt.orgsOut, months: rebuilt.monthsOut, needsReview: preservedNeedsReview, autoClassify: preservedAutoClassify };
    await dbSaveRejection(rejectionData.uploadedBy, rejectionData.fileName, rejectionData.totalOrders, rebuilt.orgsOut, rebuilt.monthsOut, preservedNeedsReview, preservedAutoClassify);
    saveJSON(REJECTION_FILE, rejectionData);
    auditLog(null, correctedBy, 'REJECTION_BULK_CORRECTION', 'Bulk-corrected ' + upd.rows.length + ' rows with reason "' + rawReason.slice(0,80) + '" \u2014 Root Cause set to "' + canonical + '"', req.headers['x-forwarded-for'] || req.ip || '');

    res.json({ success: true, updated: upd.rows.length, rootCause: canonical, totalRej: rebuilt.totalRej, totalDel: rebuilt.totalDel });
  } catch (e) {
    console.error('Rejection bulk correction error:', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});


// notes box) and push it to the DB — the dashboard rebuilds from ALL stored
// rows immediately after, exactly the same way an upload does, so the
// correction shows up right away without needing a fresh file upload. ──
app.put('/api/rejection/rows/correct', requireAuth, requireRole('superadmin', 'subadmin'), async function(req, res) {
  try {
    var orderNo = toStr(req.body.orderNo);
    var newRootCause = toStr(req.body.rootCause);
    var correctedBy = toStr(req.body.correctedBy) || 'Admin';
    if (!orderNo) return res.status(400).json({ error: 'orderNo is required' });
    if (!newRootCause) return res.status(400).json({ error: 'rootCause is required' });

    // If the typed category matches a known clean one (any casing), use the
    // canonical casing so it groups correctly with existing data instead of
    // creating a near-duplicate category.
    var canonical = newRootCause;
    var matchExisting = null;
    for (var catUpper of KNOWN_CLEAN_CATEGORIES) {
      if (catUpper === newRootCause.toUpperCase()) { matchExisting = catUpper; break; }
    }
    // Look through the learned lookup's own values for exact-casing match (KNOWN_CLEAN_CATEGORIES is uppercased-only)
    Object.values(REJECTION_LOOKUPS.reasonToRootCause).forEach(function(v) {
      if (String(v).toUpperCase() === newRootCause.toUpperCase()) canonical = v;
    });

    var ie = deriveInternalExternal(canonical);

    var upd;
    if (ie.matched) {
      upd = await pool.query(
        `UPDATE rejection_rows SET root_cause=$1, root_cause_source='manual', source=$2, corrected_by=$3, corrected_at=NOW() WHERE order_no=$4 RETURNING id`,
        [canonical, ie.binary.toUpperCase(), correctedBy, orderNo]
      );
    } else {
      upd = await pool.query(
        `UPDATE rejection_rows SET root_cause=$1, root_cause_source='manual', corrected_by=$2, corrected_at=NOW() WHERE order_no=$3 RETURNING id`,
        [canonical, correctedBy, orderNo]
      );
    }
    if (!upd.rows.length) return res.status(404).json({ error: 'No stored row found with that order number \u2014 it may be from before per-row storage existed, or already deleted.' });

    var rebuilt = await rebuildRejectionAggregateFromDB();
    var preservedNeedsReview = (rejectionData && rejectionData.needsReview || []).filter(function(n) { return n.orderNo !== orderNo; });
    var preservedAutoClassify = (rejectionData && rejectionData.autoClassify) || null;
    rejectionData = { uploadedAt: new Date().toISOString(), uploadedBy: (rejectionData && rejectionData.uploadedBy) || 'Admin', fileName: (rejectionData && rejectionData.fileName) || '', totalOrders: rebuilt.totalRej + rebuilt.totalDel, orgs: rebuilt.orgsOut, months: rebuilt.monthsOut, needsReview: preservedNeedsReview, autoClassify: preservedAutoClassify };
    await dbSaveRejection(rejectionData.uploadedBy, rejectionData.fileName, rejectionData.totalOrders, rebuilt.orgsOut, rebuilt.monthsOut, preservedNeedsReview, preservedAutoClassify);
    saveJSON(REJECTION_FILE, rejectionData);
    auditLog(null, correctedBy, 'REJECTION_MANUAL_CORRECTION', 'Corrected order ' + orderNo + ' \u2014 Root Cause set to "' + canonical + '"', req.headers['x-forwarded-for'] || req.ip || '');

    res.json({ success: true, updated: upd.rows.length, rootCause: canonical, totalRej: rebuilt.totalRej, totalDel: rebuilt.totalDel });
  } catch (e) {
    console.error('Rejection correction error:', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

// Nicely-cased canonical category list for the correction dropdown (KNOWN_CLEAN_CATEGORIES
// itself is uppercased-only, used for case-insensitive matching — this rebuilds display casing).
var CANONICAL_ROOT_CAUSES = Array.from(new Set([
  'Internal Transfer',
  'Merchandiser Unavailable on Route', 'Pending Goods Return Voucher Not Ready (GRV)', 'Customer System Down',
  'Duplicate Item Received Under Separate LPO', 'Returned — Weather/Road Closure', 'Declined — No Stock Requirement',
  'Declined — LPO Cancelled in Customer System', 'Returned — Receiving Closed for the Day', 'Declined — Insufficient Storage Space',
  'Declined — No Delivery Schedule Confirmed', 'Returned — Customer Payment Not Ready (Cheque/Cash)',
  'Declined — LPO Not Reflected in Customer System', 'Declined — Customer System Under Maintenance',
  'Declined — Insufficient Storage Capacity Today', 'Declined — Outside Receiving Hours'
].concat(Object.values(REJECTION_LOOKUPS.reasonToRootCause)))).sort();

app.get('/api/rejection/root-causes', requireAuth, function(req, res) {
  res.json({ categories: CANONICAL_ROOT_CAUSES });
});

// ── CUSTOMER VISIT PREP: typeahead search over every customer name ever
// seen in rejection_rows, ranked by rejection count so the busiest
// accounts surface first. Used by the search box on the Rejection
// dashboard to find a customer before pulling their full breakdown. ──
// Known chain/brand aliases — typing any one of these searches for the
// whole group, since the same chain can appear under several different
// names in the raw customer_name data (e.g. Carrefour orders are often
// logged under "MAF" or "C4"). Add more groups here as new aliases turn up.
var CUSTOMER_GROUP_ALIASES = {
  carrefour: ['carrefour', 'maf', 'c4'],
  maf: ['carrefour', 'maf', 'c4'],
  c4: ['carrefour', 'maf', 'c4'],
  lulu: ['lulu'],
  nesto: ['nesto'],
  spinneys: ['spinneys'],
  cash: ['cash']
};

app.get('/api/rejection/customer-search', requireAuth, async function(req, res) {
  try {
    var q = toStr(req.query.q || '').trim();
    if (q.length < 2) return res.json({ customers: [] });
    var qLower = q.toLowerCase();
    var groupKey = CUSTOMER_GROUP_ALIASES[qLower] ? qLower : null;
    var terms = CUSTOMER_GROUP_ALIASES[qLower] || [q];

    var whereParts = [], params = [];
    terms.forEach(function(t) { params.push('%' + t + '%'); whereParts.push('customer_name ILIKE $' + params.length); });

    var result = await pool.query(
      `SELECT customer_name,
              COUNT(*) FILTER (WHERE UPPER(status) IN ('REJ','REJECTION','REJECTED','R/D','HOLD','RD')) AS rejected_count,
              COUNT(*) AS total_count
       FROM rejection_rows
       WHERE customer_name IS NOT NULL AND customer_name <> '' AND (` + whereParts.join(' OR ') + `)
       GROUP BY customer_name
       HAVING COUNT(*) FILTER (WHERE UPPER(status) IN ('REJ','REJECTION','REJECTED','R/D','HOLD','RD')) > 0
       ORDER BY rejected_count DESC, customer_name ASC
       LIMIT 12`,
      params
    );
    var customers = result.rows.map(function(r) {
      return { name: r.customer_name, rejected: parseInt(r.rejected_count, 10) || 0, total: parseInt(r.total_count, 10) || 0 };
    });
    var groupOption = null;
    // When the query is a known chain alias (carrefour/maf/c4/lulu/...),
    // add a combined "whole group" option ahead of the individual branch
    // matches — one click gets every branch under that chain rolled into
    // a single view/download, instead of picking through them one by one.
    // The group total is a separate unlimited COUNT, not a sum of the
    // (LIMIT 12) rows above, so it reflects every branch, not just the
    // ones shown in the dropdown.
    if (groupKey) {
      var totalResult = await pool.query(
        `SELECT COUNT(*) FILTER (WHERE UPPER(status) IN ('REJ','REJECTION','REJECTED','R/D','HOLD','RD')) AS rejected_count,
                COUNT(DISTINCT customer_name) AS branch_count
         FROM rejection_rows
         WHERE customer_name IS NOT NULL AND customer_name <> '' AND (` + whereParts.join(' OR ') + `)`,
        params
      );
      var tr = totalResult.rows[0] || {};
      groupOption = { group: groupKey, groupLabel: CUSTOMER_GROUP_LABELS[groupKey] || (groupKey + ' Group'), rejected: parseInt(tr.rejected_count, 10) || 0, branchCount: parseInt(tr.branch_count, 10) || 0 };
    }
    res.json({ customers: customers, group: groupOption });
  } catch (e) {
    console.error('Rejection customer-search error:', e.message);
    res.status(500).json({ error: 'Search failed' });
  }
});

// ── CUSTOMER VISIT PREP: full one-customer breakdown across every row
// ever stored for them — branch(org)-wise, month-wise, root-cause,
// internal/external and food/non-food splits, plus a capped list of the
// most recent rejected orders — everything needed to brief a customer
// visit without re-deriving it from the already-capped client-side
// aggregates (which only keep the top ~100 cust+addr+root combos).
// Shared by both /customer-detail (JSON, for the on-screen panel) and
// /customer-export-excel (the downloadable workbook) so the two never
// drift out of sync with each other. ──
// Plain-English explanations for the standard root-cause categories, so a
// rep can answer "what does this actually mean?" in front of a customer
// without translating internal jargon on the spot. Falls back to the raw
// category text (title-cased into a sentence) for anything not mapped.
var ROOT_CAUSE_EXPLANATIONS = {
  'RETURNED — RECEIVING TIME EXCEEDED (RUSH DELIVERY)': 'The truck arrived but the branch\u2019s receiving window had already closed for the day, so the order was returned unopened.',
  'CUSTOMER SYSTEM DOWN': 'The customer\u2019s own ordering/receiving system was down at the time of delivery, so their team couldn\u2019t process the goods in.',
  'DECLINED — LPO DATE EXPIRED': 'The purchase order (LPO) the customer raised had passed its valid date by the time the delivery arrived.',
  'DECLINED — NO STOCK REQUIREMENT': 'The branch said they didn\u2019t actually need the stock at that time (over-ordered, or a duplicate order).',
  'DECLINED — NEAR EXPIRY PRODUCT': 'The branch felt the shelf life remaining on the product was too short for them to sell it in time.',
  'DUPLICATE ITEM RECEIVED UNDER SEPARATE LPO': 'The same item arrived twice under two different purchase orders, so one delivery was declined as a duplicate.',
  'DECLINED — BARCODE SCANNING ISSUE': 'The branch\u2019s scanner couldn\u2019t read the product barcode at receiving, so the order was rejected.',
  'DECLINED — INSUFFICIENT STORAGE SPACE': 'The branch didn\u2019t have physical shelf/storage space available to receive the stock at that time.',
  'DECLINED — PRICING ISSUE': 'The price on the invoice/LPO didn\u2019t match what the branch had agreed, so they declined to receive it.',
  'DECLINED — LPO CANCELLED IN CUSTOMER SYSTEM': 'The customer cancelled the purchase order in their own system after it was placed, before the delivery arrived.',
  'DECLINED — OUTDATED TRN NUMBER': 'The tax registration number on the invoice was out of date on the customer\u2019s records.',
  'DECLINED — STOCK COUNT IN OUTLET': 'The branch was doing a stock count / inventory audit at the time and couldn\u2019t receive new stock.',
  'RETURNED — RECEIVING CLOSED FOR THE DAY': 'The delivery arrived after the branch\u2019s receiving hours had ended for the day.',
  'DECLINED — NO DELIVERY SCHEDULE CONFIRMED': 'No delivery slot/appointment had been confirmed with the branch in advance.',
  'DECLINED — MUNICIPALITY INSPECTION': 'A municipality/health inspection was happening at the branch, so receiving was paused.',
  'DECLINED — FREEZER NOT WORKING': 'The branch\u2019s freezer/cold storage was not working, so they couldn\u2019t safely accept frozen stock.',
  'DECLINED — SHOP UNDER MAINTENANCE': 'The branch was closed or under maintenance/renovation at the time of delivery.',
  'CUSTOMER REQUESTED NEXT-DAY DELIVERY': 'The customer asked for the delivery to be rescheduled to the next day.',
  'INTERNAL TRANSFER': 'This was an internal stock movement between our own warehouses, not a real customer order — it isn\u2019t a customer rejection.',
  'DECLINED — LPO NOT REFLECTED IN CUSTOMER SYSTEM': 'The purchase order wasn\u2019t showing yet in the customer\u2019s own system when the delivery arrived.',
  'DECLINED — CUSTOMER SYSTEM UNDER MAINTENANCE': 'The customer\u2019s ordering/receiving system was offline for scheduled maintenance.',
  'DECLINED — OUTSIDE RECEIVING HOURS': 'The delivery arrived outside the branch\u2019s official receiving hours.',
  'RETURNED — CUSTOMER PAYMENT NOT READY (CHEQUE/CASH)': 'The branch wasn\u2019t able to settle payment (cheque/cash) at the point of delivery.',
  'PENDING GOODS RETURN VOUCHER NOT READY (GRV)': 'A prior return hadn\u2019t been paperworked (GRV) yet, so this delivery was held until that was cleared.',
  'MERCHANDISER UNAVAILABLE ON ROUTE': 'No merchandiser was available on the route to receive/shelve the delivery at the branch.',
  'DECLINED — INSUFFICIENT STORAGE CAPACITY TODAY': 'The branch\u2019s storage was full on that specific day (temporary, not a standing space issue).'
};
function humanizeRootCause(cause) {
  var key = String(cause || '').toUpperCase().trim();
  if (ROOT_CAUSE_EXPLANATIONS[key]) return ROOT_CAUSE_EXPLANATIONS[key];
  if (/^DRIVER SPEND/i.test(key)) return 'The driver ran long at an earlier stop on the route, so this delivery arrived too late to be received — an internal routing/scheduling issue, not caused by this customer.';
  if (!cause || key === 'UNCATEGORIZED') return 'Reason not yet categorized in the system.';
  return String(cause);
}

var CUSTOMER_GROUP_LABELS = { carrefour: 'Carrefour Group (MAF / C4 / Carrefour)', maf: 'Carrefour Group (MAF / C4 / Carrefour)', c4: 'Carrefour Group (MAF / C4 / Carrefour)', lulu: 'Lulu Group', nesto: 'Nesto Group', spinneys: 'Spinneys Group', cash: 'Cash Customers' };

async function buildCustomerBreakdown(name, monthsFilter, groupKey, typeFilter) {
  var result;
  var displayName = name;
  if (groupKey && CUSTOMER_GROUP_ALIASES[groupKey]) {
    var terms = CUSTOMER_GROUP_ALIASES[groupKey];
    var whereParts = [], params = [];
    terms.forEach(function(t) { params.push('%' + t + '%'); whereParts.push('customer_name ILIKE $' + params.length); });
    displayName = CUSTOMER_GROUP_LABELS[groupKey] || (groupKey + ' Group');
    result = await pool.query(
      `SELECT * FROM rejection_rows WHERE (` + whereParts.join(' OR ') + `) ORDER BY entry_date DESC NULLS LAST`,
      params
    );
  } else {
    result = await pool.query(
      `SELECT * FROM rejection_rows WHERE customer_name = $1 ORDER BY entry_date DESC NULLS LAST`,
      [name]
    );
  }
  var rows = result.rows;
  if (!rows.length) return { hasData: false, customer: displayName };

  // Drop stray/mistyped-year rows before any aggregation happens, so every
  // sheet and KPI (branch, month, root cause, totals) stays consistent with
  // each other — e.g. a lone "2022-06" row sitting years before the rest of
  // the data is almost certainly a data-entry typo, not real history.
  // Keep only the current data year and the one before it, based on
  // whichever year is newest in this customer's own rows.
  var yearsPresent = rows.map(function(r) { var m = /^(\d{4})/.exec(r.month_key || ''); return m ? parseInt(m[1], 10) : null; }).filter(function(y) { return y; });
  if (yearsPresent.length) {
    var maxYear = Math.max.apply(null, yearsPresent);
    rows = rows.filter(function(r) {
      var m = /^(\d{4})/.exec(r.month_key || '');
      var y = m ? parseInt(m[1], 10) : maxYear;
      return y >= maxYear - 1;
    });
  }
  if (!rows.length) return { hasData: false, customer: name };

  // Full list of months this customer has any activity in, captured BEFORE
  // the optional month filter below is applied — so the month picker on
  // screen always offers every real option, even while a filter is active.
  var availableMonths = Array.from(new Set(rows.map(function(r) { return r.month_key || 'Unknown'; }))).sort();

  if (monthsFilter && monthsFilter.length) {
    var monthsSet = {};
    monthsFilter.forEach(function(m) { monthsSet[m] = true; });
    rows = rows.filter(function(r) { return monthsSet[r.month_key || 'Unknown']; });
    if (!rows.length) return { hasData: false, customer: displayName, availableMonths: availableMonths };
  }

  if (typeFilter === 'food' || typeFilter === 'nonfood') {
    rows = rows.filter(function(r) { return typeFilter === 'food' ? !!r.is_food : !!r.is_nonfood; });
    if (!rows.length) return { hasData: false, customer: displayName, availableMonths: availableMonths };
  }

  // rejection_rows.status is stored as the short code 'rej' or 'del' (see
  // rowsToStore.push in rejectionUploadHandler — NOT the raw Excel status
  // text), so match on that rather than the long-form REJECTED/DELIVERED
  // strings the original upload-time classifier used on the raw file.
  function isRejStatus(s) { s = String(s || '').toUpperCase(); return s === 'REJ' || s === 'REJECTION' || s === 'REJECTED' || s === 'R/D' || s === 'HOLD' || s === 'RD'; }
  function isDelStatus(s) { s = String(s || '').toUpperCase(); return s === 'DEL' || s === 'DELIVERED' || s.indexOf('DELIVER') >= 0 || s === 'D'; }

  // Every rejection gets bucketed into 'all' AND its own source bucket
  // ('external' or 'internal') so the frontend can flip a toggle and
  // instantly show the customer-safe (external-only) view — root causes,
  // branch/month breakdown, value at risk and the order list all need to
  // change together, not just the KPI count. Delivered orders have no
  // source (they weren't rejected), so they're tracked once and merged
  // into every bucket's branch/month chart data for the Delivered bars.
  // byMonthOrg and byRootOrg add a second dimension — which ORG each
  // month's rejections and each root cause's cases actually happened at —
  // so a rep can answer "which branch had the stock-count issue?" without
  // re-opening the raw data.
  var BUCKETS = ['all', 'external', 'internal'];
  var acc = {};
  BUCKETS.forEach(function(b) { acc[b] = { rejected: 0, valueAtRisk: 0, foodCount: 0, nonfoodCount: 0, byOrg: {}, byMonth: {}, byMonthOrg: {}, byRoot: {}, byRootOrg: {}, byOrgRoot: {}, byOrgArea: {}, byBranch: {}, byBranchRoot: {}, byBranchMonth: {}, recent: [] }; });
  var deliveredByOrg = {}, deliveredByMonth = {}, totalDel = 0;

  rows.forEach(function(r) {
    var org = r.org || 'UNKNOWN';
    var mo = r.month_key || 'Unknown';

    if (isDelStatus(r.status)) {
      totalDel++;
      deliveredByOrg[org] = (deliveredByOrg[org] || 0) + 1;
      deliveredByMonth[mo] = (deliveredByMonth[mo] || 0) + 1;
      return;
    }
    if (!isRejStatus(r.status)) return;

    var val = parseFloat(r.value) || 0;
    var srcU = String(r.source || '').toUpperCase();
    var root = r.root_cause || 'Uncategorized';
    // "Driver Spend N Hours In <some other location> Due To Late Receiving"
    // describes a delay caused at a DIFFERENT stop on the route — it's an
    // AKI routing/scheduling issue, not something this customer did, so it
    // must never surface as External no matter how the row was tagged.
    if (/^DRIVER SPEND/i.test(root)) srcU = 'INTERNAL';
    // Prefer the specific outlet/branch address over the city-level area —
    // "area" is often just "Dubai" for every row, which would collapse all
    // branches into one bucket; "address" is the actual named branch (e.g.
    // "1705110, UNION COOP (BRANCH), Hessa Street"), same as Order Master
    // already shows in its Branch column.
    var area = r.address || r.area || 'Unknown location';
    var targets = ['all'];
    if (srcU === 'EXTERNAL') targets.push('external');
    else if (srcU === 'INTERNAL') targets.push('internal');

    targets.forEach(function(b) {
      var a = acc[b];
      a.rejected++; a.valueAtRisk += val;
      if (r.is_food) a.foodCount++; else if (r.is_nonfood) a.nonfoodCount++;
      if (!a.byOrg[org]) a.byOrg[org] = { org: org, rejected: 0, value: 0 };
      a.byOrg[org].rejected++; a.byOrg[org].value += val;
      if (!a.byMonth[mo]) a.byMonth[mo] = { month: mo, rejected: 0, value: 0 };
      a.byMonth[mo].rejected++; a.byMonth[mo].value += val;
      if (!a.byMonthOrg[mo]) a.byMonthOrg[mo] = {};
      a.byMonthOrg[mo][org] = (a.byMonthOrg[mo][org] || 0) + 1;
      a.byRoot[root] = (a.byRoot[root] || 0) + 1;
      if (!a.byRootOrg[root]) a.byRootOrg[root] = {};
      a.byRootOrg[root][org] = (a.byRootOrg[root][org] || 0) + 1;
      if (!a.byOrgRoot[org]) a.byOrgRoot[org] = {};
      a.byOrgRoot[org][root] = (a.byOrgRoot[org][root] || 0) + 1;
      if (!a.byOrgArea[org]) a.byOrgArea[org] = {};
      if (!a.byOrgArea[org][area]) a.byOrgArea[org][area] = { count: 0, value: 0 };
      a.byOrgArea[org][area].count++; a.byOrgArea[org][area].value += val;
      if (!a.byBranch[area]) a.byBranch[area] = { count: 0, value: 0 };
      a.byBranch[area].count++; a.byBranch[area].value += val;
      if (!a.byBranchRoot[area]) a.byBranchRoot[area] = {};
      a.byBranchRoot[area][root] = (a.byBranchRoot[area][root] || 0) + 1;
      if (!a.byBranchMonth[area]) a.byBranchMonth[area] = {};
      a.byBranchMonth[area][mo] = (a.byBranchMonth[area][mo] || 0) + 1;
      if (a.recent.length < 2000) {
        a.recent.push({ date: r.entry_date, order_no: r.order_no, org: org, root_cause: root, value: val, area: r.area, address: r.address, source: srcU || null, month: mo });
      }
    });
  });

  var totalOrders = acc.all.rejected + totalDel;

  function topBranchesLabel(orgCounts) {
    var list = Object.keys(orgCounts || {}).map(function(k) { return { org: k, n: orgCounts[k] }; }).sort(function(x, y) { return y.n - x.n; });
    return list.slice(0, 3).map(function(o) { return o.org + ' (' + o.n + ')'; }).join(', ') || '—';
  }

  function finalizeBucket(b) {
    var a = acc[b];
    var orgKeys = {}; Object.keys(a.byOrg).forEach(function(k) { orgKeys[k] = true; }); Object.keys(deliveredByOrg).forEach(function(k) { orgKeys[k] = true; });
    var orgList = Object.keys(orgKeys).map(function(k) {
      var o = a.byOrg[k] || { rejected: 0, value: 0 };
      var rootsForOrg = Object.keys(a.byOrgRoot[k] || {}).map(function(rk) {
        var cnt = a.byOrgRoot[k][rk];
        return { root_cause: rk, count: cnt, pct: o.rejected ? (cnt / o.rejected * 100) : 0, explanation: humanizeRootCause(rk) };
      }).sort(function(x, y) { return y.count - x.count; });
      var areasForOrg = Object.keys(a.byOrgArea[k] || {}).map(function(ak) {
        var av = a.byOrgArea[k][ak];
        var rootsForArea = Object.keys(a.byBranchRoot[ak] || {}).map(function(rk) {
          var cnt = a.byBranchRoot[ak][rk];
          return { root_cause: rk, count: cnt, pct: av.count ? (cnt / av.count * 100) : 0 };
        }).sort(function(x, y) { return y.count - x.count; });
        return { area: ak, count: av.count, pct: o.rejected ? (av.count / o.rejected * 100) : 0, value: av.value, rootCauses: rootsForArea };
      }).sort(function(x, y) { return y.count - x.count; });
      return { org: k, rejected: o.rejected, delivered: deliveredByOrg[k] || 0, value: o.value, rootCauses: rootsForOrg, areas: areasForOrg };
    }).sort(function(x, y) { return y.rejected - x.rejected; });
    // Fixed set of ORG codes to use as column order everywhere an ORG
    // breakdown is shown (known codes first, anything else after) — only
    // ORGs with at least one rejection, so zero-rejection codes like
    // TPN/TPW/UNKNOWN don't pad out Monthly Trend with empty columns.
    var orgsWithAnyRejection = {};
    orgList.forEach(function(o) { if (o.rejected > 0) orgsWithAnyRejection[o.org] = true; });
    var knownOrgOrder = ['DCV', 'DCF', 'DGC', 'DGS', 'DSN'];
    var orgColumnOrder = knownOrgOrder.filter(function(o) { return orgsWithAnyRejection[o]; })
      .concat(Object.keys(orgsWithAnyRejection).filter(function(o) { return knownOrgOrder.indexOf(o) < 0; }).sort());

    var moKeys = {}; Object.keys(a.byMonth).forEach(function(k) { moKeys[k] = true; }); Object.keys(deliveredByMonth).forEach(function(k) { moKeys[k] = true; });
    var monthList = Object.keys(moKeys).map(function(k) {
      var m = a.byMonth[k] || { rejected: 0, value: 0 };
      return { month: k, rejected: m.rejected, delivered: deliveredByMonth[k] || 0, value: m.value, byOrg: a.byMonthOrg[k] || {} };
    })
      .filter(function(m) { return m.rejected > 0; }) // skip months with nothing to report
      .sort(function(x, y) { return x.month < y.month ? -1 : (x.month > y.month ? 1 : 0); });

    var rootList = Object.keys(a.byRoot).map(function(k) {
      return {
        root_cause: k, count: a.byRoot[k], pct: a.rejected ? (a.byRoot[k] / a.rejected * 100) : 0,
        explanation: humanizeRootCause(k), topBranches: topBranchesLabel(a.byRootOrg[k])
      };
    }).sort(function(x, y) { return y.count - x.count; });

    // Real branch (outlet/address) rejection summary — separate from the
    // ORG-level byOrg above. Each entry is one physical branch with its
    // own root-cause breakdown, so clicking a branch answers "why did AWeer
    // reject 14 orders" without cross-referencing another sheet/panel.
    var branchList = Object.keys(a.byBranch).map(function(k) {
      var bd = a.byBranch[k];
      var rootsForBranch = Object.keys(a.byBranchRoot[k] || {}).map(function(rk) {
        var cnt = a.byBranchRoot[k][rk];
        return { root_cause: rk, count: cnt, pct: bd.count ? (cnt / bd.count * 100) : 0, explanation: humanizeRootCause(rk) };
      }).sort(function(x, y) { return y.count - x.count; });
      var monthsForBranch = Object.keys(a.byBranchMonth[k] || {}).map(function(mk) {
        var cnt = a.byBranchMonth[k][mk];
        return { month: mk, count: cnt, pct: bd.count ? (cnt / bd.count * 100) : 0 };
      }).sort(function(x, y) { return x.month < y.month ? -1 : (x.month > y.month ? 1 : 0); });
      return { name: k, count: bd.count, value: bd.value, pct: a.rejected ? (bd.count / a.rejected * 100) : 0, rootCauses: rootsForBranch, months: monthsForBranch };
    }).sort(function(x, y) { return y.count - x.count; });

    return {
      rejected: a.rejected,
      rejectionRate: totalOrders > 0 ? Math.round((a.rejected / totalOrders * 100) * 100) / 100 : 0,
      valueAtRisk: Math.round(a.valueAtRisk),
      foodSplit: { food: a.foodCount, nonfood: a.nonfoodCount },
      byOrg: orgList,
      orgColumnOrder: orgColumnOrder,
      byMonth: monthList,
      rootCauses: rootList,
      branches: branchList,
      recentRejections: a.recent
    };
  }

  return {
    hasData: true,
    customer: displayName,
    totalOrders: totalOrders,
    delivered: totalDel,
    sourceSplit: { internal: acc.internal.rejected, external: acc.external.rejected },
    bySource: { all: finalizeBucket('all'), external: finalizeBucket('external'), internal: finalizeBucket('internal') },
    availableMonths: availableMonths
  };
}

app.get('/api/rejection/customer-detail', requireAuth, async function(req, res) {
  try {
    var name = toStr(req.query.name || '').trim();
    var groupKey = toStr(req.query.group || '').trim().toLowerCase();
    if (!name && !groupKey) return res.status(400).json({ error: 'Customer name is required' });
    var monthsFilter = toStr(req.query.months || '').split(',').map(function(m) { return m.trim(); }).filter(Boolean);
    var typeFilter = toStr(req.query.type || '').trim().toLowerCase();
    var breakdown = await buildCustomerBreakdown(name, monthsFilter, groupKey || null, typeFilter || null);
    res.json(breakdown);
  } catch (e) {
    console.error('Rejection customer-detail error:', e.message);
    res.status(500).json({ error: 'Failed to load customer detail' });
  }
});

// ── CUSTOMER VISIT PREP: Excel download of the same breakdown shown
// on-screen (branch-wise, month-wise, root causes, recent orders) —
// the Branch-Wise chart was dropped from the on-screen panel to save
// space, but that data still needs to travel with the customer, so it
// lives here instead. Respects the same all/external/internal view
// (?source=), month selection (?months=), and Food/Non-Food filter
// (?type=) the user had on screen. ──
app.get('/api/rejection/customer-export-excel', requireAuth, async function(req, res) {
  try {
    var name = toStr(req.query.name || '').trim();
    var groupKey = toStr(req.query.group || '').trim().toLowerCase();
    if (!name && !groupKey) return res.status(400).json({ error: 'Customer name is required' });
    var src = toStr(req.query.source || 'all').toLowerCase();
    if (['all', 'external', 'internal'].indexOf(src) < 0) src = 'all';
    var monthsFilter = toStr(req.query.months || '').split(',').map(function(m) { return m.trim(); }).filter(Boolean);
    var typeFilter = toStr(req.query.type || '').trim().toLowerCase();

    var d = await buildCustomerBreakdown(name, monthsFilter, groupKey || null, typeFilter || null);
    if (!d.hasData) return res.status(404).json({ error: 'No rejection rows on file for this customer' + (monthsFilter.length ? ' in the selected month(s)' : '') });
    var b = d.bySource[src];
    var srcLabel = src === 'external' ? 'External (Customer-Facing)' : (src === 'internal' ? 'Internal (AKI-Only)' : 'All');
    var genDate = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8';
    function styleHeaderRow(row) {
      row.eachCell(function(cell) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
        cell.font = { bold: true, color: { argb: GOLD }, size: 11 };
        cell.alignment = { vertical: 'middle' };
      });
    }
    function styleSectionRow(row) {
      row.font = { bold: true, color: { argb: DARKBG }, size: 12 };
      row.eachCell(function(cell) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
    }
    function styleTotalRow(row) {
      row.eachCell(function(cell) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHTGOLD } };
        cell.font = { bold: true, color: { argb: DARKBG } };
      });
    }
    function styleTitleRow(row, span) {
      row.font = { bold: true, size: 13, color: { argb: GOLD } };
      row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
    }
    function noteRow(sheet, text, span) {
      var r = sheet.addRow([text]);
      r.font = { italic: true, size: 9, color: { argb: 'FF666666' } };
      sheet.mergeCells('A' + r.number + ':' + span + r.number);
      return r;
    }

    // Everything below is built as LIVE FORMULAS (COUNTIF/SUMIF/COUNTIFS/
    // SUMIFS) reading off the "Order Master" tab, not pre-computed numbers
    // — so the workbook re-totals itself if a row on Order Master is
    // ever corrected by hand. Delivered/Total-Orders figures are the one
    // exception: this workbook only has row-level detail for REJECTED
    // orders, so those two stay as sourced static numbers (labelled as
    // such), exactly like the reference format.
    var orgCols = b.orgColumnOrder || [];
    var omCount = b.recentRejections.length;
    var OM_START = 3;
    var OM_END = omCount > 0 ? (2 + omCount) : 3;
    var OMS = "'Order Master'!";
    var R_ORG = OMS + '$C$' + OM_START + ':$C$' + OM_END;
    var R_AREA = OMS + '$E$' + OM_START + ':$E$' + OM_END;
    var R_ROOT = OMS + '$F$' + OM_START + ':$F$' + OM_END;
    var R_MONTH = OMS + '$J$' + OM_START + ':$J$' + OM_END;
    var R_VALUE = OMS + '$I$' + OM_START + ':$I$' + OM_END;

    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI';
    wb.created = new Date();

    // ══════════════════════ Sheet 1: Executive Summary ══════════════════════
    var ex = wb.addWorksheet('Executive Summary');
    ex.columns = [{ width: 34 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 12 }, { width: 30 }, { width: 10 }, { width: 12 }, { width: 55 }];
    var exTitle = ex.addRow(['CUSTOMER VISIT — DELIVERY REJECTION SUMMARY']);
    styleTitleRow(exTitle);
    ex.mergeCells('A1:I1');
    var exSub = ex.addRow([d.customer + '   \u2014   ' + srcLabel + ' view   \u2014   Generated ' + genDate]);
    exSub.font = { bold: true, size: 12, color: { argb: DARKBG } };
    ex.mergeCells('A2:I2');
    noteRow(ex, 'This pack summarises ' + srcLabel + ' delivery rejections only' + (src === 'external' ? " \u2014 figures caused by AKI Group's own internal handling are excluded." : (src === 'internal' ? ' \u2014 customer-caused (External) figures are excluded.' : '.')), 'I');
    ex.addRow([]);

    styleSectionRow(ex.addRow(['KEY METRICS']));
    var kpiHdr = ex.addRow(['Total Orders (All Sources)', '', 'Rejected (' + srcLabel + ')', '', 'Rejection Rate', '', 'Delivered', '', 'Value At Risk (AED)']);
    kpiHdr.font = { bold: true, size: 10, color: { argb: DARKBG } };
    var totOrdersCell = 'B' + (kpiHdr.number + 1);
    var kpiVal = ex.addRow([d.totalOrders, '', { formula: 'COUNTA(' + R_ORG + ')' }, '', '', '', d.delivered, '', { formula: 'SUM(' + R_VALUE + ')' }]);
    kpiVal.font = { bold: true, size: 16, color: { argb: DARKBG } };
    kpiVal.getCell(3).font = { bold: true, size: 16, color: { argb: 'FFB00000' } };
    kpiVal.getCell(5).value = { formula: 'IFERROR(C' + kpiVal.number + '/A' + kpiVal.number + ',0)' };
    kpiVal.getCell(5).numFmt = '0.0%';
    kpiVal.getCell(5).font = { bold: true, size: 16, color: { argb: 'FFB00000' } };
    kpiVal.getCell(9).numFmt = '#,##0';
    noteRow(ex, 'Total Orders and Delivered are totals provided in the source export (order-level detail for delivered orders is not included in this workbook); Rejected and Value At Risk are calculated directly from the Order Master tab and update automatically.', 'I');
    ex.addRow([]);

    // Two tables side by side: branch summary (left) + food/non-food (right)
    var exSideHdr1 = ex.addRow(['REJECTIONS BY ORG', '', '', '', '', 'FOOD vs NON-FOOD (' + srcLabel + ')']);
    styleSectionRow(exSideHdr1);
    var exBwHdr = ex.addRow(['ORG', 'Rejected', 'Delivered', 'Value At Risk', '% of Rejections', 'Category', 'Orders', '% of Rejections']);
    styleHeaderRow(exBwHdr);
    var bwOrgsWithRejections = b.byOrg.filter(function(o) { return o.rejected > 0; });
    var exFoodRows = [['Food', b.foodSplit.food], ['Non-Food', b.foodSplit.nonfood]];
    var exSideRowsN = Math.max(bwOrgsWithRejections.length, exFoodRows.length);
    for (var si = 0; si < exSideRowsN; si++) {
      var rowArr = ['', '', '', '', '', '', '', ''];
      if (bwOrgsWithRejections[si]) {
        var o0 = bwOrgsWithRejections[si];
        rowArr[0] = o0.org; rowArr[1] = { formula: 'COUNTIF(' + R_ORG + ',"' + o0.org + '")' }; rowArr[2] = o0.delivered;
        rowArr[3] = { formula: 'SUMIF(' + R_ORG + ',"' + o0.org + '",' + R_VALUE + ')' };
      }
      if (exFoodRows[si]) { rowArr[5] = exFoodRows[si][0]; rowArr[6] = exFoodRows[si][1]; }
      var r0 = ex.addRow(rowArr);
      if (bwOrgsWithRejections[si]) {
        r0.getCell(5).value = { formula: 'IFERROR(B' + r0.number + '/$C$' + kpiVal.number + ',0)' };
        r0.getCell(5).numFmt = '0.0%';
        r0.getCell(4).numFmt = '#,##0';
      }
      if (exFoodRows[si]) {
        r0.getCell(8).value = { formula: 'IFERROR(G' + r0.number + '/$C$' + kpiVal.number + ',0)' };
        r0.getCell(8).numFmt = '0.0%';
      }
    }
    var exBwTotalRow = ex.addRow(['TOTAL',
      { formula: 'SUM(B' + (exBwHdr.number + 1) + ':B' + (exBwHdr.number + exSideRowsN) + ')' },
      bwOrgsWithRejections.reduce(function(s, o) { return s + o.delivered; }, 0),
      { formula: 'SUM(D' + (exBwHdr.number + 1) + ':D' + (exBwHdr.number + exSideRowsN) + ')' },
      1, '', '', '']);
    exBwTotalRow.getCell(5).numFmt = '0.0%'; exBwTotalRow.getCell(4).numFmt = '#,##0';
    styleTotalRow(exBwTotalRow);
    noteRow(ex, 'Total Orders (all sources) and Delivered are totals provided in the source export; Rejected and Value At Risk are calculated directly from the Order Master tab and update automatically.', 'I');
    ex.addRow([]);

    // Top 5 root causes
    styleSectionRow(ex.addRow(['TOP 5 ROOT CAUSES DRIVING REJECTIONS']));
    var exRcHdr = ex.addRow(['Root Cause', '', '', '', '', 'Count', '% of Total', '', 'What This Means for the Customer']);
    styleHeaderRow(exRcHdr);
    var top5 = b.rootCauses.slice(0, 5);
    top5.forEach(function(rc0) {
      var rr = ex.addRow([rc0.root_cause, '', '', '', '', { formula: 'COUNTIF(' + R_ROOT + ',A' + (ex.rowCount + 1) + ')' }, '', '', rc0.explanation]);
      rr.getCell(7).value = { formula: 'IFERROR(F' + rr.number + '/$C$' + kpiVal.number + ',0)' };
      rr.getCell(7).numFmt = '0.0%';
      ex.mergeCells('A' + rr.number + ':E' + rr.number);
      ex.mergeCells('I' + rr.number + ':I' + rr.number);
    });
    ex.addRow([]);
    var top5Count = top5.reduce(function(s, r) { return s + r.count; }, 0);
    var top5Pct = b.rejected ? (top5Count / b.rejected * 100) : 0;
    noteRow(ex, 'These ' + top5.length + ' causes together account for ' + top5Count + ' of the ' + b.rejected + ' rejections (' + top5Pct.toFixed(1) + '%)' + (top5[0] ? ' \u2014 led by ' + top5[0].root_cause.toLowerCase() + '.' : '.'), 'I');

    // ══════════════════════ Sheet 2: ORG Summary ══════════════════════
    var bs = wb.addWorksheet('ORG Summary');
    bs.columns = [{ width: 20 }, { width: 12 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 18 }, { width: 14 }, { width: 18 }];
    var bsTitle = bs.addRow(['ORG-WISE SUMMARY \u2014 ' + srcLabel + ' Rejections']);
    styleTitleRow(bsTitle); bs.mergeCells('A1:H1');
    bs.addRow([]);
    styleSectionRow(bs.addRow(['REJECTIONS BY ORG']));
    styleHeaderRow(bs.addRow(['ORG', 'Rejected', 'Delivered', 'Total Orders', 'Rejection Rate', 'Value At Risk (AED)', 'Value At Risk %', 'Avg AED / Rejection']));
    var bsFirstDataRow = bs.rowCount + 1;
    bwOrgsWithRejections.forEach(function(o) {
      var r0 = bs.addRow([o.org,
        { formula: 'COUNTIF(' + R_ORG + ',A' + (bs.rowCount + 1) + ')' },
        o.delivered, null,
        null,
        { formula: 'SUMIF(' + R_ORG + ',A' + (bs.rowCount + 1) + ',' + R_VALUE + ')' },
        null, null]);
      r0.getCell(4).value = { formula: 'B' + r0.number + '+C' + r0.number };
      r0.getCell(5).value = { formula: 'IFERROR(B' + r0.number + '/D' + r0.number + ',0)' }; r0.getCell(5).numFmt = '0.0%';
      r0.getCell(6).numFmt = '#,##0';
      r0.getCell(8).value = { formula: 'IFERROR(F' + r0.number + '/B' + r0.number + ',0)' }; r0.getCell(8).numFmt = '#,##0.0';
    });
    var bsLastDataRow = bs.rowCount;
    var bsTotalRow = bs.addRow(['TOTAL',
      { formula: 'SUM(B' + bsFirstDataRow + ':B' + bsLastDataRow + ')' },
      { formula: 'SUM(C' + bsFirstDataRow + ':C' + bsLastDataRow + ')' },
      { formula: 'SUM(D' + bsFirstDataRow + ':D' + bsLastDataRow + ')' },
      null,
      { formula: 'SUM(F' + bsFirstDataRow + ':F' + bsLastDataRow + ')' },
      1, null]);
    bsTotalRow.getCell(5).value = { formula: 'IFERROR(B' + bsTotalRow.number + '/D' + bsTotalRow.number + ',0)' }; bsTotalRow.getCell(5).numFmt = '0.0%';
    bsTotalRow.getCell(6).numFmt = '#,##0'; bsTotalRow.getCell(7).numFmt = '0.0%';
    bsTotalRow.getCell(8).value = { formula: 'IFERROR(F' + bsTotalRow.number + '/B' + bsTotalRow.number + ',0)' }; bsTotalRow.getCell(8).numFmt = '#,##0.0';
    bwOrgsWithRejections.forEach(function(o, i) {
      var row = bs.getRow(bsFirstDataRow + i);
      row.getCell(7).value = { formula: 'IFERROR(F' + row.number + '/$F$' + bsTotalRow.number + ',0)' };
      row.getCell(7).numFmt = '0.0%';
    });
    styleTotalRow(bsTotalRow);
    noteRow(bs, 'Delivered / Total Orders are sourced from the order management system export. Rejected, Value At Risk and Avg AED/Rejection are calculated live from the Order Master tab.', 'H');

    // ══════════════════════ Sheet 3: Monthly Trend ══════════════════════
    var mt = wb.addWorksheet('Monthly Trend');
    mt.columns = [{ width: 12 }, { width: 12 }, { width: 12 }, { width: 18 }].concat(orgCols.map(function() { return { width: 12 }; })).concat([{ width: 14 }, { width: 16 }]);
    var mtTitle = mt.addRow(['MONTH-WISE TREND \u2014 ' + srcLabel + ' Rejections']);
    styleTitleRow(mtTitle); mt.mergeCells('A1:' + String.fromCharCode(70 + orgCols.length) + '1');
    mt.addRow([]);
    styleSectionRow(mt.addRow(['REJECTIONS & VALUE AT RISK BY MONTH']));
    styleHeaderRow(mt.addRow(['Month', 'Rejected', 'Delivered', 'Value At Risk (AED)'].concat(orgCols.map(function(o) { return o + ' (Rej)'; })).concat(['Rejection Rate', 'MoM Change (Rej)'])));
    var mtFirstDataRow = mt.rowCount + 1;
    if (!b.byMonth.length) {
      mt.addRow(['No months with rejections in this view']);
    } else {
      b.byMonth.forEach(function(m, mi) {
        var rowVals = [m.month,
          { formula: 'COUNTIF(' + R_MONTH + ',A' + (mt.rowCount + 1) + ')' },
          m.delivered,
          { formula: 'SUMIF(' + R_MONTH + ',A' + (mt.rowCount + 1) + ',' + R_VALUE + ')' }
        ].concat(orgCols.map(function(o) {
          return { formula: 'COUNTIFS(' + R_MONTH + ',A' + (mt.rowCount + 1) + ',' + R_ORG + ',"' + o + '")' };
        })).concat([null, null]);
        var r0 = mt.addRow(rowVals);
        r0.getCell(4).numFmt = '#,##0';
        var rateCol = 4 + orgCols.length + 1;
        var momCol = rateCol + 1;
        r0.getCell(rateCol).value = { formula: 'IFERROR(B' + r0.number + '/(B' + r0.number + '+C' + r0.number + '),0)' };
        r0.getCell(rateCol).numFmt = '0.0%';
        if (mi === 0) { r0.getCell(momCol).value = '\u2014'; }
        else { r0.getCell(momCol).value = { formula: 'B' + r0.number + '-B' + (r0.number - 1) }; }
      });
    }
    var mtLastDataRow = mt.rowCount;
    if (b.byMonth.length) {
      var mtTotalRow = mt.addRow(['TOTAL',
        { formula: 'SUM(B' + mtFirstDataRow + ':B' + mtLastDataRow + ')' },
        { formula: 'SUM(C' + mtFirstDataRow + ':C' + mtLastDataRow + ')' },
        { formula: 'SUM(D' + mtFirstDataRow + ':D' + mtLastDataRow + ')' }
      ].concat(orgCols.map(function(o, oi) {
        var col = String.fromCharCode(69 + oi);
        return { formula: 'SUM(' + col + mtFirstDataRow + ':' + col + mtLastDataRow + ')' };
      })).concat([null, null]));
      var mtRateCol = 4 + orgCols.length + 1;
      mtTotalRow.getCell(mtRateCol).value = { formula: 'IFERROR(B' + mtTotalRow.number + '/(B' + mtTotalRow.number + '+C' + mtTotalRow.number + '),0)' };
      mtTotalRow.getCell(mtRateCol).numFmt = '0.0%';
      mtTotalRow.getCell(4).numFmt = '#,##0';
      styleTotalRow(mtTotalRow);
    }
    var todayStr = new Date().toISOString().slice(0, 10);
    noteRow(mt, 'Delivered per month is sourced from the order management system export. Rejected, Value At Risk, branch splits, Rejection Rate and MoM Change are calculated live from the Order Master tab. The most recent month may be partial (data on file through ' + genDate + ').', String.fromCharCode(69 + orgCols.length + 1));

    // ══════════════════════ Sheet 4: Root Cause Analysis ══════════════════════
    var rc = wb.addWorksheet('Root Cause Analysis');
    rc.columns = [{ width: 44 }, { width: 10 }, { width: 12 }].concat(orgCols.map(function() { return { width: 8 }; })).concat([{ width: 18 }]);
    var rcTitle = rc.addRow(['ROOT CAUSE ANALYSIS \u2014 ' + srcLabel + ' Rejections']);
    styleTitleRow(rcTitle); rc.mergeCells('A1:' + String.fromCharCode(65 + 3 + orgCols.length) + '1');
    rc.addRow([]);
    styleSectionRow(rc.addRow(['ALL ROOT CAUSES, RANKED BY VOLUME']));
    styleHeaderRow(rc.addRow(['Root Cause', 'Count', '% of Total'].concat(orgCols).concat(['Value At Risk (AED)'])));
    var rcFirstDataRow = rc.rowCount + 1;
    b.rootCauses.forEach(function(rcRow) {
      var vals = [rcRow.root_cause,
        { formula: 'COUNTIF(' + R_ROOT + ',A' + (rc.rowCount + 1) + ')' },
        null
      ].concat(orgCols.map(function(o) {
        return { formula: 'COUNTIFS(' + R_ROOT + ',A' + (rc.rowCount + 1) + ',' + R_ORG + ',"' + o + '")' };
      })).concat([
        { formula: 'SUMIF(' + R_ROOT + ',A' + (rc.rowCount + 1) + ',' + R_VALUE + ')' }
      ]);
      var r0 = rc.addRow(vals);
      r0.getCell(3).numFmt = '0.0%';
      r0.getCell(3 + orgCols.length + 1).numFmt = '#,##0';
    });
    var rcLastDataRow = rc.rowCount;
    var rcTotalRow = rc.addRow(['TOTAL',
      { formula: 'SUM(B' + rcFirstDataRow + ':B' + rcLastDataRow + ')' },
      1
    ].concat(orgCols.map(function(o, oi) {
      var col = String.fromCharCode(68 + oi);
      return { formula: 'SUM(' + col + rcFirstDataRow + ':' + col + rcLastDataRow + ')' };
    })).concat([
      { formula: 'SUM(' + String.fromCharCode(68 + orgCols.length) + rcFirstDataRow + ':' + String.fromCharCode(68 + orgCols.length) + rcLastDataRow + ')' }
    ]));
    rcTotalRow.getCell(3).numFmt = '0.0%';
    rcTotalRow.getCell(3 + orgCols.length + 1).numFmt = '#,##0';
    styleTotalRow(rcTotalRow);
    // Fix the % of Total formula now that the TOTAL row's position is known
    for (var rri = rcFirstDataRow; rri <= rcLastDataRow; rri++) {
      rc.getCell('C' + rri).value = { formula: 'IFERROR(B' + rri + '/$B$' + rcTotalRow.number + ',0)' };
    }
    noteRow(rc, 'All figures on this tab are calculated live from the Order Master tab (COUNTIF/COUNTIFS/SUMIF by Root Cause). Ranking order is fixed by current volume; re-sort manually if the mix changes materially.', String.fromCharCode(65 + 3 + orgCols.length));
    rc.addRow([]); rc.addRow([]);
    styleSectionRow(rc.addRow(['TALKING POINTS \u2014 WHAT EACH ROOT CAUSE MEANS FOR THE CUSTOMER']));
    var rcTpHdr = rc.addRow(['Root Cause', '', '', 'What This Means']);
    styleHeaderRow(rcTpHdr);
    b.rootCauses.forEach(function(rcRow) {
      var r0 = rc.addRow([rcRow.root_cause, '', '', rcRow.explanation]);
      rc.mergeCells('A' + r0.number + ':C' + r0.number);
      rc.mergeCells('D' + r0.number + ':' + String.fromCharCode(65 + 3 + orgCols.length) + r0.number);
    });

    // ══════════════════════ Sheet 5: Outlet Detail ══════════════════════
    var od = wb.addWorksheet('Outlet Detail');
    od.columns = [{ width: 14 }, { width: 55 }, { width: 12 }, { width: 12 }, { width: 18 }, { width: 42 }, { width: 12 }, { width: 12 } ];
    var odTitle = od.addRow(['OUTLET-LEVEL DETAIL \u2014 ' + srcLabel + ' Rejections']);
    styleTitleRow(odTitle); od.mergeCells('A1:H1');
    od.addRow([]);
    styleSectionRow(od.addRow(['REJECTIONS BY AREA / OUTLET (ALL BRANCHES)']));
    styleHeaderRow(od.addRow(['ORG', 'Area / Outlet (Address)', 'Rejected', '% of ORG', 'Value At Risk (AED)', 'Top Rejection Reason', 'Reason Count', 'Reason %']));
    var odFirstDataRow = od.rowCount + 1;
    bwOrgsWithRejections.forEach(function(o) {
      (o.areas || []).forEach(function(ar) {
        var topReason = (ar.rootCauses && ar.rootCauses[0]) ? ar.rootCauses[0].root_cause : '';
        var r0 = od.addRow([o.org, ar.area,
          { formula: 'COUNTIFS(' + R_ORG + ',"' + o.org + '",' + R_AREA + ',B' + (od.rowCount + 1) + ')' },
          null,
          { formula: 'SUMIFS(' + R_VALUE + ',' + R_ORG + ',"' + o.org + '",' + R_AREA + ',B' + (od.rowCount + 1) + ')' },
          topReason,
          topReason ? { formula: 'COUNTIFS(' + R_ORG + ',"' + o.org + '",' + R_AREA + ',B' + (od.rowCount + 1) + ',' + R_ROOT + ',F' + (od.rowCount + 1) + ')' } : 0,
          null
        ]);
        r0.getCell(4).value = { formula: 'IFERROR(C' + r0.number + '/COUNTIF(' + R_ORG + ',"' + o.org + '"),0)' };
        r0.getCell(4).numFmt = '0.0%';
        r0.getCell(5).numFmt = '#,##0';
        r0.getCell(8).value = { formula: 'IFERROR(G' + r0.number + '/C' + r0.number + ',0)' };
        r0.getCell(8).numFmt = '0.0%';
      });
    });
    var odLastDataRow = od.rowCount;
    if (odLastDataRow >= odFirstDataRow) {
      var odTotalRow = od.addRow(['TOTAL', '',
        { formula: 'SUM(C' + odFirstDataRow + ':C' + odLastDataRow + ')' }, '',
        { formula: 'SUM(E' + odFirstDataRow + ':E' + odLastDataRow + ')' }, '', '', '']);
      odTotalRow.getCell(5).numFmt = '#,##0';
      styleTotalRow(odTotalRow);
    }
    noteRow(od, 'All figures calculated live from the Order Master tab (COUNTIFS/SUMIFS by Branch and Area/Outlet).', 'E');

    // ══════════════════════ Sheet 6: Order Master (raw data every other
    // tab's formulas point at — always add this one last) ══════════════════════
    var om = wb.addWorksheet('Order Master');
    om.columns = [{ width: 12 }, { width: 18 }, { width: 8 }, { width: 14 }, { width: 46 }, { width: 42 }, { width: 55 }, { width: 10 }, { width: 14 }, { width: 10 }];
    var omTitleRow = om.addRow(['ORDER MASTER DATA \u2014 ' + d.customer + ' \u2014 ' + srcLabel + ' (' + omCount + ' orders)']);
    styleTitleRow(omTitleRow); om.mergeCells('A1:J1');
    styleHeaderRow(om.addRow(['Date', 'Order No', 'ORG', 'Area', 'Branch / Outlet (Address)', 'Root Cause', 'What This Means', 'Source', 'Value (AED)', 'Month']));
    b.recentRejections.forEach(function(r) {
      var dt = r.date ? new Date(r.date) : null;
      om.addRow([dt, r.order_no, r.org, r.area || '\u2014', r.address || '\u2014', r.root_cause, humanizeRootCause(r.root_cause), r.source || '\u2014', Math.round(r.value || 0), r.month || '\u2014']);
      if (dt) om.getRow(om.rowCount).getCell(1).numFmt = 'dd/mm/yyyy';
      om.getRow(om.rowCount).getCell(9).numFmt = '#,##0';
    });

    var fileName = 'Customer_Visit_' + d.customer.replace(/[^a-z0-9]+/gi, '_').slice(0, 40) + '_' + src + '.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fileName + '"');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error('Customer export-excel error:', e.message);
    res.status(500).json({ error: 'Failed to build Excel export' });
  }
});

// ── CUSTOMER VISIT — PRICE DISCREPANCY: same idea as the rejection
// breakdown, scoped to one customer, for the Customer Visit panel's
// collapsible Price Discrepancy section. Matches customer_name the same
// normalized way the price_discrepancy_module itself does (trim + collapse
// whitespace + lowercase), since spelling can vary slightly between the
// two data sources. Also supports group mode (groupKey), using the same
// CUSTOMER_GROUP_ALIASES as the rejection side, so a combined "Carrefour
// Group" view pulls price discrepancies across every branch, not just an
// exact-name match against the group's display label. ──
async function buildCustomerPriceDiscrepancies(name, groupKey) {
  var result, displayName = name;
  if (groupKey && CUSTOMER_GROUP_ALIASES[groupKey]) {
    var terms = CUSTOMER_GROUP_ALIASES[groupKey];
    var whereParts = [], params = [];
    terms.forEach(function(t) { params.push('%' + t + '%'); whereParts.push('customer_name ILIKE $' + params.length); });
    displayName = CUSTOMER_GROUP_LABELS[groupKey] || (groupKey + ' Group');
    result = await pool.query(
      `SELECT *,
              (status = 'Open') AS still_open,
              CASE WHEN status = 'Open' THEN EXTRACT(DAY FROM (NOW() - first_reported_at))::int ELSE NULL END AS days_unresolved
       FROM price_discrepancies
       WHERE (` + whereParts.join(' OR ') + `)
       ORDER BY still_open DESC, days_unresolved DESC NULLS LAST, uploaded_at DESC`,
      params
    );
  } else {
    result = await pool.query(
      `SELECT *,
              (status = 'Open') AS still_open,
              CASE WHEN status = 'Open' THEN EXTRACT(DAY FROM (NOW() - first_reported_at))::int ELSE NULL END AS days_unresolved
       FROM price_discrepancies
       WHERE lower(regexp_replace(trim(customer_name), '\\s+', ' ', 'g')) = lower(regexp_replace(trim($1), '\\s+', ' ', 'g'))
       ORDER BY still_open DESC, days_unresolved DESC NULLS LAST, uploaded_at DESC`,
      [name]
    );
  }
  var rows = result.rows;
  if (!rows.length) return { hasData: false, customer: displayName };

  var openRows = rows.filter(function(r) { return r.status === 'Open'; });
  var resolvedRows = rows.filter(function(r) { return r.status !== 'Open'; });
  var totalOpenValue = openRows.reduce(function(s, r) { return s + (Math.abs(parseFloat(r.discrepancy)) || 0); }, 0);

  var items = rows.map(function(r) {
    return {
      sku_code: r.sku_code, sku_description: r.sku_description,
      system_price: parseFloat(r.system_price) || 0, lpo_price: parseFloat(r.lpo_price) || 0,
      discrepancy: parseFloat(r.discrepancy) || 0, status: r.status,
      remarks: r.remarks || '', week_ending: r.week_ending,
      first_reported_at: r.first_reported_at, times_reported: r.times_reported || 1,
      days_unresolved: r.days_unresolved, reported_by: r.reported_by_full_name || r.reported_by_username,
      customer_name: r.customer_name
    };
  });

  return {
    hasData: true, customer: displayName,
    open: openRows.length, resolved: resolvedRows.length,
    totalOpenValue: Math.round(totalOpenValue),
    items: items
  };
}

app.get('/api/price-discrepancy/customer-detail', requireAuth, async function(req, res) {
  try {
    var name = toStr(req.query.name || '').trim();
    var groupKey = toStr(req.query.group || '').trim().toLowerCase();
    if (!name && !groupKey) return res.status(400).json({ error: 'Customer name is required' });
    var breakdown = await buildCustomerPriceDiscrepancies(name, groupKey || null);
    res.json(breakdown);
  } catch (e) {
    console.error('Customer price-discrepancy detail error:', e.message);
    res.status(500).json({ error: 'Failed to load price discrepancy detail' });
  }
});

app.get('/api/price-discrepancy/customer-export-excel', requireAuth, async function(req, res) {
  try {
    var name = toStr(req.query.name || '').trim();
    var groupKey = toStr(req.query.group || '').trim().toLowerCase();
    if (!name && !groupKey) return res.status(400).json({ error: 'Customer name is required' });
    var showColors = toStr(req.query.show || 'red').toLowerCase().split(',').map(function(s) { return s.trim(); }).filter(Boolean);
    var showRed = showColors.indexOf('red') >= 0, showGreen = showColors.indexOf('green') >= 0;
    if (!showRed && !showGreen) showRed = true; // never export nothing
    var d = await buildCustomerPriceDiscrepancies(name, groupKey || null);
    if (!d.hasData) return res.status(404).json({ error: 'No price discrepancies on file for this customer' });

    var fullCount = d.items.length;
    var exportItems = d.items.filter(function(it) { return (it.discrepancy < 0 && showRed) || (it.discrepancy >= 0 && showGreen); });

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8';
    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI';
    wb.created = new Date();

    var ws = wb.addWorksheet('Price Discrepancies');
    ws.columns = [{ width: 26 }, { width: 20 }, { width: 40 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 14 }, { width: 40 }, { width: 22 }];
    var title = ws.addRow(['PRICE DISCREPANCIES \u2014 ' + d.customer]);
    title.font = { bold: true, size: 14, color: { argb: GOLD } };
    title.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
    ws.mergeCells('A1:J1');
    var sub = ws.addRow(['Generated ' + new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })]);
    sub.font = { italic: true, size: 9, color: { argb: 'FF666666' } };
    ws.mergeCells('A2:J2');
    ws.addRow([]);
    var kpiRow = ws.addRow(['Open', d.open, 'Resolved', d.resolved]);
    kpiRow.font = { bold: true };
    ws.addRow([]);
    var filterLabel = (showRed && showGreen) ? 'Red + Green (all rows)' : (showRed ? 'Red only (price shortfalls)' : 'Green only (LPO already higher)');
    var filterRow = ws.addRow(['Filter applied: ' + filterLabel + ' \u2014 ' + exportItems.length + ' of ' + fullCount + ' total rows']);
    filterRow.font = { bold: true, size: 10, color: { argb: 'FFB00000' } };
    ws.mergeCells('A' + filterRow.number + ':J' + filterRow.number);
    var noteRow = ws.addRow(['Discrepancy = LPO price minus system price. Red = system price higher than LPO (the real issue). Green = customer\u2019s LPO already higher than system (not an issue to raise).']);
    noteRow.font = { italic: true, size: 9, color: { argb: 'FF666666' } };
    ws.mergeCells('A' + noteRow.number + ':J' + noteRow.number);
    var hdr = ws.addRow(['Branch', 'SKU Code', 'SKU Description', 'System Price', 'LPO Price', 'Discrepancy', 'Status', 'Times Reported', 'Remarks', 'Reported By']);
    hdr.eachCell(function(cell) {
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
      cell.font = { bold: true, color: { argb: GOLD }, size: 11 };
    });
    exportItems.forEach(function(it) {
      var row = ws.addRow([it.customer_name, it.sku_code, it.sku_description, it.system_price, it.lpo_price, it.discrepancy, it.status, it.times_reported, it.remarks, it.reported_by]);
      row.getCell(6).font = { color: { argb: it.discrepancy < 0 ? 'FFB00000' : 'FF1A8A3A' }, bold: true };
      if (it.status === 'Open') {
        row.getCell(7).font = { color: { argb: 'FFB00000' }, bold: true };
      } else {
        row.eachCell(function(cell) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHTGOLD } }; });
      }
    });

    var fileNameTag = (showRed && showGreen) ? '_all' : (showRed ? '_red' : '_green');
    var fileName = 'Price_Discrepancy_' + d.customer.replace(/[^a-z0-9]+/gi, '_').slice(0, 40) + fileNameTag + '.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fileName + '"');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error('Customer price-discrepancy export-excel error:', e.message);
    res.status(500).json({ error: 'Failed to build Excel export' });
  }
});


// ── AI-assisted bulk classification: when grouping alone doesn't help much
// (e.g. 158 distinct reason texts out of 166 rows — almost nothing repeats),
// send the batch to Claude once and get a best-guess category for each,
// so the human is reviewing/confirming pre-filled picks instead of reading
// and deciding every single one from a blank dropdown. ──
app.post('/api/rejection/suggest-categories', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var reasons = Array.isArray(req.body.reasons) ? req.body.reasons.slice(0, 200) : [];
    if (!reasons.length) return res.status(400).json({ error: 'No reason texts provided' });

    var prompt = 'You are classifying delivery-rejection reason texts into standard root-cause categories for a UAE FMCG distributor (AKI Group).\n\n' +
      'Known categories — use EXACTLY one of these (matching text) if it genuinely fits:\n' +
      CANONICAL_ROOT_CAUSES.map(function(c){ return '- ' + c; }).join('\n') + '\n\n' +
      'For each numbered reason text below, pick the single best-fitting category from the list above. ' +
      'Only if truly none of them fit, propose a short new category in the same style (Title Case, similar phrasing pattern, e.g. "Declined — ..." or "Returned — ...").\n\n' +
      'Reason texts:\n' +
      reasons.map(function(r,i){ return (i+1) + '. "' + r + '"'; }).join('\n') + '\n\n' +
      'Respond with ONLY a JSON array like [{"index":1,"category":"..."}, ...], one entry per numbered reason above, matching indexes exactly. No other text, no markdown fences.';

    var msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001', max_tokens: 4000,
      messages: [{ role: 'user', content: prompt }]
    });

    var text = (msg.content || []).map(function(b){ return b.text || ''; }).join('');
    var clean = text.replace(/```json|```/g, '').trim();
    var parsed;
    try { parsed = JSON.parse(clean); } catch(e) {
      return res.status(500).json({ error: 'AI response wasn\u2019t valid JSON — try again.' });
    }

    var suggestions = {};
    parsed.forEach(function(p) {
      var idx = (p.index || 0) - 1;
      if (reasons[idx] !== undefined && p.category) suggestions[reasons[idx]] = String(p.category).trim();
    });

    res.json({ success: true, suggestions: suggestions });
  } catch(e) {
    console.error('Rejection suggest-categories error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.get('/api/rejection/status', function(req, res) {
  res.setHeader('Content-Type', 'application/json');
  if (!rejectionData) return res.json({ hasData:false });
  res.json({ hasData:true, uploadedAt:rejectionData.uploadedAt, uploadedBy:rejectionData.uploadedBy, fileName:rejectionData.fileName, totalOrders:rejectionData.totalOrders, orgs:rejectionData.orgs, months:rejectionData.months, needsReupload:rejectionData.needsReupload||false, needsReview:rejectionData.needsReview||[], autoClassify:rejectionData.autoClassify||null });
});

// ── REJECTION EXCEL EXPORT (server-side, 2 sheets: Executive Summary + Detail, styled) ──
app.post('/api/rejection/export-excel', async function(req, res) {
  try {
    var body = req.body || {};
    var summary = body.summary || {};
    var detailRows = body.detailRows || [];
    var filterStr = body.filterStr || 'All Filters';

    var totalRej = summary.tRej || 0;
    var totalDel = summary.tDel || 0;
    var rate = (totalRej + totalDel) > 0 ? (totalRej/(totalRej+totalDel)*100).toFixed(2)+'%' : '0%';
    var topReasons = (summary.reasons || []).slice(0, 5);
    var topCusts = (summary.custs || []).slice(0, 5);

    // For each top root cause, find the single customer+branch driving the most cases of it
    var rootCauseTop = {};
    detailRows.forEach(function(r){
      var key = r.root || 'Unknown';
      if(!rootCauseTop[key] || (r.n||0) > rootCauseTop[key].n){
        rootCauseTop[key] = { cust: r.cust||'—', addr: r.addr||'—', n: r.n||0 };
      }
    });

    // Group detail rows by customer to find repeat offenders (same customer, multiple branch addresses)
    // Track rejection count PER branch (not just presence) so we can list the actual branches
    var custGroups = {};
    detailRows.forEach(function(r){
      var key = r.cust || 'Unknown';
      if(!custGroups[key]) custGroups[key] = { branches: {}, total: 0 };
      var addrKey = r.addr || 'Unknown address';
      custGroups[key].branches[addrKey] = (custGroups[key].branches[addrKey] || 0) + (r.n || 0);
      custGroups[key].total += (r.n || 0);
    });
    var repeatOffenders = Object.keys(custGroups)
      .map(function(name){
        var branchList = Object.keys(custGroups[name].branches)
          .map(function(a){ return { addr:a, n:custGroups[name].branches[a] }; })
          .sort(function(x,y){ return y.n - x.n; });
        return { name:name, branchCount: branchList.length, total: custGroups[name].total, branches: branchList };
      })
      .filter(function(c){ return c.branchCount >= 2; })
      .sort(function(a,b){ return (b.branchCount - a.branchCount) || (b.total - a.total); })
      .slice(0, 8);

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8', WHITE = 'FFFFFFFF';
    function styleHeaderRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
        cell.font = { bold:true, color:{argb:GOLD}, size:11 };
        cell.alignment = { vertical:'middle' };
      });
    }
    function styleSectionRow(row){
      row.font = { bold:true, color:{argb:DARKBG}, size:12 };
      row.eachCell(function(cell){ cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:GOLD} }; });
    }
    function styleTotalRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:LIGHTGOLD} };
        cell.font = { bold:true, color:{argb:DARKBG} };
      });
    }
    function styleCustRow(row){
      row.font = { bold:true };
      row.eachCell(function(cell){ cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:'FFEFEFEF'} }; });
    }

    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI';
    wb.created = new Date();

    // ---- Sheet 1: Executive Summary ----
    var es = wb.addWorksheet('Executive Summary');
    es.columns = [{width:38},{width:32},{width:14},{width:14},{width:34},{width:20},{width:20}];

    var titleRow = es.addRow(['AKI GROUP — REJECTION EXECUTIVE SUMMARY']);
    titleRow.font = { bold:true, size:15, color:{argb:GOLD} };
    es.mergeCells('A1:E1');
    es.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
    var genRow = es.addRow(['Generated', null]);
    var genCell = genRow.getCell(2);
    genCell.value = { formula: 'NOW()' };
    genCell.numFmt = 'dd/mm/yyyy hh:mm:ss';
    genCell.font = { bold: true, color: { argb: 'FFFF0000' } };
    es.addRow(['Filters Applied', filterStr]);
    es.addRow([]);
    styleSectionRow(es.addRow(['KEY METRICS']));
    es.addRow(['Total Rejections', totalRej]);
    es.addRow(['Total Delivered', totalDel]);
    es.addRow(['Rejection Rate', rate]);
    es.addRow(['Value at Risk', summary.val || '—']);
    es.addRow([]);

    // TOP 5 ROOT CAUSES — now shows WHO (customer + branch) is the biggest driver of each cause
    styleHeaderRow(es.addRow(['ROOT CAUSE', 'TOP CUSTOMER DRIVING IT', 'COUNT', '% OF TOTAL', 'TOP BRANCH (CASES)']));
    topReasons.forEach(function(r){
      var top = rootCauseTop[r.l] || { cust:'—', addr:'—', n:0 };
      es.addRow([r.l, top.cust, r.n, totalRej>0 ? ((r.n/totalRej*100).toFixed(1)+'%') : '0%', top.addr + ' ('+top.n+')']);
    });
    es.addRow([]);

    styleHeaderRow(es.addRow(['TOP 5 CUSTOMERS BY REJECTIONS', '', 'COUNT', '% OF TOTAL', '']));
    topCusts.forEach(function(c){
      es.addRow([c.n, '', c.c, totalRej>0 ? ((c.c/totalRej*100).toFixed(1)+'%') : '0%', '']);
    });
    es.addRow([]);

    // Repeat Offenders — customer name AND the actual branch addresses driving their total
    styleHeaderRow(es.addRow(['REPEAT OFFENDERS — CUSTOMER / BRANCH', '', 'BRANCHES', 'TOTAL REJECTIONS', '']));
    if(repeatOffenders.length){
      repeatOffenders.forEach(function(o){
        styleCustRow(es.addRow([o.name, '', o.branchCount, o.total, '']));
        o.branches.slice(0, 6).forEach(function(b){
          es.addRow(['     • '+b.addr, '', '', b.n, '']);
        });
      });
    } else {
      es.addRow(['No customer repeats across multiple branches in this filtered view.']);
    }
    es.addRow([]);

    styleSectionRow(es.addRow(['RECOMMENDED ACTIONS']));
    if (topReasons[0]) {
      var t0 = rootCauseTop[topReasons[0].l] || {};
      es.addRow(['1. Prioritize a fix for "'+topReasons[0].l+'" — top root cause at '+topReasons[0].n+' rejections ('+(totalRej>0?(topReasons[0].n/totalRej*100).toFixed(1):'0')+'% of total). Biggest driver: "'+(t0.cust||'—')+'" at '+(t0.addr||'—')+' ('+(t0.n||0)+' cases).']);
    }
    if (topReasons[1]) es.addRow(['2. Address "'+topReasons[1].l+'" next — '+topReasons[1].n+' rejections.']);
    if (topCusts[0]) es.addRow(['3. Engage account owner for "'+topCusts[0].n+'" — highest-rejecting customer with '+topCusts[0].c+' cases.']);
    if (repeatOffenders[0]) es.addRow(['4. Investigate "'+repeatOffenders[0].name+'" across '+repeatOffenders[0].branchCount+' branches (see list above) — recurring issue, not a one-off location problem.']);
    es.addRow(['5. Re-check merchandiser/route scheduling if route-related causes dominate the list above.']);
    es.addRow([]);

    // ---- Month-Wise Analysis — part of Executive Summary, not a separate
    // sheet, so the GM sees the full picture on the first tab without
    // clicking between sheets. ----
    var monthWise = body.monthWise || [];
    if (monthWise.length) {
      styleSectionRow(es.addRow(['MONTH-WISE ANALYSIS']));
      styleHeaderRow(es.addRow(['Month','Delivered','Rejection','Delivered %','Rejection %','Food (of Rej.)','Non-Food (of Rej.)']));
      var mwDelSum=0, mwRejSum=0, mwFoodSum=0, mwNonfoodSum=0;
      monthWise.forEach(function(r){
        mwDelSum += r.delivered||0; mwRejSum += r.rejection||0;
        mwFoodSum += r.foodRej||0; mwNonfoodSum += r.nonfoodRej||0;
        es.addRow([
          r.month, r.delivered||0, r.rejection||0,
          (r.deliveredPct||0).toFixed(1)+'%', (r.rejectionPct||0).toFixed(1)+'%',
          (r.foodRej||0)+' ('+(r.foodPct||0).toFixed(1)+'%)',
          (r.nonfoodRej||0)+' ('+(r.nonfoodPct||0).toFixed(1)+'%)'
        ]);
      });
      var mwTotal = mwDelSum + mwRejSum;
      styleTotalRow(es.addRow([
        'TOTAL', mwDelSum, mwRejSum,
        mwTotal ? (100*mwDelSum/mwTotal).toFixed(1)+'%' : '0%',
        mwTotal ? (100*mwRejSum/mwTotal).toFixed(1)+'%' : '0%',
        mwRejSum ? mwFoodSum+' ('+(100*mwFoodSum/mwRejSum).toFixed(1)+'%)' : '0 (0%)',
        mwRejSum ? mwNonfoodSum+' ('+(100*mwNonfoodSum/mwRejSum).toFixed(1)+'%)' : '0 (0%)'
      ]));
      var noteRow = es.addRow(['Food/Non-Food % = share of that month\u2019s rejections specifically, not of total orders. "Frozen" is not currently tracked as a separate category.']);
      noteRow.font = { italic: true, size: 9, color: { argb: 'FF888888' } };
      es.addRow([]);
    }

    // ---- Sheet 2: Rejection Detail ----
    var det = wb.addWorksheet('Rejection Detail');
    det.columns = [{width:5},{width:32},{width:45},{width:30},{width:14},{width:12}];
    styleHeaderRow(det.addRow(['#','Customer Name','Full Address','Final Root Cause','Rejection Count','% of Total']));
    var sumN = 0;
    detailRows.forEach(function(r, i){
      var pct = totalRej>0 ? ((r.n/totalRej*100).toFixed(2)+'%') : '0%';
      sumN += (r.n||0);
      det.addRow([i+1, r.cust||'', r.addr||'', r.root||'', r.n||0, pct]);
    });
    var pctTotal = totalRej>0 ? ((sumN/totalRej*100).toFixed(2)+'%') : '0%';
    styleTotalRow(det.addRow(['', 'TOTAL', '', '', sumN, pctTotal]));

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Rejection_Report_'+Date.now()+'.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch(e) {
    console.error('rejection export-excel error:', e.message);
    res.status(500).json({ error: 'Export failed: '+e.message });
  }
});

// ── DISPATCH DROP ANALYSIS + ROUTE SUMMARY EXCEL EXPORT ──
app.post('/api/dispatch/export-excel', async function(req, res) {
  try {
    var body = req.body || {};
    var ca = body.cost_analysis || {};
    var topRoutes = body.top_routes || [];
    var repeatLocs = body.repeat_locations || [];
    var cityTypeCross = body.city_type_cross || {};

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8', REQBLUE = 'FFD6E8FF', AVOIDRED = 'FFFDE0DE';
    function styleHeaderRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
        cell.font = { bold:true, color:{argb:GOLD}, size:11 };
        cell.alignment = { vertical:'middle' };
      });
    }
    function styleSectionRow(row){
      row.font = { bold:true, color:{argb:DARKBG}, size:12 };
      row.eachCell(function(cell){ cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:GOLD} }; });
    }
    function styleTotalRow(row){
      row.eachCell(function(cell){
        cell.fill = { type:'pattern', pattern:'solid', fgColor:{argb:LIGHTGOLD} };
        cell.font = { bold:true, color:{argb:DARKBG} };
      });
    }

    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI';
    wb.created = new Date();

    var avoidableCount = ca.repeat_location_avoidable_count || 0;
    var totalRepeatCount = ca.repeat_location_count || 0;
    var avoidableRows = repeatLocs.filter(function(l){ return !l.is_legitimate_split; });
    var requiredRows = repeatLocs.filter(function(l){ return l.is_legitimate_split; });

    // ---- Sheet 1: Executive Summary ----
    var es = wb.addWorksheet('Executive Summary');
    es.columns = [{width:36},{width:22},{width:16},{width:16},{width:16}];

    var titleRow = es.addRow(['AKI GROUP — DAILY DISPATCH DROP ANALYSIS']);
    titleRow.font = { bold:true, size:15, color:{argb:GOLD} };
    es.mergeCells('A1:E1');
    es.getRow(1).fill = { type:'pattern', pattern:'solid', fgColor:{argb:DARKBG} };
    es.addRow(['Date', body.date || new Date().toLocaleDateString('en-AE')]);
    es.addRow(['Generated', new Date().toLocaleString('en-AE')]);
    es.addRow([]);

    styleSectionRow(es.addRow(['KEY METRICS']));
    es.addRow(['Total Orders', body.total_orders || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Value (AED)', Math.round(body.total_value || 0)]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Routes', body.total_routes || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Drivers', body.total_drivers || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Total Drops', body.total_drops || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['Own Fleet Drops', ca.own_fleet_drops || 0]).getCell(2).numFmt = '#,##0';
    es.addRow(['3PL Drops (billed separately)', ca.pl_drops || 0]).getCell(2).numFmt = '#,##0';
    es.addRow([]);

    styleSectionRow(es.addRow(['REPEAT-VISIT ADDRESSES — ACTION SUMMARY']));
    es.addRow(['Total addresses visited by 2+ routes today', totalRepeatCount]);
    es.addRow(['— Required (different order types, separate trucks needed)', requiredRows.length]);
    es.addRow(['— Avoidable (same order type, worth questioning)', avoidableCount]);
    es.addRow([]);

    var RED = 'FFB0201A';
    var exceptionRows = avoidableRows.filter(function(l){ return l.is_high_value_exception; });
    var genuineAvoidableRows = avoidableRows.filter(function(l){ return !l.is_high_value_exception; });

    if (exceptionRows.length) {
      var exceptionsByValue = exceptionRows.slice().sort(function(a, b) { return (b.total_value || 0) - (a.total_value || 0); });
      styleHeaderRow(es.addRow(['⚠ HIGH-VALUE EXCEPTIONS (>AED 100,000)', 'CUSTOMER', 'ROUTES (TYPE · ORDERS · VALUE)', 'TOTAL VALUE (AED)', '']));
      es.addRow(['These are same-order-type duplicates, but the value is large enough that they may be a genuinely large order needing a capacity split — verify before treating as a routing mistake.']);
      exceptionsByValue.forEach(function(l){
        var routesLabel = (l.route_types||[]).map(function(rt){ return rt.route+' ('+rt.types.join('+')+', '+(rt.order_count||0)+' orders, AED '+(rt.value||0).toLocaleString()+')'; }).join(', ') || (l.routes||[]).join(', ');
        var row = es.addRow([l.location_id, l.customer, routesLabel, l.total_value || 0, '']);
        row.getCell(4).numFmt = '#,##0';
        row.font = { color:{argb:'FF8B6914'} };
      });
      es.addRow([]);
    }

    if (genuineAvoidableRows.length) {
      styleHeaderRow(es.addRow(['TOP AVOIDABLE REPEAT DROPS', 'CUSTOMER', 'ROUTES (TYPE · ORDERS · VALUE)', 'TOTAL VALUE (AED)', '']));
      genuineAvoidableRows.slice(0, 10).forEach(function(l){
        var routesLabel = (l.route_types||[]).map(function(rt){ return rt.route+' ('+rt.types.join('+')+', '+(rt.order_count||0)+' orders, AED '+(rt.value||0).toLocaleString()+')'; }).join(', ') || (l.routes||[]).join(', ');
        var row = es.addRow([l.location_id, l.customer, routesLabel, l.total_value || 0, '']);
        row.getCell(4).numFmt = '#,##0';
        row.font = { color:{argb:RED} };
      });
      es.addRow([]);
    }

    styleSectionRow(es.addRow(['RECOMMENDED ACTIONS']));
    if (genuineAvoidableRows.length) {
      var biggest = genuineAvoidableRows[0];
      es.addRow(['1. Raise the ' + genuineAvoidableRows.length + ' avoidable repeat-drop addresses with the transport team — same order type sent on 2+ separate trucks to the same address.']);
      es.addRow(['2. Start with "' + biggest.customer + '" (Location ' + biggest.location_id + ') — highest combined value at AED ' + (biggest.total_value||0).toLocaleString() + ' split across routes ' + (biggest.routes||[]).join(', ') + '.']);
      es.addRow(['3. See "Repeat Location Detail" tab for the full list with per-route order values — check whether each split was due to genuine order size before assuming it was a routing error.']);
    } else {
      es.addRow(['No avoidable repeat drops found today — all multi-route visits were legitimate Food/Non-Food/3PL splits.']);
    }
    if (exceptionRows.length) {
      es.addRow(['4. ' + exceptionRows.length + ' high-value exception(s) flagged above (>AED 100,000) — review order size before raising these with the transport team, as a large order may genuinely require 2 trucks.']);
    }

    // ---- Sheet 2: Route Summary ----
    var rt = wb.addWorksheet('Route Summary');
    rt.columns = [{width:14},{width:30},{width:14},{width:12},{width:12},{width:14},{width:16}];
    styleHeaderRow(rt.addRow(['Route', 'Type', 'Vehicle', 'Orders', 'Drivers', 'Locations (Drops)', 'Value (AED)']));
    var sumOrders = 0, sumDrops = 0, sumValue = 0;
    topRoutes.forEach(function(r){
      sumOrders += r.orders || 0; sumDrops += r.drops || 0; sumValue += r.value || 0;
      var typeLabel = (r.types || []).join(', ');
      var vehicleLabel = r.isPartitionVehicle ? 'Partition' : 'Single-Type';
      var row = rt.addRow([r.route, typeLabel, vehicleLabel, r.orders || 0, r.driverCount || 0, r.drops || 0, Math.round(r.value || 0)]);
      row.getCell(7).numFmt = '#,##0';
      if (r.isPartitionVehicle) { row.getCell(3).font = { bold:true, color:{argb:GOLD} }; }
    });
    var totalRow = rt.addRow(['TOTAL', '', '', sumOrders, '', sumDrops, Math.round(sumValue)]);
    totalRow.getCell(7).numFmt = '#,##0';
    styleTotalRow(totalRow);

    // ---- Sheet 3: Transport Cost (rate-card based, only if truck-type column present) ----
    if (body.truck_cost_estimate && body.truck_cost_estimate.available) {
      var tce = body.truck_cost_estimate;
      var tc = wb.addWorksheet('Transport Cost');
      tc.columns = [{width:26},{width:14},{width:14},{width:16}];
      styleSectionRow(tc.addRow(['SUMMARY']));
      tc.addRow(['Total Estimated Cost (AED)', tce.total_estimated_cost || 0]).getCell(2).numFmt = '#,##0';
      tc.addRow(['Total Drops Billed', tce.total_drops_billed || 0]);
      tc.addRow(['Total Vehicles Used', tce.total_vehicles || 0]);
      tc.addRow([]);

      styleHeaderRow(tc.addRow(['BY VEHICLE TYPE', 'Rate (AED/Drop)', 'Drops Billed', 'Est. Cost (AED)']));
      (tce.by_type || []).forEach(function(t){
        var row = tc.addRow([t.label, t.rate, t.drop_count, t.estimated_cost]);
        row.getCell(4).numFmt = '#,##0';
      });
      tc.addRow([]);

      styleHeaderRow(tc.addRow(['BY REGION', '', 'Drops', 'Est. Cost (AED)']));
      (tce.by_region || []).forEach(function(r){
        var row = tc.addRow([r.city, '', r.total_drops, r.total_cost]);
        row.getCell(4).numFmt = '#,##0';
      });
      tc.addRow([]);

      styleHeaderRow(tc.addRow(['BY FOOD / NON-FOOD', '', 'Drops', 'Est. Cost (AED)']));
      (tce.by_food_type || []).forEach(function(f){
        var row = tc.addRow([f.category, '', f.drop_count, f.estimated_cost]);
        row.getCell(4).numFmt = '#,##0';
      });

      var unmatchedKeysExport = Object.keys(tce.unmatched_truck_types || {});
      if (unmatchedKeysExport.length) {
        tc.addRow([]);
        var warnRow = tc.addRow(['⚠ Unmatched truck-type text (excluded from cost, not guessed):']);
        warnRow.font = { italic: true, color: {argb:'FFB0201A'} };
        unmatchedKeysExport.forEach(function(k){
          tc.addRow(['  ' + k, '', tce.unmatched_truck_types[k], '']);
        });
      }
    }

    // ---- Sheet 4: In-House vs Hired Drivers ----
    if (body.driver_source_split) {
      var dss = body.driver_source_split;
      var dh = wb.addWorksheet('In-House vs Hired Drivers');
      dh.columns = [{width:30},{width:32},{width:16},{width:12},{width:12},{width:16}];
      styleSectionRow(dh.addRow(['SUMMARY']));
      dh.addRow(['In-House Drivers', dss.inhouse.driver_count || 0, '', dss.inhouse.orders||0, dss.inhouse.drops||0, dss.inhouse.value||0]).getCell(6).numFmt = '#,##0';
      var hiredSummaryRow = dh.addRow(['Hired Drivers (no name on file)', dss.hired.driver_count || 0, '', dss.hired.orders||0, dss.hired.drops||0, dss.hired.value||0]);
      hiredSummaryRow.getCell(6).numFmt = '#,##0';
      hiredSummaryRow.font = { color:{argb:'FFB0201A'} };
      dh.addRow([]);
      styleHeaderRow(dh.addRow(['Phone/ID', 'Customer', 'Type', 'Orders', 'Drops', 'Value (AED)']));
      (dss.hired_driver_details || []).forEach(function(d){
        var custLabel = (d.customers && d.customers.length) ? d.customers.join(', ') : '—';
        var typeLabel2 = (d.types && d.types.length) ? d.types.join(', ') : '—';
        var row2 = dh.addRow([d.name, custLabel, typeLabel2, d.orders || 0, d.drops || 0, d.value || 0]);
        row2.getCell(6).numFmt = '#,##0';
        row2.getCell(1).font = { color:{argb:'FFB0201A'} };
      });
    }

    // ---- Sheet 5: Repeat Location Detail (one row per individual order for full traceability) ----
    var rl = wb.addWorksheet('Repeat Location Detail');
    rl.columns = [{width:14},{width:30},{width:40},{width:12},{width:26},{width:10},{width:14},{width:14},{width:16}];
    styleHeaderRow(rl.addRow(['Location ID', 'Customer', 'Address', 'Route', 'Order Code', 'Type', 'Order Value (AED)', 'Status', 'Location Total (AED)']));
    var rlRowCount = 1;
    repeatLocs.forEach(function(l){
      var statusText = l.is_legitimate_split ? 'Required' : (l.is_high_value_exception ? 'Exception — Review' : 'Avoidable');
      var statusColor = l.is_legitimate_split ? REQBLUE : (l.is_high_value_exception ? 'FFFFE9A8' : AVOIDRED);
      var fontColor = l.is_legitimate_split ? 'FF1B5E9E' : (l.is_high_value_exception ? 'FF8B6914' : 'FFB0201A');
      (l.route_types || []).forEach(function(rt2){
        var orders = (rt2.orders && rt2.orders.length) ? rt2.orders : [{ order_code:'', type:(rt2.types||[])[0]||'', value: rt2.value }];
        orders.forEach(function(ord){
          var row = rl.addRow([
            l.location_id,
            l.customer,
            l.address,
            rt2.route,
            ord.order_code || '—',
            ord.type || '',
            ord.value || 0,
            statusText,
            l.total_value || 0
          ]);
          rlRowCount++;
          row.getCell(7).numFmt = '#,##0';
          row.getCell(9).numFmt = '#,##0';
          row.getCell(8).fill = { type:'pattern', pattern:'solid', fgColor:{argb: statusColor} };
          row.getCell(8).font = { bold:true, color:{argb: fontColor} };
          row.getCell(9).font = { bold:true, color:{argb: fontColor} };
        });
      });
    });
    // Enable AutoFilter on the header row so the Status column (blue=Required, red=Avoidable,
    // amber=Exception) can be filtered directly in Excel — every row now carries its own
    // Status value, so filtering shows complete, self-contained rows, not blank gaps.
    rl.autoFilter = { from: { row: 1, column: 1 }, to: { row: rlRowCount, column: 9 } };

    var buf = await wb.xlsx.writeBuffer();
    res.setHeader('Content-Disposition', 'attachment; filename="Dispatch_Drop_Analysis_'+Date.now()+'.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(Buffer.from(buf));
  } catch(e) {
    console.error('dispatch export-excel error:', e.message);
    res.status(500).json({ error: 'Export failed: '+e.message });
  }
});

// ── SHARED PASSWORD CHECK FOR UPLOAD STARS ──
app.post('/api/backlog/check-password', requireAuth, async function(req, res) {
  try {
    var { password } = req.body;
    if (!password) return res.json({ ok: false });
    // Check against the logged-in user's own password
    var result = await pool.query('SELECT password_hash FROM users WHERE id=$1', [req.user.uid]);
    if (!result.rows[0]) return res.json({ ok: false });
    var match = await bcrypt.compare(password, result.rows[0].password_hash);
    res.json({ ok: match });
  } catch(e) { res.status(500).json({ ok: false, error: e.message }); }
});

// Caption must always have full DB access, regardless of what's loaded in the requesting browser's
// current session/tab. Pull directly from the server-side DB-backed globals every time.
function buildServerDataContext() {
  var parts = [];
  function addJSON(obj, label, maxLen) {
    try {
      if (obj === null || obj === undefined) { parts.push('=== ' + label + ': no file uploaded ==='); return; }
      var s = JSON.stringify(obj);
      if (s && s.length > maxLen) s = s.substring(0, maxLen) + '...(truncated)';
      parts.push('=== ' + label + ' (live from database) ===\n' + s);
    } catch(e) { parts.push('=== ' + label + ': error reading data ==='); }
  }

  addJSON(rejectionData ? { orgs: rejectionData.orgs, months: rejectionData.months } : null, 'REJECTION_DB', 4500);
  addJSON(currentDispatch ? { date: currentDispatch.date, summary: currentDispatch.summary } : null, 'DISPATCH_DB_LATEST', 3000);
  addJSON(Object.keys(dispatchHistory || {}).sort().reverse().slice(0, 14), 'DISPATCH_AVAILABLE_DATES', 500);
  addJSON(salesData ? salesData.summary : null, 'ORDER_BOOKING_DB (Todays Order Booking dashboard — BOOKED orders, NOT dispatch)', 3000);
  addJSON(returnsData ? returnsData.summary : null, 'RETURNS_DB', 3000);
  addJSON(backlogData ? backlogData.summary : null, 'BACKLOG_DB (WH Backlog)', 3000);
  addJSON(automationData ? { summary: automationData.summary, latestMonth: automationData.latestMonth, latestRate: automationData.latestRate, sortedMonths: automationData.sortedMonths } : null, 'AUTOMATION_DB', 2500);
  addJSON(deliveryData ? deliveryData.summary : null, 'DELIVERY_DB', 2000);
  addJSON(genInfoData ? genInfoData.rows : null, 'TEAM_DB', 3000);

  return parts.join('\n\n');
}

app.post('/api/voice', requireAuth, async function(req, res) {
  try {
    var text = req.body.text || '';
    var clientContext = req.body.context || '';
    var context = clientContext;
    try {
      context = clientContext + '\n\n' + buildServerDataContext();
    } catch(ctxErr) {
      console.error('buildServerDataContext failed, falling back to client-only context:', ctxErr.message);
    }
    var tab = req.body.tab || 'dispatch';
    var history = req.body.history || [];
    var lang = req.body.lang || 'English';

    var isFrederic = /i am fred|i.m fred|this is fred|hello.*fred|frederic here|i am frederic|frederic speaking/i.test(text.trim());

    var systemPrompt =
      'You are CAPTION, an operations assistant for AKI, a UAE logistics company. ' +
      'Built by Azhar (Mohammed Azharuddin, Customer Service and Operations at AKI). ' +
      'Azhar reports to Mr. Frederic Fleureau, GM Supply Chain and Operations Consumer at AKI. ' +
      'LANGUAGE: Respond ONLY in ' + lang + '. ' +
      (isFrederic ? 'FREDERIC MODE: Address as boss. Questions outside data: say Boss that is outside my scope. ' : '') +
      'CONVERSATION: Ongoing conversation. Remember everything. Answer follow-ups naturally. ' +
      'HOW TO READ THE DATA: ' +
      'SCREEN_NOW = exactly what the screen shows right now with all active filters. Use this ONLY when the user asks about "now" / "current view" / doesn\'t name a specific month. ' +
      'If the user names a SPECIFIC month (e.g. "June rejection details") and that month is NOT the one currently on screen, do NOT ask them to apply the filter — you already have every month\'s numbers in the MONTHLY section below. Read that month\'s data directly and answer immediately with the real numbers, in the SAME reply, every time. Never say "please apply the filter" or "I need to see that data" — you already have it. Set action=filter so the screen catches up to match your answer, but the spoken answer must stand on its own regardless of what the screen does. ' +
      'SCREEN_NOW tRej = TOTAL REJECTIONS on screen. ' +
      'SCREEN_NOW tDel = DELIVERED on screen. ' +
      'SCREEN_NOW contrib = CONTRIBUTION TO REJECTION RATE shown on screen. ' +
      'SCREEN_NOW val = VALUE AT RISK on screen. ' +
      'SCREEN_NOW foodRej = food type rejection count. nonFoodRej = non-food type rejection count. ' +
      'When food or nonfood type filter is active: DELIVERED on screen shows total YTD food/nonfood deliveries (large number like 57237 or 68265). ' +
      'For the actual June food delivered: add DCV.del[June] + DCF.del[June] from ByORG Monthly section. ' +
      'For actual June nonfood delivered: add DGC.del[June] + DGS.del[June] + DSN.del[June] from ByORG Monthly section. ' +
      'MONTHLY section = data for each month without day filter. ' +
      'Days section under each month = day-by-day breakdown. ' +
      'ByORG section = per ORG stats. ownRate = ORG rejection rate. contribBadge = small % shown on ORG card. ' +
      'RATE: overall rate = tRej/(tRej+tDel)*100. When type filter active use contrib from SCREEN_NOW. ' +
      'IMPORTANT: When day filter + type filter (food/nonfood) are both active: ' +
      'reason and customer counts are estimates scaled from all-type data. ' +
      'They may not add up exactly to tRej. Always state tRej as the exact total. ' +
      'Do NOT sum up reason counts and claim that is the total — use tRej from SCREEN_NOW. ' +
      'Keep answers 2 to 3 sentences. Lead with the actual number/answer in the first sentence — busy managers are listening, not reading, so do not warm up with preamble. If data missing say please upload the file. ' +
      'For dispatch/driver questions use AllDrivers section. ' +
      'Phone numbers: say plus then digits in groups. ' +
      'ORDER BOOKING vs DISPATCH — THESE ARE DIFFERENT THINGS, NEVER MIX THEM UP: ' +
      'ORDER_BOOKING section = orders that were BOOKED/PLACED (the "Today\'s Order Booking" dashboard). Use this whenever the user says "booking", "booked", "order booking", or "placed". ' +
      'DISPATCH section = orders that were physically DISPATCHED/sent out to customers already (a separate warehouse operation, often reflecting orders booked on a PREVIOUS day). Use this only when the user says "dispatch", "dispatched", or "delivery/route" questions. ' +
      'If the user asks "today\'s order booking" or "what did we book today", answer from ORDER_BOOKING, never from DISPATCH — booking and dispatch can be completely different numbers since today\'s dispatch is often yesterday\'s booking catching up. ' +
      'Set action=filter to filter dashboard. Set action=navigate to go to another dashboard. ' +
      'REJECTION FILTER: If the user asks about a specific month, ORG, food/non-food type, or external/internal source ' +
      'on the rejection dashboard, set action=filter and put the plain keywords in action_detail — ' +
      'e.g. action_detail="june non-food external" or action_detail="month=6 food". Always answer using the SAME period the user asked about, not the full year, unless they said "YTD" or "all months". ' +
      'SALON: Salon sales/rejections = the DGC org (label it "Salon" when the user says "salon"). ' +
      'CONTEXT SOURCES: the data below has two parts. The first part (SCREEN_NOW, MONTHLY, etc.) reflects exactly what is on the user\'s screen right now, including active filters — use it for "what am I looking at" style questions. ' +
      'The second part (sections ending in _DB, e.g. REJECTION_DB, DISPATCH_DB_LATEST, ORDER_BOOKING_DB, RETURNS_DB, BACKLOG_DB, AUTOMATION_DB, DELIVERY_DB, TEAM_DB) is pulled directly from the database every time and is ALWAYS complete and current, regardless of which dashboard tab the user currently has open or what they\'ve loaded this session. If a dashboard-specific section from the first part is missing or says "no file", fall back to the matching _DB section before ever saying data is unavailable — only say data is missing if the _DB section for that topic also says "no file uploaded". ' +
      'ALL DATA: ' + context.substring(0, 24000) +
      ' Reply ONLY in JSON: {"answer":"your answer","action":"none or filter or navigate","action_detail":"value","action_label":"label"}';

    var messages = [];
    var recent = history.slice(-16);
    for (var i = 0; i < recent.length; i++) {
      messages.push({ role: recent[i].role === 'assistant' ? 'assistant' : 'user', content: recent[i].content });
    }
    messages.push({ role: 'user', content: text });

    var msg = await anthropic.messages.create({
      model: 'claude-haiku-4-5-20251001',
      max_tokens: 600,
      system: systemPrompt,
      messages: messages
    });

    var raw = (msg.content[0].text || '').trim();
    var parsed;
    try {
      var m2 = raw.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(m2 ? m2[0] : raw);
    } catch(e) {
      parsed = { answer: raw, action: 'none', action_label: '' };
    }
    res.json({ success: true, result: parsed });
  } catch(e) {
    console.error('/api/voice error:', e.message, e.stack);
    res.status(500).json({ error: e.message });
  }
});

app.post('/api/chat', function(req, res) {
  try {
    var prompt=req.body.prompt, history=req.body.history||[];
    var messages=history.slice(-10).map(function(h){return{role:h.role==='assistant'?'assistant':'user',content:h.content};});
    if(!messages.length||messages[messages.length-1].content!==prompt) messages.push({role:'user',content:prompt});
    anthropic.messages.create({model:'claude-haiku-4-5-20251001',max_tokens:2000,system:'You are AZHAR-AI, a professional executive assistant for a UAE logistics company.',messages:messages})
      .then(function(msg){res.json({result:msg.content[0].text});})
      .catch(function(e){res.status(500).json({error:e.message});});
  } catch(e){if(!res.headersSent)res.status(500).json({error:e.message});}
});

app.post('/api/excel', upload.single('file'), function(req, res) {
  try {
    var question=req.body.question||'Analyse this data', dataText='';
    if(req.file){var ext2=path.extname(req.file.originalname||'').toLowerCase();dataText=(ext2==='.xlsx'||ext2==='.xls')?XLSX.utils.sheet_to_csv(XLSX.read(req.file.buffer,{type:'buffer'}).Sheets[XLSX.read(req.file.buffer,{type:'buffer'}).SheetNames[0]]):req.file.buffer.toString('utf8');}
    anthropic.messages.create({model:'claude-haiku-4-5-20251001',max_tokens:2000,messages:[{role:'user',content:question+(dataText?'\n\nData:\n'+dataText.substring(0,8000):'')}]})
      .then(function(msg){res.json({result:msg.content[0].text});})
      .catch(function(e){res.status(500).json({error:e.message});});
  } catch(e){if(!res.headersSent)res.status(500).json({error:e.message});}
});


// ─────────────────────────────────────────────
//  RETURNS DATA  (DB + file fallback)
// ─────────────────────────────────────────────
var returnsData = null;
var RETURNS_FILE = path.join(DATA_DIR, 'returns.json');

async function dbSaveReturns(uploadedBy, fileName, totalOrders, summary) {
  try {
    await pool.query('DELETE FROM returns_data');
    await pool.query(
      'INSERT INTO returns_data (uploaded_by, file_name, total_orders, summary) VALUES ($1, $2, $3, $4)',
      [uploadedBy, fileName, totalOrders, JSON.stringify(summary)]
    );
    return true;
  } catch(e) {
    console.error('DB save returns error:', e.message);
    return false;
  }
}

async function loadReturnsFromDB() {
  try {
    var res = await pool.query('SELECT * FROM returns_data ORDER BY uploaded_at DESC LIMIT 1');
    if (res.rows[0]) {
      returnsData = {
        uploadedAt: res.rows[0].uploaded_at,
        uploadedBy: res.rows[0].uploaded_by,
        fileName: res.rows[0].file_name,
        totalOrders: res.rows[0].total_orders,
        summary: res.rows[0].summary
      };
      console.log('Loaded returns from DB');
      return true;
    }
  } catch(e) { console.error('DB load returns:', e.message); }
  var saved = loadJSON(RETURNS_FILE);
  if (saved) { returnsData = saved; console.log('Loaded returns from file'); }
  return false;
}
loadReturnsFromDB();

app.post('/api/returns/upload', async function(req, res) {
  try {
    var summary = (typeof req.body.summary === 'string') ? JSON.parse(req.body.summary || '{}') : (req.body.summary || {});
    var uploadedBy = req.body.uploadedBy || 'Admin';
    var fileName = req.body.fileName || (req.file && req.file.originalname) || 'returns.csv';
    var totalOrders = parseInt(req.body.totalOrders) || 0;
    var rawRows = Array.isArray(req.body.rows) ? req.body.rows : [];
    returnsData = {
      uploadedAt: new Date().toISOString(),
      uploadedBy: uploadedBy,
      fileName: fileName,
      totalOrders: totalOrders,
      summary: summary
    };
    var dbOk = await dbSaveReturns(uploadedBy, fileName, totalOrders, summary);
    // Refresh row-level table so the Full Analyzed Report export has
    // live detail to compute from, not just the pre-aggregated summary.
    var rowsSaved = 0;
    if (rawRows.length) {
      try {
        await pool.query('DELETE FROM returns_rows');
        var chunkSize = 500;
        for (var ci = 0; ci < rawRows.length; ci += chunkSize) {
          var chunk = rawRows.slice(ci, ci + chunkSize);
          var vals = [], params = [], pi = 1;
          chunk.forEach(function(r) {
            vals.push('($' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ',$' + (pi++) + ')');
            params.push(
              toStr(r.wh || '').toUpperCase(),
              toStr(r.cat || ''),
              toStr(r.bu || ''),
              parseFloat(r.tot) || 0,
              parseInt(r.month) || 0,
              parseInt(r.day) || 0,
              toStr(r.dateStr || ''),
              toStr(r.cust || ''),
              toStr(r.area || ''),
              toStr(r.orderNo || ''),
              toStr(r.reason || '')
            );
          });
          await pool.query(
            'INSERT INTO returns_rows (warehouse, category, bu, value, month, day, date_str, customer, area, order_no, reason) VALUES ' + vals.join(','),
            params
          );
          rowsSaved += chunk.length;
        }
      } catch(rowErr) {
        console.error('Returns row-level save error:', rowErr.message);
      }
    }
    saveJSON(RETURNS_FILE, returnsData);
    var summaryKeys = Object.keys(summary);
    console.log('Returns saved:', totalOrders, 'orders, summary keys:', summaryKeys, dbOk ? '(DB+file)' : '(file only)', rowsSaved, 'raw rows stored');
    res.json({ success: true, rowsSaved: rowsSaved });
  } catch(e) {
    console.error('Returns upload error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: e.message });
  }
});

app.delete('/api/returns/clear', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM returns_data');
    await pool.query('DELETE FROM returns_rows');
    returnsData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

var RETURNS_SUMMARY_VERSION = 'v3';
app.get('/api/returns/status', function(req, res) {
  if (!returnsData) return res.json({ hasData: false });
  if (!returnsData.summary || returnsData.summary.version !== RETURNS_SUMMARY_VERSION) {
    return res.json({ hasData: false, reason: 'version_mismatch' });
  }
  res.json({
    hasData: true,
    uploadedAt: returnsData.uploadedAt,
    uploadedBy: returnsData.uploadedBy,
    fileName: returnsData.fileName,
    totalOrders: returnsData.totalOrders,
    summary: returnsData.summary
  });
});

// ── MARKET RETURNS: Full Analyzed Report (Excel) ──────────────────────────
// Mirrors the Customer Visit Prep export: multi-sheet ExcelJS workbook, dark
// gold theme, live SUMIF/COUNTIF/SUMIFS formulas reading off a raw "Return
// Master" tab (not frozen numbers) so the workbook re-totals itself if a row
// is corrected by hand. Built for the CEO-level ask: month-by-month returns
// by Food / Non-Food, in both % share and absolute AED, plus supporting
// detail (warehouse, category, area, return reason) and the full row-level
// data behind it.
app.get('/api/returns/export-excel', requireAuth, async function(req, res) {
  try {
    var rowsRes = await pool.query('SELECT * FROM returns_rows ORDER BY month, day');
    var rows = rowsRes.rows || [];
    if (!rows.length) return res.status(404).json({ error: 'No Market Returns row-level data on file. Re-upload the returns file to enable this export.' });

    var MONTH_NAMES = ['', 'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    var genDate = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

    // Aggregate for the Executive Summary / sheet ordering (formulas do the
    // real math in the workbook; these are just used to decide row order).
    var totalVal = 0, invVal = 0, grvVal = 0, invCount = 0, grvCount = 0;
    var foodVal = 0, nfVal = 0;
    var whMap = {}, monthSet = {}, areaMap = {}, reasonMap = {};
    rows.forEach(function(r) {
      var v = parseFloat(r.value) || 0;
      totalVal += v;
      if (r.category === 'Invoice Cancellation') { invVal += v; invCount++; }
      else if (r.category === 'GRV-Return') { grvVal += v; grvCount++; }
      if (r.bu === 'Food') foodVal += v;
      else if (r.bu === 'Non-Food') nfVal += v;
      if (r.warehouse) whMap[r.warehouse] = (whMap[r.warehouse] || 0) + v;
      if (r.month) monthSet[r.month] = true;
      if (r.area) areaMap[r.area] = (areaMap[r.area] || 0) + v;
      if (r.reason) reasonMap[r.reason] = (reasonMap[r.reason] || 0) + v;
    });
    var months = Object.keys(monthSet).map(Number).sort(function(a, b) { return a - b; });
    var whKeys = Object.keys(whMap).sort(function(a, b) { return whMap[b] - whMap[a]; });
    var areaKeys = Object.keys(areaMap).sort(function(a, b) { return areaMap[b] - areaMap[a]; });
    var reasonKeys = Object.keys(reasonMap).sort(function(a, b) { return reasonMap[b] - reasonMap[a]; });

    var GOLD = 'FFC9A84C', DARKBG = 'FF1A1E26', LIGHTGOLD = 'FFF5E9C8';
    function styleHeaderRow(row) {
      row.eachCell(function(cell) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } };
        cell.font = { bold: true, color: { argb: GOLD }, size: 11 };
        cell.alignment = { vertical: 'middle' };
      });
    }
    function styleSectionRow(row) {
      row.font = { bold: true, color: { argb: DARKBG }, size: 12 };
      row.eachCell(function(cell) { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GOLD } }; });
    }
    function styleTotalRow(row) {
      row.eachCell(function(cell) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHTGOLD } };
        cell.font = { bold: true, color: { argb: DARKBG } };
      });
    }
    function styleTitleRow(row) { row.font = { bold: true, size: 13, color: { argb: GOLD } }; row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: DARKBG } }; }
    function noteRow(sheet, text, span) {
      var r = sheet.addRow([text]);
      r.font = { italic: true, size: 9, color: { argb: 'FF666666' } };
      sheet.mergeCells('A' + r.number + ':' + span + r.number);
      return r;
    }

    var RM_START = 3;
    var RM_END = 2 + rows.length;
    var RMS = "'Return Master'!";
    var R_MONTH = RMS + '$A$' + RM_START + ':$A$' + RM_END;
    var R_WH = RMS + '$C$' + RM_START + ':$C$' + RM_END;
    var R_CAT = RMS + '$D$' + RM_START + ':$D$' + RM_END;
    var R_BU = RMS + '$E$' + RM_START + ':$E$' + RM_END;
    var R_VAL = RMS + '$F$' + RM_START + ':$F$' + RM_END;
    var R_AREA = RMS + '$G$' + RM_START + ':$G$' + RM_END;
    var R_REASON = RMS + '$I$' + RM_START + ':$I$' + RM_END;

    var ExcelJS = require('exceljs');
    var wb = new ExcelJS.Workbook();
    wb.creator = 'AZHAR-AI'; wb.created = new Date();

    // ══════════════════════ Sheet 1: Executive Summary ══════════════════════
    var ex = wb.addWorksheet('Executive Summary');
    ex.columns = [{ width: 26 }, { width: 16 }, { width: 14 }, { width: 16 }, { width: 14 }, { width: 26 }, { width: 14 }, { width: 14 }];
    var exTitle = ex.addRow(['MARKET RETURNS — FULL ANALYZED REPORT']); styleTitleRow(exTitle); ex.mergeCells('A1:H1');
    var exSub = ex.addRow(['AKI Group  —  Invoice Cancellations · GRV Returns · BU-wise  —  Generated ' + genDate]);
    exSub.font = { bold: true, size: 12, color: { argb: DARKBG } }; ex.mergeCells('A2:H2');
    ex.addRow([]);

    styleSectionRow(ex.addRow(['KEY METRICS']));
    var kpiHdr = ex.addRow(['Total Returns Value (AED)', '', 'Invoice Cancellations', '', 'GRV Returns', '', 'Total Orders', '']);
    kpiHdr.font = { bold: true, size: 10, color: { argb: DARKBG } };
    var kpiVal = ex.addRow([
      { formula: 'SUM(' + R_VAL + ')' }, '',
      { formula: 'COUNTIF(' + R_CAT + ',"Invoice Cancellation")' }, '',
      { formula: 'COUNTIF(' + R_CAT + ',"GRV-Return")' }, '',
      { formula: 'COUNTA(' + R_MONTH + ')' }, ''
    ]);
    kpiVal.font = { bold: true, size: 16, color: { argb: DARKBG } };
    kpiVal.getCell(1).numFmt = '#,##0'; kpiVal.getCell(1).font = { bold: true, size: 16, color: { argb: 'FFB00000' } };
    ex.addRow([]);

    styleSectionRow(ex.addRow(['FOOD vs NON-FOOD — OVERALL SPLIT']));
    var buHdr = ex.addRow(['Category', 'Value (AED)', '% of Total']); styleHeaderRow(buHdr);
    var buFirstRow = ex.rowCount + 1;
    ['Food', 'Non-Food'].forEach(function(b) {
      var r0 = ex.addRow([b, { formula: 'SUMIF(' + R_BU + ',A' + (ex.rowCount + 1) + ',' + R_VAL + ')' }, '']);
      r0.getCell(2).numFmt = '#,##0';
    });
    var buLastRow = ex.rowCount;
    for (var bi = buFirstRow; bi <= buLastRow; bi++) {
      ex.getCell('C' + bi).value = { formula: 'IFERROR(B' + bi + '/$B$' + kpiVal.number + ',0)' };
      ex.getCell('C' + bi).numFmt = '0.0%';
    }
    var buTotalRow = ex.addRow(['TOTAL', { formula: 'SUM(B' + buFirstRow + ':B' + buLastRow + ')' }, { formula: 'SUM(C' + buFirstRow + ':C' + buLastRow + ')' }]);
    buTotalRow.getCell(2).numFmt = '#,##0'; buTotalRow.getCell(3).numFmt = '0.0%';
    styleTotalRow(buTotalRow);
    ex.addRow([]);

    styleSectionRow(ex.addRow(['BY CATEGORY']));
    var catHdr = ex.addRow(['Category', 'Orders', 'Value (AED)', '% of Total']); styleHeaderRow(catHdr);
    var catFirstRow = ex.rowCount + 1;
    ['Invoice Cancellation', 'GRV-Return'].forEach(function(c) {
      var r0 = ex.addRow([c,
        { formula: 'COUNTIF(' + R_CAT + ',A' + (ex.rowCount + 1) + ')' },
        { formula: 'SUMIF(' + R_CAT + ',A' + (ex.rowCount + 1) + ',' + R_VAL + ')' }, '']);
      r0.getCell(3).numFmt = '#,##0';
    });
    var catLastRow = ex.rowCount;
    for (var cci = catFirstRow; cci <= catLastRow; cci++) {
      ex.getCell('D' + cci).value = { formula: 'IFERROR(C' + cci + '/$B$' + kpiVal.number + ',0)' };
      ex.getCell('D' + cci).numFmt = '0.0%';
    }
    noteRow(ex, 'All figures on this tab are calculated live from the Return Master tab and update automatically if any row is corrected.', 'H');

    // ══════════════════════ Sheet 2: Monthly Trend (Food vs Non-Food) ══════════════════════
    var mt = wb.addWorksheet('Monthly Trend');
    mt.columns = [{ width: 10 }, { width: 16 }, { width: 12 }, { width: 16 }, { width: 12 }, { width: 16 }, { width: 12 }, { width: 12 }];
    var mtTitle = mt.addRow(['MONTH-BY-MONTH RETURNS — FOOD vs NON-FOOD']); styleTitleRow(mtTitle); mt.mergeCells('A1:H1');
    mt.addRow([]);
    styleSectionRow(mt.addRow(['VALUE (AED) AND % SHARE BY MONTH']));
    styleHeaderRow(mt.addRow(['Month', 'Food (AED)', 'Food %', 'Non-Food (AED)', 'Non-Food %', 'Total (AED)', 'Orders', 'MoM Change']));
    var mtFirstRow = mt.rowCount + 1;
    months.forEach(function(m, mi) {
      var mLabel = MONTH_NAMES[m] || m;
      var r0 = mt.addRow([mLabel,
        { formula: 'SUMIFS(' + R_VAL + ',' + R_MONTH + ',' + m + ',' + R_BU + ',"Food")' }, '',
        { formula: 'SUMIFS(' + R_VAL + ',' + R_MONTH + ',' + m + ',' + R_BU + ',"Non-Food")' }, '',
        { formula: 'SUMIF(' + R_MONTH + ',' + m + ',' + R_VAL + ')' },
        { formula: 'COUNTIF(' + R_MONTH + ',' + m + ')' }, '']);
      r0.getCell(2).numFmt = '#,##0'; r0.getCell(4).numFmt = '#,##0'; r0.getCell(6).numFmt = '#,##0';
      r0.getCell(3).value = { formula: 'IFERROR(B' + r0.number + '/F' + r0.number + ',0)' }; r0.getCell(3).numFmt = '0.0%';
      r0.getCell(5).value = { formula: 'IFERROR(D' + r0.number + '/F' + r0.number + ',0)' }; r0.getCell(5).numFmt = '0.0%';
      if (mi === 0) { r0.getCell(8).value = '—'; }
      else { r0.getCell(8).value = { formula: 'IFERROR((F' + r0.number + '-F' + (r0.number - 1) + ')/F' + (r0.number - 1) + ',0)' }; r0.getCell(8).numFmt = '0.0%'; }
    });
    var mtLastRow = mt.rowCount;
    if (months.length) {
      var mtTotalRow = mt.addRow(['TOTAL',
        { formula: 'SUM(B' + mtFirstRow + ':B' + mtLastRow + ')' }, { formula: 'IFERROR(B' + (mtLastRow + 1) + '/F' + (mtLastRow + 1) + ',0)' },
        { formula: 'SUM(D' + mtFirstRow + ':D' + mtLastRow + ')' }, { formula: 'IFERROR(D' + (mtLastRow + 1) + '/F' + (mtLastRow + 1) + ',0)' },
        { formula: 'SUM(F' + mtFirstRow + ':F' + mtLastRow + ')' },
        { formula: 'SUM(G' + mtFirstRow + ':G' + mtLastRow + ')' }, '']);
      mtTotalRow.getCell(2).numFmt = '#,##0'; mtTotalRow.getCell(3).numFmt = '0.0%';
      mtTotalRow.getCell(4).numFmt = '#,##0'; mtTotalRow.getCell(5).numFmt = '0.0%';
      mtTotalRow.getCell(6).numFmt = '#,##0';
      styleTotalRow(mtTotalRow);
    }
    noteRow(mt, 'All figures calculated live from the Return Master tab (SUMIFS/COUNTIF by Month and BU). The most recent month may be partial — data on file through ' + genDate + '.', 'H');

    // ══════════════════════ Sheet 3: By Warehouse ══════════════════════
    var wh = wb.addWorksheet('By Warehouse');
    wh.columns = [{ width: 14 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 14 }, { width: 14 }];
    var whTitle = wh.addRow(['RETURNS BY WAREHOUSE']); styleTitleRow(whTitle); wh.mergeCells('A1:F1');
    wh.addRow([]);
    styleHeaderRow(wh.addRow(['Warehouse', 'Orders', 'Invoice Cancel', 'GRV Return', 'Value (AED)', '% of Total']));
    var whFirstRow = wh.rowCount + 1;
    whKeys.forEach(function(w) {
      var r0 = wh.addRow([w,
        { formula: 'COUNTIF(' + R_WH + ',A' + (wh.rowCount + 1) + ')' },
        { formula: 'COUNTIFS(' + R_WH + ',A' + (wh.rowCount + 1) + ',' + R_CAT + ',"Invoice Cancellation")' },
        { formula: 'COUNTIFS(' + R_WH + ',A' + (wh.rowCount + 1) + ',' + R_CAT + ',"GRV-Return")' },
        { formula: 'SUMIF(' + R_WH + ',A' + (wh.rowCount + 1) + ',' + R_VAL + ')' }, '']);
      r0.getCell(5).numFmt = '#,##0';
    });
    var whLastRow = wh.rowCount;
    for (var whi = whFirstRow; whi <= whLastRow; whi++) {
      wh.getCell('F' + whi).value = { formula: 'IFERROR(E' + whi + '/SUM($E$' + whFirstRow + ':$E$' + whLastRow + '),0)' };
      wh.getCell('F' + whi).numFmt = '0.0%';
    }
    var whTotalRow = wh.addRow(['TOTAL',
      { formula: 'SUM(B' + whFirstRow + ':B' + whLastRow + ')' },
      { formula: 'SUM(C' + whFirstRow + ':C' + whLastRow + ')' },
      { formula: 'SUM(D' + whFirstRow + ':D' + whLastRow + ')' },
      { formula: 'SUM(E' + whFirstRow + ':E' + whLastRow + ')' }, 1]);
    whTotalRow.getCell(5).numFmt = '#,##0'; whTotalRow.getCell(6).numFmt = '0.0%';
    styleTotalRow(whTotalRow);
    noteRow(wh, 'Calculated live from the Return Master tab.', 'F');

    // ══════════════════════ Sheet 4: By Area & Reason ══════════════════════
    var ar = wb.addWorksheet('Area & Reason');
    ar.columns = [{ width: 14 }, { width: 14 }, { width: 12 }, { width: 4 }, { width: 34 }, { width: 12 }, { width: 12 }];
    var arTitle = ar.addRow(['RETURNS BY EMIRATE / AREA  &  BY RETURN REASON']); styleTitleRow(arTitle); ar.mergeCells('A1:G1');
    ar.addRow([]);
    var arHdr = ar.addRow(['Area', 'Value (AED)', '% of Total', '', 'Return Reason', 'Value (AED)', '% of Total']); styleHeaderRow(arHdr);
    var arRowsN = Math.max(areaKeys.length, reasonKeys.length);
    var arFirstRow = ar.rowCount + 1;
    for (var ai = 0; ai < arRowsN; ai++) {
      var rowArr = ['', '', '', '', '', '', ''];
      if (areaKeys[ai]) rowArr[0] = areaKeys[ai];
      if (reasonKeys[ai]) rowArr[4] = reasonKeys[ai];
      var r0 = ar.addRow(rowArr);
      if (areaKeys[ai]) {
        r0.getCell(2).value = { formula: 'SUMIF(' + R_AREA + ',A' + r0.number + ',' + R_VAL + ')' }; r0.getCell(2).numFmt = '#,##0';
        r0.getCell(3).value = { formula: 'IFERROR(B' + r0.number + '/SUM(' + R_VAL + '),0)' }; r0.getCell(3).numFmt = '0.0%';
      }
      if (reasonKeys[ai]) {
        r0.getCell(6).value = { formula: 'SUMIF(' + R_REASON + ',E' + r0.number + ',' + R_VAL + ')' }; r0.getCell(6).numFmt = '#,##0';
        r0.getCell(7).value = { formula: 'IFERROR(F' + r0.number + '/SUM(' + R_VAL + '),0)' }; r0.getCell(7).numFmt = '0.0%';
      }
    }
    noteRow(ar, 'Calculated live from the Return Master tab. Reason column depends on the source upload including a Return Reason field.', 'G');

    // ══════════════════════ Sheet 5: Return Master (raw data) ══════════════════════
    var rm = wb.addWorksheet('Return Master');
    rm.columns = [{ width: 8 }, { width: 6 }, { width: 12 }, { width: 20 }, { width: 12 }, { width: 14 }, { width: 14 }, { width: 16 }, { width: 34 } ];
    var rmTitle = rm.addRow(['RETURN MASTER DATA (' + rows.length + ' rows)']); styleTitleRow(rmTitle); rm.mergeCells('A1:I1');
    styleHeaderRow(rm.addRow(['Month', 'Day', 'Warehouse', 'Category', 'BU', 'Value (AED)', 'Area', 'Order No', 'Return Reason']));
    rows.forEach(function(r) {
      var r0 = rm.addRow([r.month, r.day, r.warehouse, r.category, r.bu, Math.round(parseFloat(r.value) || 0), r.area || '—', r.order_no || '—', r.reason || '—']);
      r0.getCell(6).numFmt = '#,##0';
    });

    var fileName = 'Market_Returns_Full_Analyzed_Report_' + new Date().toISOString().slice(0, 10) + '.xlsx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="' + fileName + '"');
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error('Returns export-excel error:', e.message);
    if (!res.headersSent) res.status(500).json({ error: 'Failed to build Excel export' });
  }
});

// ─── AUTH SYSTEM ──────────────────────────────────────────────────────────

// Audit log helper
async function auditLog(userId, username, action, details, ip) {
  try {
    await pool.query(
      'INSERT INTO audit_log (user_id, username, action, details, ip_address) VALUES ($1,$2,$3,$4,$5)',
      [userId||null, username||'system', action, details||'', ip||'']
    );
  } catch(e) { console.error('Audit log error:', e.message); }
}

// Auth middleware
async function requireAuth(req, res, next) {
  var token = req.headers['x-auth-token'] || req.headers['authorization'];
  if (token && token.startsWith('Bearer ')) token = token.slice(7);
  if (!token) return res.status(401).json({ error: 'Not authenticated' });
  try {
    var sess = await pool.query(
      'SELECT s.*, u.id as uid, u.username, u.role, u.dashboards, u.full_name, u.active, u.must_change_password, u.horeca_salesperson_name FROM sessions s JOIN users u ON u.id=s.user_id WHERE s.token=$1 AND s.expires_at>NOW()',
      [token]
    );
    if (!sess.rows[0]) return res.status(401).json({ error: 'Session expired' });
    if (!sess.rows[0].active) return res.status(403).json({ error: 'Account disabled' });
    req.user = sess.rows[0];
    next();
  } catch(e) { res.status(500).json({ error: e.message }); }
}

function requireRole(...roles) {
  return function(req, res, next) {
    if (!req.user) return res.status(401).json({ error: 'Not authenticated' });
    if (!roles.includes(req.user.role)) return res.status(403).json({ error: 'Access denied' });
    next();
  };
}

// ── HoReCa Order Module (new tables, shares this pool/auth, isolated from other dashboards) ──
require('./horeca_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);

// ── PPT Polish Module (upload a rough deck, get it improved via Claude, keeping original images) ──
require('./ppt_polish_module')(app, requireAuth, upload, anthropic, auditLog);

// ── Brand Presentations Module (CEO-level brand decks) ──
require('./brand_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);

// ── Vehicle Master Module (fleet registry, used as fallback for transport cost estimate) ──
var VEHICLE_MASTER_MAP = {};
require('./vehicle_master_module')(app, pool, requireAuth, requireRole, upload, auditLog, VEHICLE_MASTER_MAP);

// ── Aujan Pallet Collection & Recovery Tracking Module (Phase 1) ──
require('./pallet_module')(app, pool, requireAuth, requireRole, upload, auditLog);
require('./cs_workload_module')(app, pool, requireAuth, requireRole, upload, auditLog, anthropic);
require('./price_discrepancy_module')(app, pool, requireAuth, requireRole, upload, auditLog);
require('./task_tracker_module')(app, pool, requireAuth, requireRole, auditLog);
require('./grv_return_module')(app, pool, requireAuth, requireRole, upload, auditLog);
require('./trip_module')(app, pool, requireAuth, requireRole, upload, auditLog, bcrypt);
require('./transport_cost_module')(app, pool, requireAuth, requireRole);

// ── LOGIN ──
app.post('/api/auth/login', loginLimiter, async function(req, res) {
  try {
    var { username, password } = req.body;
    if (!username || !password) return res.status(400).json({ error: 'Username and password required' });
    var result = await pool.query('SELECT * FROM users WHERE username=$1', [username.toLowerCase().trim()]);
    var user = result.rows[0];
    if (!user) return res.status(401).json({ error: 'Invalid username or password' });
    if (!user.active) return res.status(403).json({ error: 'Account is disabled. Contact admin.' });
    var match = await bcrypt.compare(password, user.password_hash);
    if (!match) return res.status(401).json({ error: 'Invalid username or password' });
    // Create session token
    var token = crypto.randomBytes(32).toString('hex');
    var expires = new Date(Date.now() + 8 * 60 * 60 * 1000); // 8 hours
    var ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || '';
    await pool.query('INSERT INTO sessions (token, user_id, expires_at, ip_address) VALUES ($1,$2,$3,$4)',
      [token, user.id, expires, ip]);
    // Update last login
    await pool.query('UPDATE users SET last_login=NOW() WHERE id=$1', [user.id]);
    // Audit
    await auditLog(user.id, user.username, 'LOGIN', 'Successful login', ip);
    res.json({
      success: true,
      token: token,
      user: {
        id: user.id,
        username: user.username,
        full_name: user.full_name,
        role: user.role,
        dashboards: user.dashboards,
        must_change_password: user.must_change_password
      }
    });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── LOGOUT ──
app.post('/api/auth/logout', requireAuth, async function(req, res) {
  try {
    var token = req.headers['x-auth-token'] || (req.headers['authorization']||'').replace('Bearer ','');
    await pool.query('DELETE FROM sessions WHERE token=$1', [token]);
    await auditLog(req.user.uid, req.user.username, 'LOGOUT', '', req.headers['x-forwarded-for']||'');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── GET ME ──
app.get('/api/auth/me', requireAuth, function(req, res) {
  res.json({
    id: req.user.uid,
    username: req.user.username,
    full_name: req.user.full_name,
    role: req.user.role,
    dashboards: req.user.dashboards,
    must_change_password: req.user.must_change_password
  });
});

// ── CHANGE PASSWORD ──
app.post('/api/auth/change-password', requireAuth, async function(req, res) {
  try {
    var { current_password, new_password } = req.body;
    if (!new_password || new_password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
    var result = await pool.query('SELECT password_hash FROM users WHERE id=$1', [req.user.uid]);
    var match = await bcrypt.compare(current_password, result.rows[0].password_hash);
    if (!match) return res.status(401).json({ error: 'Current password incorrect' });
    var hash = await bcrypt.hash(new_password, 10);
    await pool.query('UPDATE users SET password_hash=$1, must_change_password=false WHERE id=$2', [hash, req.user.uid]);
    await auditLog(req.user.uid, req.user.username, 'CHANGE_PASSWORD', 'Password changed by user', '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── USER MANAGEMENT (superadmin only) ──
app.get('/api/users', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var result = await pool.query('SELECT id, username, full_name, role, dashboards, active, created_at, last_login, must_change_password, upload_exempt_until FROM users ORDER BY created_at');
    res.json({ users: result.rows });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/users', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var { username, password, full_name, role, dashboards } = req.body;
    if (!username || !password) return res.status(400).json({ error: 'Username and password required' });
    var hash = await bcrypt.hash(password, 10);
    var dbs = dashboards || ['dispatch','rejection','summary','email','invoice','backlog','returns','sales','automation'];
    var result = await pool.query(
      'INSERT INTO users (username, password_hash, full_name, role, dashboards, created_by) VALUES ($1,$2,$3,$4,$5,$6) RETURNING id',
      [username.toLowerCase().trim(), hash, full_name||username, role||'viewer', JSON.stringify(dbs), req.user.username]
    );
    await auditLog(req.user.uid, req.user.username, 'CREATE_USER', 'Created user: '+username+' role: '+role, '');
    res.json({ success: true, id: result.rows[0].id });
  } catch(e) {
    if (e.message.includes('unique')) return res.status(400).json({ error: 'Username already exists' });
    res.status(500).json({ error: e.message });
  }
});

app.put('/api/users/:id', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var { full_name, role, dashboards, active } = req.body;
    var dbs = dashboards ? JSON.stringify(dashboards) : null;
    await pool.query(
      'UPDATE users SET full_name=COALESCE($1,full_name), role=COALESCE($2,role), dashboards=COALESCE($3::jsonb,dashboards), active=COALESCE($4,active) WHERE id=$5',
      [full_name||null, role||null, dbs, active!=null?active:null, req.params.id]
    );
    await auditLog(req.user.uid, req.user.username, 'UPDATE_USER', 'Updated user ID: '+req.params.id, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Temporarily lifts the sub-admin 2-per-10-hour upload cooldown for one user, for a
// given number of days — e.g. someone catching up on several missing days at once
// (their own uploads still go through normal validation; this only removes the
// rate-limit that would otherwise block them after 2 uploads in a 10-hour window).
// Passing days=0 (or omitting it) revokes an active exemption immediately.
app.post('/api/users/:id/upload-exemption', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var days = Math.max(0, parseInt(req.body.days, 10) || 0);
    var untilVal = days > 0 ? new Date(Date.now() + days * 86400000) : null;
    var r = await pool.query('UPDATE users SET upload_exempt_until=$1 WHERE id=$2 RETURNING username', [untilVal, req.params.id]);
    if (!r.rows[0]) return res.status(404).json({ error: 'User not found' });
    var msg = days > 0
      ? 'Granted ' + r.rows[0].username + ' an upload-cooldown exemption for ' + days + ' day(s), until ' + untilVal.toLocaleString('en-AE')
      : 'Revoked upload-cooldown exemption for ' + r.rows[0].username;
    await auditLog(req.user.uid, req.user.username, 'UPLOAD_EXEMPTION', msg, req.headers['x-forwarded-for'] || req.ip || '');
    res.json({ success: true, upload_exempt_until: untilVal });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/users/:id', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    if (parseInt(req.params.id) === req.user.uid) return res.status(400).json({ error: 'Cannot delete your own account' });
    var u = await pool.query('SELECT username FROM users WHERE id=$1', [req.params.id]);
    await pool.query('DELETE FROM users WHERE id=$1', [req.params.id]);
    await auditLog(req.user.uid, req.user.username, 'DELETE_USER', 'Deleted user: '+(u.rows[0]||{}).username, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/users/:id/reset-password', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var { new_password } = req.body;
    if (!new_password || new_password.length < 6) return res.status(400).json({ error: 'Password must be at least 6 characters' });
    var hash = await bcrypt.hash(new_password, 10);
    await pool.query('UPDATE users SET password_hash=$1, must_change_password=true WHERE id=$2', [hash, req.params.id]);
    var u = await pool.query('SELECT username FROM users WHERE id=$1', [req.params.id]);
    await auditLog(req.user.uid, req.user.username, 'RESET_PASSWORD', 'Reset password for: '+(u.rows[0]||{}).username, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ── EMERGENCY ADMIN RESET (remove after first use) ──
app.get('/api/version', function(req, res) {
  res.json({ version: 'V4.1', date: '2026-06-05', status: 'running', auth: 'active' });
});

app.get('/api/setup/reset-admin', async function(req, res) {
  try {
    var hash = await bcrypt.hash('YAmaha100@', 10);
    var check = await pool.query("SELECT id FROM users WHERE username='azhar'");
    if (check.rows.length === 0) {
      await pool.query(
        "INSERT INTO users (username, password_hash, full_name, role, active) VALUES ($1,$2,$3,$4,true)",
        ['azhar', hash, 'Mohammed Azharuddin', 'superadmin']
      );
      res.json({ success: true, message: 'Admin user CREATED. Login: azhar / YAmaha100@' });
    } else {
      // Force reset — update password AND ensure active=true
      await pool.query("UPDATE users SET password_hash=$1, active=true, must_change_password=false WHERE username='azhar'", [hash]);
      // Clear all old sessions for this user
      var uid = check.rows[0].id;
      await pool.query("DELETE FROM sessions WHERE user_id=$1", [uid]);
      res.json({ success: true, message: 'Password RESET. All sessions cleared. Login: azhar / YAmaha100@' });
    }
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── AUDIT LOG ──
app.get('/api/audit', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    var result = await pool.query('SELECT * FROM audit_log ORDER BY created_at DESC LIMIT 500');
    res.json({ logs: result.rows });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── AUDIT UPLOAD ACTIONS ──
// Patch existing upload endpoints to log actions
// (handled via middleware injection in each upload route)

// STATIC - MUST BE LAST
app.get('/', function(req, res) {
  var p1=path.join(__dirname,'public','index.html'), p2=path.join(__dirname,'index.html'), p3=path.join(__dirname,'azhar-ai-v4.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  if(fs.existsSync(p3))return res.sendFile(p3);
  res.status(404).json({error:'index.html not found'});
});
app.get('/orders', function(req, res) {
  var p1=path.join(__dirname,'public','orders.html'), p2=path.join(__dirname,'orders.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'orders.html not found'});
});
app.get('/brands', function(req, res) {
  var p1=path.join(__dirname,'public','brand_frontend.html'), p2=path.join(__dirname,'brand_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'brand_frontend.html not found'});
});
app.get('/pallets', function(req, res) {
  var p1=path.join(__dirname,'public','pallet_frontend.html'), p2=path.join(__dirname,'pallet_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'pallet_frontend.html not found'});
});
app.get('/transport-cost', function(req, res) {
  var p1=path.join(__dirname,'public','transport_cost_frontend.html'), p2=path.join(__dirname,'transport_cost_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'transport_cost_frontend.html not found'});
});
app.get('/cs-workload', function(req, res) {
  var p1=path.join(__dirname,'public','cs_workload_frontend.html'), p2=path.join(__dirname,'cs_workload_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'cs_workload_frontend.html not found'});
});
app.get('/price-discrepancy', function(req, res) {
  var p1=path.join(__dirname,'public','price_discrepancy_frontend.html'), p2=path.join(__dirname,'price_discrepancy_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'price_discrepancy_frontend.html not found'});
});
app.get('/tasks-dashboard', function(req, res) {
  var p1=path.join(__dirname,'public','task_tracker_frontend.html'), p2=path.join(__dirname,'task_tracker_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'task_tracker_frontend.html not found'});
});
app.get('/grv-return', function(req, res) {
  var p1=path.join(__dirname,'public','grv_return_frontend.html'), p2=path.join(__dirname,'grv_return_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'grv_return_frontend.html not found'});
});
app.get('/grv-cs', function(req, res) {
  var p1=path.join(__dirname,'public','grv_cs_frontend.html'), p2=path.join(__dirname,'grv_cs_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'grv_cs_frontend.html not found'});
});
app.get('/trip', function(req, res) {
  var p1=path.join(__dirname,'public','trip_frontend.html'), p2=path.join(__dirname,'trip_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'trip_frontend.html not found'});
});
app.get('/trip-login', function(req, res) {
  var p1=path.join(__dirname,'public','trip_login_frontend.html'), p2=path.join(__dirname,'trip_login_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'trip_login_frontend.html not found'});
});
app.get('/trip-portal', function(req, res) {
  var p1=path.join(__dirname,'public','trip_portal_frontend.html'), p2=path.join(__dirname,'trip_portal_frontend.html');
  if(fs.existsSync(p1))return res.sendFile(p1);
  if(fs.existsSync(p2))return res.sendFile(p2);
  res.status(404).json({error:'trip_portal_frontend.html not found'});
});
app.use(express.static(path.join(__dirname,'public')));
app.use(express.static(__dirname));

app.use(function(err,req,res,next){
  console.error('Global error:',err.message);
  if(!res.headersSent)res.status(500).json({error:err.message||'Server error'});
});

// ─── DAILY SALES ──────────────────────────────────────────
var salesData = null;

async function loadSalesFromDB() {
  try {
    await pool.query('CREATE TABLE IF NOT EXISTS sales_data (id SERIAL PRIMARY KEY, uploaded_at TIMESTAMPTZ DEFAULT NOW(), uploaded_by TEXT, file_name TEXT, total_orders INT, summary JSONB)');
    var res = await pool.query('SELECT * FROM sales_data ORDER BY uploaded_at DESC LIMIT 1');
    if (res.rows[0]) {
      salesData = { uploadedAt: res.rows[0].uploaded_at, fileName: res.rows[0].file_name, totalOrders: res.rows[0].total_orders, summary: res.rows[0].summary };
      console.log('Loaded sales from DB');
    }
  } catch(e) { console.error('DB load sales:', e.message); }
}
loadSalesFromDB();

app.delete('/api/sales/clear', requireAuth, requireRole('superadmin'), async function(req, res) {
  try { await pool.query('DELETE FROM sales_data'); salesData = null; res.json({ success: true }); }
  catch(e) { res.status(500).json({ error: e.message }); }
});

app.get('/api/sales/status', function(req, res) {
  if (!salesData) return res.json({ hasData: false });
  if (!salesData.summary || salesData.summary.version !== 'v1') return res.json({ hasData: false });
  res.json({ hasData: true, uploadedAt: salesData.uploadedAt, fileName: salesData.fileName, totalOrders: salesData.totalOrders, summary: salesData.summary });
});

app.post('/api/sales/upload', async function(req, res) {
  try {
    var summary = (typeof req.body.summary === 'string') ? JSON.parse(req.body.summary) : (req.body.summary || {});
    var fileName = req.body.fileName || 'sales.xlsx';
    var totalOrders = parseInt(req.body.totalOrders) || 0;
    salesData = { uploadedAt: new Date(), fileName: fileName, totalOrders: totalOrders, summary: summary };
    try {
      await pool.query('CREATE TABLE IF NOT EXISTS sales_data (id SERIAL PRIMARY KEY, uploaded_at TIMESTAMPTZ DEFAULT NOW(), uploaded_by TEXT, file_name TEXT, total_orders INT, summary JSONB)');
      await pool.query('DELETE FROM sales_data');
      await pool.query('INSERT INTO sales_data (uploaded_by, file_name, total_orders, summary) VALUES ($1,$2,$3,$4)', ['Admin', fileName, totalOrders, JSON.stringify(summary)]);
    } catch(dbErr) { console.error('Sales DB save:', dbErr.message); }
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ─── GENERAL INFO ──────────────────────────────────────────────────────────
var genInfoData = null;
async function loadGenInfoFromDB() {
  try {
    await pool.query('CREATE TABLE IF NOT EXISTS geninfo_data (id SERIAL PRIMARY KEY, uploaded_at TIMESTAMPTZ DEFAULT NOW(), file_name TEXT, total_members INT, rows JSONB)');
    await pool.query(`CREATE TABLE IF NOT EXISTS automation_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_records INT,
      summary JSONB,
      rows JSONB
    )`);
    var r = await pool.query('SELECT * FROM geninfo_data ORDER BY uploaded_at DESC LIMIT 1');
    if (r.rows[0]) {
      genInfoData = { fileName: r.rows[0].file_name, totalMembers: r.rows[0].total_members, rows: r.rows[0].rows };
      console.log('GenInfo loaded:', genInfoData.fileName, genInfoData.totalMembers, 'members');
    }
  } catch(e) { console.error('GenInfo DB load:', e.message); }
}
loadGenInfoFromDB();

app.get('/api/geninfo/status', requireAuth, function(req, res) {
  if (!genInfoData) return res.json({ hasData: false });
  res.json({ hasData: true, fileName: genInfoData.fileName, totalMembers: genInfoData.totalMembers, rows: genInfoData.rows });
});

app.post('/api/geninfo/upload', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { rows, fileName, totalMembers } = req.body;
    if (!rows || !rows.length) return res.status(400).json({ error: 'No rows provided' });
    genInfoData = { rows, fileName: fileName || 'team.xlsx', totalMembers: totalMembers || rows.length };
    try {
      await pool.query('DELETE FROM geninfo_data');
      await pool.query('INSERT INTO geninfo_data (file_name, total_members, rows) VALUES ($1,$2,$3)',
        [genInfoData.fileName, genInfoData.totalMembers, JSON.stringify(rows)]);
      console.log('GenInfo saved to DB:', rows.length, 'members');
    } catch(dbErr) { console.error('GenInfo DB save:', dbErr.message); }
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'GenInfo: ' + genInfoData.fileName, '');
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.delete('/api/geninfo/clear', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM geninfo_data');
    genInfoData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});


// ─── TWILIO VOIP ────────────────────────────────────────────────────────────
// STATUS: READY — Set these 4 env vars in Render to activate:
//   TWILIO_ACCOUNT_SID   → from twilio.com console
//   TWILIO_AUTH_TOKEN    → from twilio.com console
//   TWILIO_PHONE_NUMBER  → your Twilio number e.g. +12015551234
//   TWILIO_TWIML_APP_SID → create TwiML App in Twilio console, set Voice URL to:
//                          https://azr-operations.com/api/voip/twiml

var TWILIO_CONFIGURED = !!(
  process.env.TWILIO_ACCOUNT_SID &&
  process.env.TWILIO_AUTH_TOKEN &&
  process.env.TWILIO_PHONE_NUMBER &&
  process.env.TWILIO_TWIML_APP_SID
);

if (TWILIO_CONFIGURED) {
  console.log('✅ Twilio VoIP: CONFIGURED and ready');
} else {
  console.log('⚠ Twilio VoIP: Not configured (set env vars to activate)');
}

// Check VoIP status + issue browser token
app.get('/api/voip/status', requireAuth, async function(req, res) {
  if (!TWILIO_CONFIGURED) return res.json({ configured: false, reason: 'Missing env vars' });
  var twilio;
  try { twilio = require('twilio'); } catch(e) {
    return res.json({ configured: false, reason: 'twilio package missing' });
  }
  try {
    var accountSid  = process.env.TWILIO_ACCOUNT_SID;
    var apiKey      = process.env.TWILIO_API_KEY;
    var apiSecret   = process.env.TWILIO_API_SECRET;
    var twimlAppSid = process.env.TWILIO_TWIML_APP_SID;
    var identity    = (req.user.username || 'azhar_user').replace(/[^a-zA-Z0-9_]/g, '_');

    // AccessToken with API Key — required for Twilio JS SDK v1.x and v2.x
    var AccessToken = twilio.jwt.AccessToken;
    var VoiceGrant  = AccessToken.VoiceGrant;
    var grant = new VoiceGrant({
      outgoingApplicationSid: twimlAppSid,
      incomingAllow: false
    });
    var token = new AccessToken(accountSid, apiKey, apiSecret, {
      identity: identity,
      ttl: 3600
    });
    token.addGrant(grant);
    var jwt = token.toJwt();
    console.log('✅ VoIP token generated for:', identity);
    res.json({ configured: true, token: jwt });
  } catch(e) {
    console.error('❌ VoIP token error:', e.message);
    res.json({ configured: false, error: e.message });
  }
});

// TwiML — tells Twilio what to do when call connects (dial out to real number)
app.post('/api/voip/twiml', function(req, res) {
  var to = req.body.To || req.query.To;
  var callerId = process.env.TWILIO_PHONE_NUMBER || '';
  res.set('Content-Type', 'text/xml');
  if (to && to.startsWith('+')) {
    // Dial with two-way audio bridge
    res.send('<?xml version="1.0" encoding="UTF-8"?>' +
      '<Response>' +
        '<Dial callerId="' + callerId + '" timeout="30" record="do-not-record">' +
          '<Number>' + to + '</Number>' +
        '</Dial>' +
      '</Response>');
  } else if (to && to.startsWith('client:')) {
    // Browser client call
    res.send('<?xml version="1.0" encoding="UTF-8"?>' +
      '<Response>' +
        '<Dial callerId="' + callerId + '">' +
          '<Client>' + to.replace('client:','') + '</Client>' +
        '</Dial>' +
      '</Response>');
  } else {
    res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Say>Call configuration error.</Say></Response>');
  }
});

// Initiate outbound call from server (alternative method)
app.post('/api/voip/call', requireAuth, async function(req, res) {
  if (!TWILIO_CONFIGURED) return res.status(503).json({ error: 'VoIP not configured' });
  try {
    var { to, from_number, from_name } = req.body;
    console.log('VoIP call request - to:', to, 'from:', from_number);
    if (!to) return res.status(400).json({ error: 'No destination number' });

    // ── SECURITY: Only allow calls to registered General Info numbers ──
    if (genInfoData && genInfoData.rows && genInfoData.rows.length) {
      var cleanTo = to.replace(/\s+/g, '').replace(/^00/, '+');
      var allowed = genInfoData.rows.some(function(row) {
        var contact = String(row['CONTACT'] || row['contact'] || '').replace(/\s+/g, '');
        if (!contact) return false;
        if (!contact.startsWith('+')) contact = '+' + contact;
        return contact === cleanTo ||
               contact.replace('+','') === cleanTo.replace('+','') ||
               cleanTo.endsWith(contact.slice(-9));
      });
      if (!allowed) {
        await auditLog(req.user.uid, req.user.username, 'VOIP_BLOCKED', 'Blocked: ' + to, '');
        return res.status(403).json({ error: 'Number not registered in General Info.' });
      }
    }

    var twilio = require('twilio');
    var client = twilio(process.env.TWILIO_ACCOUNT_SID, process.env.TWILIO_AUTH_TOKEN);
    var conferenceName = 'bridge-' + Date.now();
    var baseUrl = 'https://azr-operations.com';

    if (from_number && from_number.startsWith('+')) {
      // Call caller first — when they answer, Twilio dials the destination
      var encodedTo = encodeURIComponent(to);
      var c1 = await client.calls.create({
        to: from_number,
        from: process.env.TWILIO_PHONE_NUMBER,
        url: baseUrl + '/api/voip/dial?to=' + encodedTo
      });
      console.log('VoIP BRIDGE: ' + from_number + ' -> ' + to);
      await auditLog(req.user.uid, req.user.username, 'VOIP_BRIDGE', from_number + ' <-> ' + to, '');
      res.json({ success: true, mode: 'bridge' });
    } else {
      // Outbound only
      var call = await client.calls.create({
        to: to,
        from: process.env.TWILIO_PHONE_NUMBER,
        url: baseUrl + '/api/voip/twiml'
      });
      console.log('VoIP OUTBOUND: ' + req.user.username + ' -> ' + to);
      await auditLog(req.user.uid, req.user.username, 'VOIP_CALL', 'Called: ' + to, '');
      res.json({ success: true, mode: 'outbound' });
    }
  } catch(e) {
    console.error('VoIP call error FULL:', JSON.stringify(e));
    console.error('VoIP call error msg:', e.message);
    console.error('TWILIO_ACCOUNT_SID set:', !!process.env.TWILIO_ACCOUNT_SID);
    console.error('TWILIO_AUTH_TOKEN set:', !!process.env.TWILIO_AUTH_TOKEN);
    console.error('TWILIO_PHONE_NUMBER set:', !!process.env.TWILIO_PHONE_NUMBER);
    res.status(500).json({ error: e.message });
  }
});

// ── CONFERENCE BRIDGE TwiML ──
app.get('/api/voip/dial', function(req, res) {
  var to = req.query.to || '';
  res.set('Content-Type', 'text/xml');
  if (to) {
    res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Dial callerId="' + 
      (process.env.TWILIO_PHONE_NUMBER || '') + '" timeout="30">' + to + '</Dial></Response>');
  } else {
    res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Say>Connection error.</Say></Response>');
  }
});

app.get('/api/voip/conference', function(req, res) {
  var room = req.query.room || 'azhar-default';
  res.set('Content-Type', 'text/xml');
  res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Dial><Conference startConferenceOnEnter="true" endConferenceOnExit="true" maxParticipants="2" record="do-not-record">' + room + '</Conference></Dial></Response>');
});

app.post('/api/voip/conference', function(req, res) {
  var room = req.query.room || req.body.room || 'azhar-default';
  res.set('Content-Type', 'text/xml');
  res.send('<?xml version="1.0" encoding="UTF-8"?><Response><Dial><Conference startConferenceOnEnter="true" endConferenceOnExit="true" maxParticipants="2" record="do-not-record">' + room + '</Conference></Dial></Response>');
});

// ─── VOIP DEBUG (temporary) ─────────────────────────────────────────────────
app.get('/api/voip/debug', function(req, res) {
  res.json({
    configured: TWILIO_CONFIGURED,
    has_sid: !!process.env.TWILIO_ACCOUNT_SID,
    has_token: !!process.env.TWILIO_AUTH_TOKEN,
    has_phone: !!process.env.TWILIO_PHONE_NUMBER,
    has_twiml: !!process.env.TWILIO_TWIML_APP_SID,
    has_api_key: !!process.env.TWILIO_API_KEY,
    has_api_secret: !!process.env.TWILIO_API_SECRET,
    twilio_pkg: (function(){ try{ require('twilio'); return 'OK'; }catch(e){ return e.message; }})()
  });
});

// ─── SERVE TWILIO SDK ────────────────────────────────────────────────────────
app.get('/twilio-sdk.js', function(req, res) {
  try {
    // Try to serve from npm package
    var sdkPath = require.resolve('@twilio/voice-sdk/dist/twilio.js');
    res.setHeader('Content-Type', 'application/javascript');
    res.setHeader('Cache-Control', 'public, max-age=86400');
    require('fs').createReadStream(sdkPath).pipe(res);
  } catch(e) {
    try {
      var sdkPath2 = require.resolve('@twilio/voice-sdk/dist/twilio.min.js');
      res.setHeader('Content-Type', 'application/javascript');
      require('fs').createReadStream(sdkPath2).pipe(res);
    } catch(e2) {
      // Fallback: redirect to CDN (browser can access even if server can't)
      res.redirect('https://sdk.twilio.com/js/client/v1.14/twilio.js');
    }
  }
});

// ══════════════════════════════════════════════════════════
// AUTOMATION TRACKING API ROUTES — PER-MONTH MERGE
// ══════════════════════════════════════════════════════════
var automationData = null;

// New table structure: one row per month
// automation_months: month TEXT PK, auto INT, manual INT, total INT, org_data JSONB, updated_at, updated_by, file_name

async function initAutomationMonthsTable() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS automation_months (
      month TEXT PRIMARY KEY,
      auto_count INT DEFAULT 0,
      manual_count INT DEFAULT 0,
      total_count INT DEFAULT 0,
      org_data JSONB DEFAULT '{}'::jsonb,
      updated_at TIMESTAMPTZ DEFAULT NOW(),
      updated_by TEXT,
      file_name TEXT
    )`);
    console.log('automation_months table ready');
  } catch(e) { console.error('automation_months init:', e.message); }
}
initAutomationMonthsTable();

async function loadAutomationFromDB() {
  try {
    // Load from new per-month table
    var r = await pool.query('SELECT * FROM automation_months ORDER BY month');
    if (r.rows.length) {
      var monthData = {}, orgData = {}, totalAuto = 0, totalManual = 0;
      var latestFile = '', latestBy = '', latestAt = null;
      r.rows.forEach(function(row) {
        monthData[row.month] = { auto: row.auto_count, manual: row.manual_count, total: row.total_count };
        totalAuto   += row.auto_count;
        totalManual += row.manual_count;
        // Merge org data
        var od = row.org_data || {};
        Object.keys(od).forEach(function(org) {
          if (!orgData[org]) orgData[org] = { total:0, auto:0, manual:0, ots:[] };
          orgData[org].total  += od[org].total  || 0;
          orgData[org].auto   += od[org].auto   || 0;
          orgData[org].manual += od[org].manual || 0;
          (od[org].ots||[]).forEach(function(ot){
            if (orgData[org].ots.indexOf(ot) === -1) orgData[org].ots.push(ot);
          });
        });
        if (!latestAt || new Date(row.updated_at) > new Date(latestAt)) {
          latestAt = row.updated_at; latestFile = row.file_name; latestBy = row.updated_by;
        }
      });
      var total = totalAuto + totalManual;
      var rate  = total ? +(totalAuto/total*100).toFixed(2) : 0;
      var months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      var sortedMonths = months.filter(function(m){ return monthData[m]; });
      var latestMonth  = sortedMonths[sortedMonths.length-1] || '';
      var latestRate   = latestMonth ? +(monthData[latestMonth].auto/monthData[latestMonth].total*100).toFixed(2) : 0;
      automationData = {
        uploadedAt: latestAt, uploadedBy: latestBy, fileName: latestFile,
        totalRecords: total, rows: monthData, orgRows: orgData,
        summary: { total, auto: totalAuto, manual: totalManual, rate, orgRows: orgData },
        sortedMonths, latestMonth, latestRate
      };
      console.log('Automation loaded from DB:', total, 'total records across', sortedMonths.length, 'months:', sortedMonths.join(', '));
    }
  } catch(e) { console.error('Automation DB load:', e.message); }
}
loadAutomationFromDB();

app.get('/api/automation/status', requireAuth, function(req, res) {
  if (!automationData) return res.json({ hasData: false });
  res.json({
    hasData: true,
    uploadedAt:   automationData.uploadedAt,
    uploadedBy:   automationData.uploadedBy,
    fileName:     automationData.fileName,
    totalRecords: automationData.totalRecords,
    summary:      automationData.summary,
    rows:         automationData.rows,
    orgRows:      automationData.orgRows || {}
  });
});

app.post('/api/automation/upload', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { rows, orgRows, fileName, totalRecords, summary } = req.body;
    if (!rows) return res.status(400).json({ error: 'No data provided' });

    // rows = { Jan: {auto, manual, total}, Feb: {...}, ... }
    // orgRows = { 'Victory-Food': {total, auto, manual, ots:[...]}, ... }
    var monthsUpdated = [];

    for (var month in rows) {
      var md = rows[month];
      var od = {};
      // Build org_data for this month from orgRows — approximate split
      if (orgRows) {
        Object.keys(orgRows).forEach(function(org) {
          od[org] = orgRows[org]; // store full org totals per upload
        });
      }
      await pool.query(`
        INSERT INTO automation_months (month, auto_count, manual_count, total_count, org_data, updated_at, updated_by, file_name)
        VALUES ($1, $2, $3, $4, $5, NOW(), $6, $7)
        ON CONFLICT (month) DO UPDATE SET
          auto_count   = EXCLUDED.auto_count,
          manual_count = EXCLUDED.manual_count,
          total_count  = EXCLUDED.total_count,
          org_data     = EXCLUDED.org_data,
          updated_at   = NOW(),
          updated_by   = EXCLUDED.updated_by,
          file_name    = EXCLUDED.file_name
      `, [month, md.auto||0, md.manual||0, md.total||0, JSON.stringify(od), req.user.username, fileName||'automation.xlsx']);
      monthsUpdated.push(month);
    }

    console.log('Automation months upserted:', monthsUpdated.join(', '));
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Automation: ' + fileName + ' months: ' + monthsUpdated.join(','), '');

    // Reload full aggregated data from DB
    await loadAutomationFromDB();
    res.json({ success: true, monthsUpdated: monthsUpdated, totalRecords: automationData ? automationData.totalRecords : totalRecords });
  } catch(e) {
    console.error('Automation upload error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/automation/clear', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM automation_months');
    automationData = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Per-month delete endpoint (optional future use)
app.delete('/api/automation/month/:month', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM automation_months WHERE month=$1', [req.params.month]);
    await loadAutomationFromDB();
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

var PORT=process.env.PORT||3000;
// ══════════════════════════════════════════════════════════
// DELIVERY SCHEDULE COMPLIANCE API ROUTES
// ══════════════════════════════════════════════════════════
var deliveryScheduleLookup = null;
var deliveryData = null;

async function initDeliveryTables() {
  try {
    await pool.query(`CREATE TABLE IF NOT EXISTS delivery_schedule (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      customer_count INT,
      lookup JSONB
    )`);
    await pool.query(`CREATE TABLE IF NOT EXISTS delivery_data (
      id SERIAL PRIMARY KEY,
      uploaded_at TIMESTAMPTZ DEFAULT NOW(),
      uploaded_by TEXT,
      file_name TEXT,
      total_orders INT,
      summary JSONB
    )`);
    // Load existing data
    var sr = await pool.query('SELECT * FROM delivery_schedule ORDER BY uploaded_at DESC LIMIT 1');
    if (sr.rows.length) {
      // Convert string keys to integers after JSONB parse
      var rawLookup = sr.rows[0].lookup || {};
      deliveryScheduleLookup = {};
      Object.keys(rawLookup).forEach(function(k) {
        deliveryScheduleLookup[parseInt(k)] = rawLookup[k];
      });
      console.log('Delivery schedule loaded:', Object.keys(deliveryScheduleLookup).length, 'customers');
    }
    var dr = await pool.query('SELECT * FROM delivery_data ORDER BY uploaded_at DESC LIMIT 1');
    if (dr.rows.length) {
      deliveryData = { summary: dr.rows[0].summary, uploadedBy: dr.rows[0].uploaded_by, fileName: dr.rows[0].file_name, totalOrders: dr.rows[0].total_orders };
      console.log('Delivery data loaded:', dr.rows[0].total_orders, 'orders');
    }
  } catch(e) { console.error('Delivery init:', e.message); }
}
initDeliveryTables();

app.get('/api/delivery/status', requireAuth, function(req, res) {
  res.json({
    hasSchedule: !!deliveryScheduleLookup,
    scheduleCustomers: deliveryScheduleLookup ? Object.keys(deliveryScheduleLookup).length : 0,
    scheduleLookup: deliveryScheduleLookup || {},
    hasData: !!deliveryData,
    summary: deliveryData ? deliveryData.summary : null,
    fileName: deliveryData ? deliveryData.fileName : null,
    uploadedBy: deliveryData ? deliveryData.uploadedBy : null,
    totalOrders: deliveryData ? deliveryData.totalOrders : 0
  });
});

app.post('/api/delivery/schedule', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { lookup, fileName, customerCount } = req.body;
    if (!lookup) return res.status(400).json({ error: 'No schedule data' });
    // Store with integer keys in memory
    deliveryScheduleLookup = {};
    Object.keys(lookup).forEach(function(k) {
      deliveryScheduleLookup[parseInt(k)] = lookup[k];
    });
    await pool.query('DELETE FROM delivery_schedule');
    await pool.query('INSERT INTO delivery_schedule (uploaded_by, file_name, customer_count, lookup) VALUES ($1,$2,$3,$4)',
      [req.user.username, fileName||'schedule.xlsx', customerCount||0, JSON.stringify(lookup)]);
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Delivery Schedule: ' + fileName, '');
    res.json({ success: true, customerCount: customerCount });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

app.post('/api/delivery/data', requireAuth, requireRole('superadmin','subadmin'), async function(req, res) {
  try {
    var { summary, fileName, totalOrders } = req.body;
    if (!summary) return res.status(400).json({ error: 'No data' });
    deliveryData = { summary, fileName: fileName||'oracle.xlsx', uploadedBy: req.user.username, totalOrders: totalOrders||0 };
    await pool.query('DELETE FROM delivery_data');
    await pool.query('INSERT INTO delivery_data (uploaded_by, file_name, total_orders, summary) VALUES ($1,$2,$3,$4)',
      [req.user.username, deliveryData.fileName, deliveryData.totalOrders, JSON.stringify(summary)]);
    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Delivery Data: ' + fileName, '');
    res.json({ success: true, totalOrders: deliveryData.totalOrders });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ── Server-side Oracle classification (avoids browser freeze) ────────
// Parses a "Request Date" style value from an Oracle/Excel export into a UTC calendar date.
// Handles: genuine JS Date objects (from real Excel date-serial cells), ISO text strings with
// a timezone offset (e.g. Oracle's "2026-01-19T00:00:00.000+04:00" — common when Request Date
// is exported as TEXT, not a real date cell), plain "YYYY-MM-DD", and true numeric Excel serials.
// Always returns the literal calendar date as written — never shifts it via timezone conversion,
// since a delivery "day" is a business date, not a precise instant.
function parseOrderDate(dateRaw) {
  if (dateRaw instanceof Date) {
    return new Date(Date.UTC(dateRaw.getFullYear(), dateRaw.getMonth(), dateRaw.getDate()));
  }
  var s = String(dateRaw || '').trim();
  if (!s) return null;

  // ISO-style date text, with or without time/offset: take the Y-M-D digits literally.
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) {
    return new Date(Date.UTC(parseInt(iso[1]), parseInt(iso[2]) - 1, parseInt(iso[3])));
  }
  // DD/MM/YYYY or MM/DD/YYYY style — Oracle exports for this region use DD/MM/YYYY.
  var slash = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (slash) {
    var yy = parseInt(slash[3]); if (yy < 100) yy += 2000;
    return new Date(Date.UTC(yy, parseInt(slash[1]) - 1, parseInt(slash[2])));
  }
  // Only treat as a real Excel serial number if the ENTIRE string is numeric (no letters/dashes) —
  // this is what previously misfired on ISO text like "2026-01-19T..." (parseFloat saw "2026" and
  // treated it as a serial, collapsing every 2026 date onto the same fake day).
  if (/^\d+(\.\d+)?$/.test(s)) {
    var f = parseFloat(s);
    if (!isNaN(f) && f > 1000) return new Date(Math.round((f - 25569) * 86400 * 1000));
  }
  // Last resort: let JS try, but normalize to a UTC calendar date (not a shifted instant).
  var fallback = new Date(s);
  if (!isNaN(fallback.getTime())) {
    return new Date(Date.UTC(fallback.getUTCFullYear(), fallback.getUTCMonth(), fallback.getUTCDate()));
  }
  return null;
}

// Collapses case/whitespace-only duplicates (e.g. "e-commerce", "E-commerce", "E-Commerce" from
// different Oracle exports) into one consistent, nicely-cased channel label used as the grouping key.
function normalizeChannelName(raw) {
  var s = String(raw || '').trim().replace(/\s+/g, ' ');
  if (!s) return '';
  return s.split(' ').map(function(word) {
    return word.split('-').map(function(part) {
      return part ? part.charAt(0).toUpperCase() + part.slice(1).toLowerCase() : part;
    }).join('-');
  }).join(' ');
}

app.post('/api/delivery/classify', requireAuth, requireRole('superadmin','subadmin'), upload.single('file'), async function(req, res) {
  try {
    if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
    if (!deliveryScheduleLookup || Object.keys(deliveryScheduleLookup).length === 0) {
      return res.status(400).json({ error: 'No schedule loaded on server. Please upload schedule first.' });
    }

    console.log('DS Classify: Reading', req.file.originalname, req.file.size, 'bytes');
    var wb = XLSX.read(req.file.buffer, { type: 'buffer', cellDates: true, cellNF: false, cellHTML: false, cellFormula: false });

    // Find correct sheet
    var sheetName = wb.SheetNames.find(function(s){ return s.trim() === 'Data'; })
      || wb.SheetNames.find(function(s){ return s.toUpperCase().includes('MASTER'); })
      || wb.SheetNames[0];

    console.log('DS Classify: Using sheet', sheetName, 'of', wb.SheetNames);
    var data = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
    console.log('DS Classify: Rows', data.length);

    if (!data.length) return res.status(400).json({ error: 'No data rows found in file. Sheet: ' + sheetName });

    var keys = Object.keys(data[0]);
    var colSiteId  = keys.find(function(k){ return /site.?id/i.test(k); }) || '';
    var colDate    = keys.find(function(k){ return /^rsd$/i.test(k.trim()) || /request.?date/i.test(k) || /invoice.?date/i.test(k); }) || '';
    var colTemp    = keys.find(function(k){ return /ambient.*frozen|frozen.*ambient/i.test(k) || k.trim() === 'Ambient / Frozen'; }) || '';
    var colChannel = keys.find(function(k){ return /channel/i.test(k); }) || '';
    var colMonth   = keys.find(function(k){ return k.trim().toUpperCase() === 'MONTH'; }) || '';
    var colCust    = keys.find(function(k){ return /customer.?name|customer_name/i.test(k); }) || '';
    var colOrg     = keys.find(function(k){ return k.trim().toUpperCase() === 'ORG'; })
      || keys.find(function(k){ return k.trim().toUpperCase() === 'WAREHOUSE'; }) || '';
    var colOrderType = keys.find(function(k){ return /order.?type/i.test(k); }) || '';
    var colStatus  = keys.find(function(k){ return k.trim().toUpperCase() === 'STATUS'; }) || '';

    console.log('DS Classify columns: siteId=' + colSiteId + ' date=' + colDate + ' temp=' + colTemp + ' channel=' + colChannel + ' month=' + colMonth + ' org=' + colOrg + ' status=' + colStatus);

    // Only these orgs are considered part of this delivery-schedule analysis.
    var ALLOWED_ORGS = ['DGC', 'DCV', 'DGS'];
    // Only these order statuses count as real deliveries (cancelled/rejected never happened).
    var ALLOWED_STATUSES = ['BOOKED', 'CLOSED'];

    var DAYS = ['SUN','MON','TUE','WED','THU','FRI','SAT'];
    var MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    function getMonthKey(row) {
      if (colMonth && row[colMonth]) return String(row[colMonth]).trim();
      var dateRaw = row[colDate];
      if (!dateRaw) return '';
      var parsedDt = parseOrderDate(dateRaw);
      if (!parsedDt) return '';
      return MONTH_NAMES[parsedDt.getUTCMonth()] + '-' + String(parsedDt.getUTCFullYear()).slice(2);
    }

    // Pre-pass: any month with under 10 orders is treated as a data-entry date typo
    // (e.g. "Jan-23" or "Dec-26" with a single stray order) rather than real volume,
    // and excluded entirely so it doesn't clutter the pivot with near-empty columns.
    var JUNK_MONTH_THRESHOLD = 10;
    var monthCounts = {};
    data.forEach(function(row) {
      if (colOrg) { var ov = String(row[colOrg]||'').trim().toUpperCase(); if (ov && ALLOWED_ORGS.indexOf(ov) === -1) return; }
      if (colStatus) { var sv = String(row[colStatus]||'').trim().toUpperCase(); if (sv && ALLOWED_STATUSES.indexOf(sv) === -1) return; }
      var m = getMonthKey(row);
      if (m) monthCounts[m] = (monthCounts[m] || 0) + 1;
    });
    var junkMonths = {};
    Object.keys(monthCounts).forEach(function(m) { if (monthCounts[m] < JUNK_MONTH_THRESHOLD) junkMonths[m] = true; });

    var monthData = {}, channelData = {}, dayData = {}, noSchedCusts = {}, oosCusts = {};
    var tempData = { Ambient:{scheduled:0,oos:0,noSched:0,total:0}, Frozen:{scheduled:0,oos:0,noSched:0,total:0} };
    var scheduled = 0, oos = 0, noSched = 0;
    var skippedOrg = 0, skippedStatus = 0, skippedJunkMonth = 0;

    data.forEach(function(row) {
      // Filter: org must be one of the allowed orgs (if an org/warehouse column exists)
      if (colOrg) {
        var orgVal = String(row[colOrg]||'').trim().toUpperCase();
        if (orgVal && ALLOWED_ORGS.indexOf(orgVal) === -1) { skippedOrg++; return; }
      }
      // Filter: exclude cancelled/rejected orders (if a status column exists)
      if (colStatus) {
        var statusVal = String(row[colStatus]||'').trim().toUpperCase();
        if (statusVal && ALLOWED_STATUSES.indexOf(statusVal) === -1) { skippedStatus++; return; }
      }
      // Filter: exclude stray/junk months (likely date typos with under 10 total orders)
      var monthCheck = getMonthKey(row);
      if (monthCheck && junkMonths[monthCheck]) { skippedJunkMonth++; return; }

      var siteRaw = row[colSiteId];
      var site = null;
      try { site = siteRaw ? parseInt(parseFloat(String(siteRaw))) : null; } catch(e){}

      var day = null, month = '';
      var dateRaw = row[colDate];
      if (dateRaw) {
        try {
          var parsedDt = parseOrderDate(dateRaw);
          if (parsedDt) {
            day = DAYS[parsedDt.getUTCDay()];
            if (colMonth && row[colMonth]) {
              month = String(row[colMonth]).trim();
            } else {
              var mNames = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
              month = mNames[parsedDt.getUTCMonth()] + '-' + String(parsedDt.getUTCFullYear()).slice(2);
            }
          }
        } catch(e){}
      }
      if (!month && colMonth && row[colMonth]) month = String(row[colMonth]).trim();

      var temp    = String(row[colTemp]||'').trim().toUpperCase();
      var channel = normalizeChannelName(String(row[colChannel]||'').trim());
      var org     = String(row[colOrg]||'').trim();
      var cust    = String(row[colCust]||'').trim();
      // Get channel and org from schedule if not in oracle
      if ((!channel || !org) && site && deliveryScheduleLookup[site]) {
        var sl2 = deliveryScheduleLookup[site];
        if (!channel && sl2.accountType) channel = normalizeChannelName(sl2.accountType);
        if (!org && sl2.org) org = sl2.org;
      }

      var status;
      if (!site || !deliveryScheduleLookup[site]) {
        status = 'No Schedule';
        if (site) {
          if (!noSchedCusts[site]) noSchedCusts[site] = { name: cust, orders: 0, months: {}, monthCounts: {}, channel: channel, org: org };
          noSchedCusts[site].orders++;
          if (month) {
            noSchedCusts[site].months[month] = 1;
            noSchedCusts[site].monthCounts[month] = (noSchedCusts[site].monthCounts[month] || 0) + 1;
          }
        }
      } else {
        var sl = deliveryScheduleLookup[site];
        var day3 = day ? day.substring(0,3) : null;
        if (day3 && sl.days && (day3 in sl.days)) {
          var schedTemp = sl.days[day3].replace(/\s/g,'').toUpperCase();
          if (temp === 'FROZEN') {
            status = schedTemp.indexOf('FROZEN') !== -1 ? 'Scheduled' : 'Out of Schedule';
          } else {
            status = schedTemp ? 'Scheduled' : 'Out of Schedule';
          }
        } else {
          status = 'Out of Schedule';
        }
        if (status === 'Out of Schedule') {
          if (!oosCusts[site]) oosCusts[site] = { name: cust, orders: 0, schedDays: sl.days ? Object.keys(sl.days).join(', ') : '', orderedDay: day3||'?', channel: channel, org: org, monthCounts: {} };
          oosCusts[site].orders++;
          if (month) oosCusts[site].monthCounts[month] = (oosCusts[site].monthCounts[month] || 0) + 1;
        }
      }

      if (status === 'Scheduled') scheduled++;
      else if (status === 'Out of Schedule') oos++;
      else noSched++;
      var statusKey = status==='Scheduled'?'scheduled':status==='Out of Schedule'?'oos':'noSched';
      // Track by temperature
      var tempKey = (temp === 'FROZEN') ? 'Frozen' : 'Ambient';
      if (!tempData[tempKey]) tempData[tempKey] = {scheduled:0,oos:0,noSched:0,total:0};
      tempData[tempKey][statusKey]++;
      tempData[tempKey].total++;

      if (month) {
        if (!monthData[month]) monthData[month] = { scheduled:0, oos:0, noSched:0, total:0, cat:{ Ambient:{scheduled:0,oos:0,noSched:0,total:0}, Frozen:{scheduled:0,oos:0,noSched:0,total:0} }, channels:{} };
        monthData[month][statusKey]++;
        monthData[month].total++;
        monthData[month].cat[tempKey][statusKey]++;
        monthData[month].cat[tempKey].total++;
        // Per-month, per-channel breakdown — needed so the Month filter can correctly
        // recompute the channel rows in the pivot table (they used to stay stuck showing
        // all-time totals regardless of which month was selected).
        if (channel) {
          if (!monthData[month].channels[channel]) monthData[month].channels[channel] = { scheduled:0, oos:0, noSched:0, total:0, cat:{ Ambient:{scheduled:0,oos:0,noSched:0,total:0}, Frozen:{scheduled:0,oos:0,noSched:0,total:0} } };
          monthData[month].channels[channel][statusKey]++;
          monthData[month].channels[channel].total++;
          monthData[month].channels[channel].cat[tempKey][statusKey]++;
          monthData[month].channels[channel].cat[tempKey].total++;
        }
      }
      if (channel) {
        if (!channelData[channel]) channelData[channel] = { scheduled:0, oos:0, noSched:0, total:0, cat:{ Ambient:{scheduled:0,oos:0,noSched:0,total:0}, Frozen:{scheduled:0,oos:0,noSched:0,total:0} } };
        channelData[channel][statusKey]++;
        channelData[channel].total++;
        channelData[channel].cat[tempKey][statusKey]++;
        channelData[channel].cat[tempKey].total++;
      }
      if (day) {
        var d3 = day.substring(0,3);
        if (!dayData[d3]) dayData[d3] = { scheduled:0, oos:0, noSched:0, total:0 };
        dayData[d3][statusKey]++;
        dayData[d3].total++;
      }
    });

    var total = scheduled + oos + noSched;
    // Match Rejection YTD's rate formula: (count/total*100) to 2 decimal places, not rounded to a whole number.
    var sp = total ? parseFloat((scheduled/total*100).toFixed(2)) : 0;
    var op = total ? parseFloat((oos/total*100).toFixed(2)) : 0;
    var np = total ? parseFloat((noSched/total*100).toFixed(2)) : 0;

    var noSchedArr = Object.keys(noSchedCusts).map(function(s){
      var d = noSchedCusts[s];
      return { site:s, name:d.name, orders:d.orders, months:Object.keys(d.months).join(', '), monthCounts:d.monthCounts||{}, channel:d.channel, org:d.org };
    }).sort(function(a,b){ return b.orders-a.orders; }).slice(0,50);

    var oosArr = Object.keys(oosCusts).map(function(s){
      var d = oosCusts[s];
      return { site:s, name:d.name, orders:d.orders, schedDays:d.schedDays, orderedDay:d.orderedDay, monthCounts:d.monthCounts||{}, channel:d.channel, org:d.org };
    }).sort(function(a,b){ return b.orders-a.orders; }).slice(0,50);

    // Rebuild monthData with keys inserted in true chronological order (e.g. Jan-26, Feb-26, Mar-26...)
    // — object key order in JS preserves insertion order for non-numeric keys, so this fixes the
    // pivot header/columns and the Month filter dropdown, which were previously showing months in
    // whatever order they first appeared in the file (e.g. Apr, Feb, Jan, Mar, May).
    var MONTH_ORDER = { Jan:0,Feb:1,Mar:2,Apr:3,May:4,Jun:5,Jul:6,Aug:7,Sep:8,Oct:9,Nov:10,Dec:11 };
    var sortedMonthKeys = Object.keys(monthData).sort(function(a, b) {
      var pa = a.split('-'), pb = b.split('-');
      var ya = parseInt(pa[1]), yb = parseInt(pb[1]);
      if (ya !== yb) return ya - yb;
      return (MONTH_ORDER[pa[0]] || 0) - (MONTH_ORDER[pb[0]] || 0);
    });
    var sortedMonthData = {};
    sortedMonthKeys.forEach(function(m) { sortedMonthData[m] = monthData[m]; });
    monthData = sortedMonthData;

    var summary = {
      total:total, scheduled:scheduled, oos:oos, noSched:noSched,
      schedPct:sp, oosPct:op, noSchedPct:np,
      monthData:monthData, channelData:channelData, dayData:dayData, tempData:tempData,
      noSchedCustomers:noSchedArr, oosCustomers:oosArr,
      skippedOrg:skippedOrg, skippedStatus:skippedStatus, skippedJunkMonth:skippedJunkMonth,
      rowsInFile: data.length
    };

    // Save to DB
    await pool.query('DELETE FROM delivery_data');
    await pool.query('INSERT INTO delivery_data (uploaded_by, file_name, total_orders, summary) VALUES ($1,$2,$3,$4)',
      [req.user.username, req.file.originalname, total, JSON.stringify(summary)]);
    deliveryData = { summary, fileName: req.file.originalname, uploadedBy: req.user.username, totalOrders: total };

    await auditLog(req.user.uid, req.user.username, 'UPLOAD', 'Delivery Oracle: ' + req.file.originalname + ' ' + total + ' orders', '');
    console.log('DS Classify complete:', total, 'orders — Scheduled:', sp + '%', 'OOS:', op + '%', 'NoSched:', np + '%',
      '| Excluded — org:', skippedOrg, 'status:', skippedStatus, 'junk-month:', skippedJunkMonth, 'of', data.length, 'rows in file',
      Object.keys(junkMonths).length ? '(junk months: ' + Object.keys(junkMonths).join(', ') + ')' : '');

    res.json({ success: true, summary: summary, totalOrders: total });
  } catch(e) {
    console.error('DS Classify error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.delete('/api/delivery/schedule/clear', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM delivery_schedule');
    deliveryScheduleLookup = null;
    res.json({ success: true });
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// Kill switch — wipes BOTH the schedule master and the classified Oracle data,
// so you can start completely fresh (e.g. after debugging a bad upload).
app.delete('/api/delivery/reset-all', requireAuth, requireRole('superadmin'), async function(req, res) {
  try {
    await pool.query('DELETE FROM delivery_schedule');
    await pool.query('DELETE FROM delivery_data');
    deliveryScheduleLookup = null;
    deliveryData = null;
    await auditLog(req.user.uid, req.user.username, 'DELETE', 'Delivery Schedule: full reset (schedule + oracle data wiped)', '');
    console.log('DS Reset-all: schedule + oracle data cleared by', req.user.username);
    res.json({ success: true });
  } catch(e) {
    console.error('DS reset-all error:', e.message);
    res.status(500).json({ error: e.message });
  }
});

app.listen(PORT,function(){console.log('AZHAR-AI server running on port '+PORT+(process.env.DATABASE_URL?' with PostgreSQL':' file-only mode'));});
