// ============================================================
// TRANSPORT COST CONTROL MODULE
// Mount with: require('./transport_cost_module')(app, pool, requireAuth, requireRole);
//
// This module reads data that Daily Dispatch and Rejection YTD already capture on
// upload (order_tracking + rejection_rows) — it has no upload of its own. Six things,
// matching what was asked for:
//   1) Daily / Monthly / YTD period toggle, Food/Non-Food + BU-wise + Combined, everywhere.
//   2) Value-tier split: orders & drops at/above AED 1,000 vs below, BU-wise + combined.
//   3) Low-value order tracker: how many low-value orders go out vs their share of total
//      value — the same drop cost applies whether the order is worth AED 50 or AED 5,000.
//   4) Re-delivery tracking: how many re-deliveries per day, attempt distribution
//      (2nd/3rd/4th+), estimated extra cost from the transport rate card.
//   5) Full drop counts, day-wise and BU-wise.
//   6) Rejection -> re-delivery: rejected orders that DID get re-delivered, by BU and
//      by attempt number (1st/2nd/3rd+), so a rejection's downstream cost is visible.
// Every section: BU-wise rows + one Combined (All BU) row, Excel download, color coding.
// ============================================================

module.exports = function (app, pool, requireAuth, requireRole) {

  // Mirrors server.js's TRUCK_RATE_CARD "Multi" rates (a re-delivery almost always goes
  // out on a shared multi-drop route, not a dedicated bulk truck, so Multi is the right
  // default) — kept as a small local constant rather than importing server.js's larger
  // rate table, since only the temperature split matters here. This is an ESTIMATE.
  var REDELIVERY_RATE = { FROZEN: 120, AMBIENT: 104 };
  function estimateRedeliveryRate(truckTypeRaw, temperatureRaw) {
    var t = String(truckTypeRaw || '').toUpperCase();
    if (t.indexOf('FROZEN') !== -1) return REDELIVERY_RATE.FROZEN;
    if (t.indexOf('AMBIENT') !== -1 || t.indexOf('CHILL') !== -1) return REDELIVERY_RATE.AMBIENT;
    var temp = String(temperatureRaw || '').toUpperCase();
    if (temp.indexOf('FROZEN') !== -1 || temp.indexOf('FREEZ') !== -1) return REDELIVERY_RATE.FROZEN;
    return REDELIVERY_RATE.AMBIENT; // most common bucket when temperature isn't on file
  }
  var LOW_VALUE_THRESHOLD = 1000;

  // ── Period resolution: daily (a single date), monthly (a calendar month), or
  // YTD (Jan 1 of the given year through today, or through Dec 31 for a past year). ──
  function resolvePeriod(q) {
    var period = q.period || 'daily';
    var today = new Date().toISOString().split('T')[0];
    if (period === 'monthly') {
      var month = q.month || today.slice(0, 7); // 'YYYY-MM'
      var y = +month.slice(0, 4), m = +month.slice(5, 7);
      var lastDay = new Date(y, m, 0).getDate();
      return { type: 'monthly', date_from: month + '-01', date_to: month + '-' + String(lastDay).padStart(2, '0'), label: month };
    }
    if (period === 'ytd') {
      var year = q.year || today.slice(0, 4);
      var to = (year === today.slice(0, 4)) ? today : (year + '-12-31');
      return { type: 'ytd', date_from: year + '-01-01', date_to: to, label: year + ' YTD' };
    }
    var d = q.date || today;
    return { type: 'daily', date_from: d, date_to: d, label: d };
  }

  // ── Raw data fetch ──
  // Rows within the period (for value-tier, drops, and to know WHICH orders appear that
  // day) plus a lighter order_code+date history up to date_to (unbounded start) so a
  // re-delivery's attempt number is correct even when attempt #1 was before the period.
  async function fetchPeriodRows(dateFrom, dateTo) {
    var r = await pool.query(
      `SELECT order_code, date_key::text AS date_key, customer, value, route, org, drop_type,
         temperature, source_label, city, location_id, truck_type
       FROM order_tracking WHERE date_key >= $1 AND date_key <= $2 ORDER BY date_key ASC`,
      [dateFrom, dateTo]);
    return r.rows;
  }
  async function fetchHistoryUpTo(dateTo) {
    var r = await pool.query(
      `SELECT order_code, date_key::text AS date_key FROM order_tracking WHERE date_key <= $1 ORDER BY date_key ASC`,
      [dateTo]);
    return r.rows;
  }
  // Salon (DIP warehouse) rows run on an unconfirmed transport billing model and never
  // appear in transport's own file — same exclusion used everywhere else on this
  // dashboard (Transport Cost Reconciliation, Duplicate Drops), applied here too so this
  // page's numbers stay consistent with those.
  function isExcludedRow(row) {
    return String(row.source_label || '').toUpperCase() === 'SALON';
  }
  function bucketOf(dropType) {
    var t = String(dropType || '').toLowerCase();
    if (t === 'food') return 'Food';
    if (t === 'nonfood') return 'Non-Food';
    return 'Other';
  }

  // ── 1+2) Value-tier split: orders & drops at/above AED 1,000 vs below, by BU, by
  // Food/Non-Food, plus one Combined (All BU) row. A "drop" here is a distinct
  // (date, location) pair — the same definition used across the rest of the dashboard —
  // credited to whichever tier(s) had an order there that day (a location with both a
  // high- and a low-value order that day counts as a drop in both tiers, since the
  // physical stop happened either way; that's intentional, not double-counting cost). ──
  function buildValueTier(rows) {
    var byOrgType = {}; // "ORG||TYPE" -> { orders_hi, orders_lo, value_hi, value_lo, drops_hi:Set, drops_lo:Set }
    var combined = { orders_hi: 0, orders_lo: 0, value_hi: 0, value_lo: 0, drops_hi: new Set(), drops_lo: new Set() };
    rows.forEach(function (r) {
      if (isExcludedRow(r)) return;
      var org = r.org || '(Unknown)';
      var type = bucketOf(r.drop_type);
      var key = org + '||' + type;
      if (!byOrgType[key]) byOrgType[key] = { org: org, type: type, orders_hi: 0, orders_lo: 0, value_hi: 0, value_lo: 0, drops_hi: new Set(), drops_lo: new Set() };
      var bucket = byOrgType[key];
      var val = +r.value || 0;
      var dropKey = r.date_key + '::' + (r.location_id || r.customer || r.order_code);
      var isHigh = val >= LOW_VALUE_THRESHOLD;
      if (isHigh) { bucket.orders_hi++; bucket.value_hi += val; bucket.drops_hi.add(dropKey); combined.orders_hi++; combined.value_hi += val; combined.drops_hi.add(dropKey); }
      else { bucket.orders_lo++; bucket.value_lo += val; bucket.drops_lo.add(dropKey); combined.orders_lo++; combined.value_lo += val; combined.drops_lo.add(dropKey); }
    });
    var out = Object.keys(byOrgType).sort().map(function (k) {
      var b = byOrgType[k];
      return { org: b.org, type: b.type, orders_ge1k: b.orders_hi, orders_lt1k: b.orders_lo, drops_ge1k: b.drops_hi.size, drops_lt1k: b.drops_lo.size, value_ge1k: Math.round(b.value_hi), value_lt1k: Math.round(b.value_lo) };
    });
    return {
      rows: out,
      combined: { org: 'ALL (Combined)', type: 'All Types', orders_ge1k: combined.orders_hi, orders_lt1k: combined.orders_lo, drops_ge1k: combined.drops_hi.size, drops_lt1k: combined.drops_lo.size, value_ge1k: Math.round(combined.value_hi), value_lt1k: Math.round(combined.value_lo) }
    };
  }

  // ── 3) Low-value order tracker — same numbers as buildValueTier, reshaped to answer
  // "how many low-value orders go out compared with [total] value", BU-wise + combined. ──
  function buildLowValue(valueTier) {
    function shape(row) {
      var totalOrders = row.orders_ge1k + row.orders_lt1k;
      var totalValue = row.value_ge1k + row.value_lt1k;
      return {
        org: row.org, type: row.type,
        low_value_orders: row.orders_lt1k, total_orders: totalOrders,
        low_value_pct: totalOrders ? +((row.orders_lt1k / totalOrders) * 100).toFixed(1) : 0,
        low_value_total_aed: row.value_lt1k, total_value_aed: totalValue,
        low_value_share_of_value_pct: totalValue ? +((row.value_lt1k / totalValue) * 100).toFixed(1) : 0
      };
    }
    return { rows: valueTier.rows.map(shape), combined: shape(valueTier.combined) };
  }

  // ── 4) Re-delivery tracking — attempt number = position of this occurrence in that
  // order's full date history up to date_to. attempt_number > 1 means this occurrence IS
  // a re-delivery (the earlier attempt(s) failed / weren't collected as delivered). ──
  function buildRedelivery(rows, historyRows, dateFrom, dateTo) {
    var historyByCode = {};
    historyRows.forEach(function (h) {
      if (!historyByCode[h.order_code]) historyByCode[h.order_code] = [];
      var dates = historyByCode[h.order_code];
      if (dates[dates.length - 1] !== h.date_key) dates.push(h.date_key); // dedupe same-day dup rows
    });
    var daily = {}; // date -> { orders:Set, cost }
    var byBu = {};  // org -> { orders:Set, cost }
    var byAttempt = { second: 0, third: 0, fourth_plus: 0 };
    var seenOrderAttempt = new Set(); // avoid double counting same order twice on the same day (multi-row same order/day)
    rows.forEach(function (r) {
      if (isExcludedRow(r)) return;
      var dates = historyByCode[r.order_code] || [];
      var attemptNumber = dates.indexOf(r.date_key) + 1;
      if (attemptNumber <= 1) return; // first-ever attempt — not a re-delivery
      var dedupeKey = r.order_code + '::' + r.date_key;
      if (seenOrderAttempt.has(dedupeKey)) return;
      seenOrderAttempt.add(dedupeKey);
      var cost = estimateRedeliveryRate(r.truck_type, r.temperature);
      if (!daily[r.date_key]) daily[r.date_key] = { orders: new Set(), cost: 0 };
      daily[r.date_key].orders.add(r.order_code); daily[r.date_key].cost += cost;
      var org = r.org || '(Unknown)';
      if (!byBu[org]) byBu[org] = { orders: new Set(), cost: 0 };
      byBu[org].orders.add(r.order_code); byBu[org].cost += cost;
      if (attemptNumber === 2) byAttempt.second++;
      else if (attemptNumber === 3) byAttempt.third++;
      else byAttempt.fourth_plus++;
    });
    var dailyRows = Object.keys(daily).sort().map(function (d) {
      return { date: d, orders_redelivered: daily[d].orders.size, estimated_cost: Math.round(daily[d].cost) };
    });
    var buRows = Object.keys(byBu).sort().map(function (org) {
      return { org: org, orders_redelivered: byBu[org].orders.size, estimated_cost: Math.round(byBu[org].cost) };
    });
    var totalOrders = new Set(); var totalCost = 0;
    dailyRows.forEach(function (d) { totalCost += d.estimated_cost; });
    Object.keys(daily).forEach(function (d) { daily[d].orders.forEach(function (o) { totalOrders.add(o); }); });
    return { daily: dailyRows, by_bu: buRows, by_attempt: byAttempt, total_orders_redelivered: totalOrders.size, total_estimated_cost: Math.round(totalCost), note: 'Cost is an ESTIMATE (temperature-based Multi-route rate) — no per-order truck-type was on file for every row.' };
  }

  // ── 5) Full drops — day-wise and BU-wise, distinct (date, location) pairs. ──
  function buildDrops(rows) {
    var dailyBu = {}; // "date||org" -> Set(locations)
    var totalsByBu = {}; // org -> Set(locations) — NOTE: this is a period total, so it is
    // NOT simply the sum of the daily figures (the same location on different days is
    // correctly two drops there, but the same location on the SAME day across multiple
    // rows is correctly one).
    rows.forEach(function (r) {
      if (isExcludedRow(r)) return;
      var org = r.org || '(Unknown)';
      var loc = r.location_id || r.customer || r.order_code;
      var dbKey = r.date_key + '||' + org;
      if (!dailyBu[dbKey]) dailyBu[dbKey] = new Set();
      dailyBu[dbKey].add(loc);
      var totKey = r.date_key + '::' + loc;
      if (!totalsByBu[org]) totalsByBu[org] = new Set();
      totalsByBu[org].add(totKey);
    });
    var dailyRows = Object.keys(dailyBu).sort().map(function (k) {
      var parts = k.split('||'); return { date: parts[0], org: parts[1], drops: dailyBu[k].size };
    });
    var buRows = Object.keys(totalsByBu).sort().map(function (org) { return { org: org, drops: totalsByBu[org].size }; });
    var combinedTotal = buRows.reduce(function (s, r) { return s + r.drops; }, 0);
    return { daily_by_bu: dailyRows, totals_by_bu: buRows, combined_total_drops: combinedTotal };
  }

  // ── 6) Rejection -> Re-delivery — rejected orders in this period, whether they were
  // re-delivered afterward, and how many attempts it took, BU-wise. Matches
  // rejection_rows.order_no against order_tracking.order_code — IF the two files use the
  // same order-numbering scheme. If Rejection YTD's order numbers use a different scheme
  // than the dispatch file's Order Code / Task ID, this section will show few or no
  // matches rather than wrong numbers — worth a quick check against a known real case
  // once both are loaded for the same period. ──
  async function buildRejectionRedelivery(dateFrom, dateTo) {
    var rejR = await pool.query(
      `SELECT order_no, org, entry_date::text AS entry_date FROM rejection_rows
       WHERE status = 'rej' AND entry_date >= $1 AND entry_date <= $2 AND order_no IS NOT NULL AND order_no <> ''`,
      [dateFrom, dateTo]);
    if (!rejR.rows.length) return { by_bu: [], rows: [], total_rejected: 0, total_redelivered: 0, note: 'No rejection rows with an order number in this period.' };
    var orderNos = rejR.rows.map(function (r) { return r.order_no; });
    var trackR = await pool.query(
      `SELECT order_code, date_key::text AS date_key FROM order_tracking WHERE order_code = ANY($1) ORDER BY date_key ASC`,
      [orderNos]);
    var deliveriesByCode = {};
    trackR.rows.forEach(function (t) {
      if (!deliveriesByCode[t.order_code]) deliveriesByCode[t.order_code] = [];
      deliveriesByCode[t.order_code].push(t.date_key);
    });
    var byBu = {}; // org -> { attempt_1, attempt_2, attempt_3_plus, not_redelivered, total }
    var detailRows = [];
    var totalRedelivered = 0;
    rejR.rows.forEach(function (rej) {
      var org = rej.org || '(Unknown)';
      if (!byBu[org]) byBu[org] = { attempt_1: 0, attempt_2: 0, attempt_3_plus: 0, not_redelivered: 0, total: 0 };
      byBu[org].total++;
      var afterDates = (deliveriesByCode[rej.order_no] || []).filter(function (d) { return d > rej.entry_date; });
      var attempts = afterDates.length;
      var status;
      if (attempts === 0) { byBu[org].not_redelivered++; status = 'Not yet re-delivered'; }
      else {
        totalRedelivered++;
        if (attempts === 1) byBu[org].attempt_1++;
        else if (attempts === 2) byBu[org].attempt_2++;
        else byBu[org].attempt_3_plus++;
        status = attempts + ' attempt(s) after rejection';
      }
      detailRows.push({ order_no: rej.order_no, org: org, rejected_date: rej.entry_date, redelivery_attempts: attempts, redelivery_dates: afterDates, status: status });
    });
    var buRows = Object.keys(byBu).sort().map(function (org) { return Object.assign({ org: org }, byBu[org]); });
    return { by_bu: buRows, rows: detailRows, total_rejected: rejR.rows.length, total_redelivered: totalRedelivered };
  }

  // ── Combined summary endpoint ──
  async function buildSummary(q) {
    var period = resolvePeriod(q);
    var rows = await fetchPeriodRows(period.date_from, period.date_to);
    var historyRows = await fetchHistoryUpTo(period.date_to);
    var valueTier = buildValueTier(rows);
    var lowValue = buildLowValue(valueTier);
    var redelivery = buildRedelivery(rows, historyRows, period.date_from, period.date_to);
    var drops = buildDrops(rows);
    var rejectionRedelivery = await buildRejectionRedelivery(period.date_from, period.date_to);
    var orgs = Array.from(new Set(rows.map(function (r) { return r.org || '(Unknown)'; }))).sort();
    return { period: period, orgs: orgs, value_tier: valueTier, low_value: lowValue, redelivery: redelivery, drops: drops, rejection_redelivery: rejectionRedelivery };
  }

  app.get('/api/transport-cost/summary', requireAuth, async function (req, res) {
    try { res.json(await buildSummary(req.query)); }
    catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Excel export — one workbook, one sheet per section, full color coding. ──
  app.get('/api/transport-cost/export', requireAuth, async function (req, res) {
    try {
      var ExcelJS = require('exceljs');
      var data = await buildSummary(req.query);
      var RED = 'FFF8D7DA', AMBER = 'FFFDF2CE', GREEN = 'FFDCF5E3';
      var wb = new ExcelJS.Workbook();
      wb.creator = 'AZHAR-AI'; wb.created = new Date();

      function titleRow(ws, text, span) { var t = ws.addRow([text]); ws.mergeCells('A' + t.number + ':' + span + t.number); t.font = { bold: true, size: 14 }; return t; }
      function headerRow(ws, cells) { var r = ws.addRow(cells); r.font = { bold: true }; r.eachCell(function (c) { c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F2B44' } }; c.font = { bold: true, color: { argb: 'FFFFFFFF' } }; }); return r; }

      // Sheet 1: Executive Summary
      var es = wb.addWorksheet('Executive Summary');
      es.columns = [{ width: 34 }, { width: 22 }];
      titleRow(es, 'TRANSPORT COST CONTROL — EXECUTIVE SUMMARY', 'B');
      es.addRow(['Period', data.period.label + ' (' + data.period.date_from + ' to ' + data.period.date_to + ')']);
      es.addRow(['Downloaded', new Date().toLocaleString('en-AE')]);
      es.addRow([]);
      headerRow(es, ['Metric', 'Value']);
      es.addRow(['Orders \u2265 AED 1,000 (Combined)', data.value_tier.combined.orders_ge1k]);
      es.addRow(['Orders < AED 1,000 (Combined)', data.value_tier.combined.orders_lt1k]);
      es.addRow(['Drops \u2265 AED 1,000 (Combined)', data.value_tier.combined.drops_ge1k]);
      es.addRow(['Drops < AED 1,000 (Combined)', data.value_tier.combined.drops_lt1k]);
      var lowRow = es.addRow(['Low-Value Orders % of All Orders (Combined)', data.low_value.combined.low_value_pct + '%']);
      if (data.low_value.combined.low_value_pct > 40) lowRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      es.addRow(['Total Drops (Combined)', data.drops.combined_total_drops]);
      es.addRow(['Orders Re-delivered', data.redelivery.total_orders_redelivered]);
      var costRow = es.addRow(['Estimated Re-delivery Cost (AED)', data.redelivery.total_estimated_cost]);
      if (data.redelivery.total_estimated_cost > 0) costRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      es.addRow(['Rejected Orders in Period', data.rejection_redelivery.total_rejected]);
      es.addRow(['...of which Re-delivered', data.rejection_redelivery.total_redelivered]);
      es.addRow([]);
      var noteRow = es.addRow([data.redelivery.note]); es.mergeCells('A' + noteRow.number + ':B' + noteRow.number); noteRow.font = { italic: true, size: 10, color: { argb: 'FF888888' } };

      // Sheet 2: Value Tier (Food/Non-Food x >=1k/<1k, BU-wise + Combined)
      var vt = wb.addWorksheet('Value Tier');
      vt.columns = [{ width: 16 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 16 }];
      titleRow(vt, 'ORDER VALUE TIER \u2014 ' + data.period.label, 'H');
      vt.addRow([]);
      headerRow(vt, ['BU', 'Type', 'Orders \u22651k', 'Orders <1k', 'Drops \u22651k', 'Drops <1k', 'Value \u22651k (AED)', 'Value <1k (AED)']);
      data.value_tier.rows.concat([data.value_tier.combined]).forEach(function (r) {
        var isCombined = r.org.indexOf('Combined') !== -1;
        var row = vt.addRow([r.org, r.type, r.orders_ge1k, r.orders_lt1k, r.drops_ge1k, r.drops_lt1k, r.value_ge1k, r.value_lt1k]);
        row.getCell(7).numFmt = '#,##0'; row.getCell(8).numFmt = '#,##0';
        if (isCombined) row.eachCell(function (c) { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; });
      });

      // Sheet 3: Low-Value Orders
      var lv = wb.addWorksheet('Low-Value Orders');
      lv.columns = [{ width: 16 }, { width: 14 }, { width: 16 }, { width: 14 }, { width: 14 }, { width: 18 }, { width: 16 }, { width: 20 }];
      titleRow(lv, 'LOW-VALUE ORDERS (< AED 1,000) \u2014 ' + data.period.label, 'H');
      lv.addRow([]);
      headerRow(lv, ['BU', 'Type', 'Low-Value Orders', 'Total Orders', 'Low-Value %', 'Low-Value Total (AED)', 'Total Value (AED)', 'Low-Value Share of Value %']);
      data.low_value.rows.concat([data.low_value.combined]).forEach(function (r) {
        var isCombined = r.org.indexOf('Combined') !== -1;
        var row = lv.addRow([r.org, r.type, r.low_value_orders, r.total_orders, r.low_value_pct + '%', r.low_value_total_aed, r.total_value_aed, r.low_value_share_of_value_pct + '%']);
        row.getCell(6).numFmt = '#,##0'; row.getCell(7).numFmt = '#,##0';
        var fill = r.low_value_pct >= 50 ? RED : (r.low_value_pct >= 25 ? AMBER : GREEN);
        row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        if (isCombined) row.eachCell(function (c) { c.font = { bold: true }; });
      });

      // Sheet 4: Re-delivery
      var rd = wb.addWorksheet('Re-delivery');
      rd.columns = [{ width: 14 }, { width: 20 }, { width: 18 }];
      titleRow(rd, 'RE-DELIVERY \u2014 ' + data.period.label, 'C');
      rd.addRow([]);
      rd.addRow(['Attempt Distribution', '', '']);
      headerRow(rd, ['Attempt #', 'Count', '']);
      rd.addRow(['2nd attempt', data.redelivery.by_attempt.second, '']);
      rd.addRow(['3rd attempt', data.redelivery.by_attempt.third, '']);
      rd.addRow(['4th+ attempt', data.redelivery.by_attempt.fourth_plus, '']);
      rd.addRow([]);
      headerRow(rd, ['Date', 'Orders Re-delivered', 'Estimated Cost (AED)']);
      data.redelivery.daily.forEach(function (r) {
        var row = rd.addRow([r.date, r.orders_redelivered, r.estimated_cost]);
        row.getCell(3).numFmt = '#,##0';
        if (r.estimated_cost > 0) row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      });
      rd.addRow([]);
      headerRow(rd, ['BU', 'Orders Re-delivered', 'Estimated Cost (AED)']);
      data.redelivery.by_bu.forEach(function (r) {
        var row = rd.addRow([r.org, r.orders_redelivered, r.estimated_cost]);
        row.getCell(3).numFmt = '#,##0';
        if (r.estimated_cost > 0) row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      });

      // Sheet 5: Drops
      var dr = wb.addWorksheet('Drops');
      dr.columns = [{ width: 14 }, { width: 16 }, { width: 12 }];
      titleRow(dr, 'FULL DROPS \u2014 ' + data.period.label, 'C');
      dr.addRow([]);
      headerRow(dr, ['BU', '', 'Total Drops']);
      data.drops.totals_by_bu.forEach(function (r) { dr.addRow([r.org, '', r.drops]); });
      var combinedDropsRow = dr.addRow(['ALL (Combined)', '', data.drops.combined_total_drops]); combinedDropsRow.eachCell(function (c) { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; });
      dr.addRow([]);
      headerRow(dr, ['Date', 'BU', 'Drops']);
      data.drops.daily_by_bu.forEach(function (r) { dr.addRow([r.date, r.org, r.drops]); });

      // Sheet 6: Rejection -> Re-delivery
      var rr = wb.addWorksheet('Rejection to Re-delivery');
      rr.columns = [{ width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 12 }];
      titleRow(rr, 'REJECTED ORDERS \u2014 RE-DELIVERY STATUS \u2014 ' + data.period.label, 'F');
      rr.addRow([]);
      headerRow(rr, ['BU', '1 Attempt', '2 Attempts', '3+ Attempts', 'Not Yet Re-delivered', 'Total Rejected']);
      data.rejection_redelivery.by_bu.forEach(function (r) {
        var row = rr.addRow([r.org, r.attempt_1, r.attempt_2, r.attempt_3_plus, r.not_redelivered, r.total]);
        if (r.not_redelivered > 0) row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      });
      rr.addRow([]);
      var rrNote = rr.addRow(['Order-number matching depends on Rejection YTD and Daily Dispatch using the same order-numbering scheme \u2014 verify against a known case if this section looks sparse.']);
      rr.mergeCells('A' + rrNote.number + ':F' + rrNote.number); rrNote.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
      rr.addRow([]);
      headerRow(rr, ['Order No', 'BU', 'Rejected Date', 'Re-delivery Attempts', 'Status']);
      data.rejection_redelivery.rows.forEach(function (r) {
        var row = rr.addRow([r.order_no, r.org, r.rejected_date, r.redelivery_attempts, r.status]);
        if (r.redelivery_attempts === 0) row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
        else row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
      });

      var buf = await wb.xlsx.writeBuffer();
      res.setHeader('Content-Disposition', 'attachment; filename="Transport_Cost_Control_' + data.period.type + '_' + Date.now() + '.xlsx"');
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(Buffer.from(buf));
    } catch (e) { res.status(500).json({ error: 'Export failed: ' + e.message }); }
  });

};
