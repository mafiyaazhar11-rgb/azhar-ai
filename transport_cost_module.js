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

  // ── Period resolution: daily (a single date), monthly (a calendar month), YTD (Jan 1
  // of the given year through today, or through Dec 31 for a past year), or custom (any
  // From/To range — multiple days, or several months like May through July at once). ──
  function resolvePeriod(q) {
    var period = q.period || 'daily';
    var today = new Date().toISOString().split('T')[0];
    if (period === 'custom') {
      var from = q.date_from || today, to = q.date_to || today;
      if (to < from) { var tmp = to; to = from; from = tmp; } // guard against a reversed range
      var fromLabel = new Date(from + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      var toLabel = new Date(to + 'T00:00:00').toLocaleDateString('en-GB', { day: 'numeric', month: 'short', year: 'numeric' });
      return { type: 'custom', date_from: from, date_to: to, label: fromLabel + ' \u2013 ' + toLabel };
    }
    if (period === 'monthly') {
      var month = q.month || today.slice(0, 7); // 'YYYY-MM'
      var y = +month.slice(0, 4), m = +month.slice(5, 7);
      var lastDay = new Date(y, m, 0).getDate();
      return { type: 'monthly', date_from: month + '-01', date_to: month + '-' + String(lastDay).padStart(2, '0'), label: month };
    }
    if (period === 'ytd') {
      var year = q.year || today.slice(0, 4);
      var to2 = (year === today.slice(0, 4)) ? today : (year + '-12-31');
      return { type: 'ytd', date_from: year + '-01-01', date_to: to2, label: year + ' YTD' };
    }
    var d = q.date || today;
    return { type: 'daily', date_from: d, date_to: d, label: d };
  }

  // ── Averages — mainly useful for a multi-day/multi-month Custom range: raw totals
  // over 3 months are hard to compare with a single day's totals, so this expresses the
  // same numbers as a per-day and per-month rate, using the number of days/months that
  // ACTUALLY have data (not the calendar span), so a partial month doesn't skew the
  // average low. ──
  function buildAverages(valueTier, drops, redelivery, distinctDates) {
    var days = distinctDates.length;
    var months = Array.from(new Set(distinctDates.map(function (d) { return d.slice(0, 7); }))).length;
    function per(total, n) { return n ? +(total / n).toFixed(1) : 0; }
    var totalOrders = valueTier.combined.orders_ge1k + valueTier.combined.orders_lt1k;
    var totalValue = valueTier.combined.value_ge1k + valueTier.combined.value_lt1k;
    return {
      days_with_data: days, months_spanned: months,
      avg_orders_per_day: per(totalOrders, days),
      avg_drops_per_day: per(drops.combined_total_drops, days),
      avg_value_per_day_aed: Math.round(per(totalValue, days)),
      avg_orders_per_month: per(totalOrders, months),
      avg_drops_per_month: per(drops.combined_total_drops, months),
      avg_value_per_month_aed: Math.round(per(totalValue, months)),
      avg_low_value_pct: valueTier.combined.orders_ge1k + valueTier.combined.orders_lt1k
        ? +(( (valueTier.combined.orders_lt1k) / totalOrders) * 100).toFixed(1) : 0,
      avg_redelivery_cost_per_month_aed: Math.round(per(redelivery.total_estimated_cost, months))
    };
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
  // WHITELIST, not a blacklist — only these are real BUs we track transport cost against.
  // Anything else (3PL, TOUR-xxx trip/route codes, other divisions' org codes, etc.) gets
  // excluded automatically rather than needing a new blacklist entry every time a new
  // stray value shows up in the source files.
  var VALID_ORGS = ['DCV', 'DGC', 'DGS', 'DSN', 'DCF', 'DPS'];
  function isExcludedRow(row) {
    if (String(row.source_label || '').toUpperCase() === 'SALON') return true;
    if (VALID_ORGS.indexOf(String(row.org || '').toUpperCase().trim()) === -1) return true;
    // Internal transfers (warehouse-to-warehouse, hub-to-hub) carry no real transport
    // cost or customer-facing order value — they shouldn't count as a "low-value order"
    // or occupy a drop slot on this page.
    if (String(row.customer || '').toUpperCase().indexOf('INTERNAL') !== -1) return true;
    return false;
  }
  function bucketOf(dropType) {
    var t = String(dropType || '').toLowerCase();
    if (t === 'food') return 'Food';
    if (t === 'nonfood') return 'Non-Food';
    return 'Other';
  }
  // Frozen vs Ambient — the dimension that actually decides whether multiple orders to
  // the same stop can travel (and be COMBINED, see below) on the same truck.
  function tempBucketOf(temperatureRaw, truckTypeRaw) {
    var t = (String(temperatureRaw || '') + ' ' + String(truckTypeRaw || '')).toUpperCase();
    if (t.indexOf('FROZEN') !== -1 || t.indexOf('FREEZ') !== -1) return 'Frozen';
    return 'Ambient'; // covers Ambient, Chiller, and unspecified
  }

  // ── COMBINED DROPS — the core building block for sections 1, 2, 3 and 5.
  // "Combined" means: multiple orders (e.g. AED 500 + 200 + 300) sharing the SAME
  // (date, route, location) — i.e. the same truck's stop at that address that day — ride
  // together and their VALUES ARE SUMMED before checking the AED 1,000 threshold. A drop
  // is judged as a whole, not order-by-order.
  //
  // The grouping key is (date, route, location) — the EXACT definition already used and
  // validated in Transport Cost Reconciliation (total_drops_reconcile), which is checked
  // against transport's own reported numbers. An earlier version of this grouped by
  // (date, location, temperature) instead, which silently collapsed two separate route
  // visits to the same address on the same day into one drop whenever they happened to
  // share a temperature category — undercounting real truck stops. Matching the proven
  // definition here fixes that: two different ROUTES to the same location on the same
  // day are correctly two separate drops (and two separate charges), never combined.
  // Rows with no route on file aren't grouped with anything else at that location (each
  // becomes its own drop, keyed by order code) rather than risk merging genuinely
  // separate visits under a shared blank key. ──
  function buildCombinedDrops(rows) {
    var groups = {};
    rows.forEach(function (r) {
      if (isExcludedRow(r)) return;
      var loc = r.location_id || r.customer || r.order_code;
      var routeKey = r.route ? r.route : ('(no route)::' + r.order_code);
      var key = r.date_key + '::' + routeKey + '::' + loc;
      if (!groups[key]) groups[key] = { date: r.date_key, location: loc, org: r.org || '(Unknown)', value: 0, orders: [] };
      groups[key].value += (+r.value || 0);
      groups[key].orders.push(r);
    });
    // Temperature (Ambient/Frozen) is a per-group DISPLAY label, not part of the grouping
    // key — a route's truck is one temperature type in practice, so the first order's
    // temperature represents the whole group.
    return Object.keys(groups).map(function (k) {
      var g = groups[k];
      g.temp = tempBucketOf(g.orders[0].temperature, g.orders[0].truck_type);
      return g;
    });
  }

  // ── 1+2) Value-tier split: orders & drops at/above AED 1,000 vs below, by BU, by
  // temperature (Ambient/Frozen — shown as the Type label, since a route's truck is one
  // temperature in practice), plus one TOTAL (All BU) row. Classification happens at the
  // COMBINED DROP level (see above, grouped by date+route+location): a drop with 3
  // orders worth 500+200+300 = AED 1,000 counts as one ≥1k drop with 3 orders in it — the
  // orders are not each judged against the threshold individually. ──
  function buildValueTier(rows, threshold) {
    threshold = threshold || LOW_VALUE_THRESHOLD;
    var drops = buildCombinedDrops(rows);
    var byOrgTemp = {};
    var total = { orders_hi: 0, orders_lo: 0, value_hi: 0, value_lo: 0, drops_hi: 0, drops_lo: 0, pure_drops_lo: 0, pure_orders_lo: 0 };
    drops.forEach(function (g) {
      var key = g.org + '||' + g.temp;
      if (!byOrgTemp[key]) byOrgTemp[key] = { org: g.org, type: g.temp, orders_hi: 0, orders_lo: 0, value_hi: 0, value_lo: 0, drops_hi: 0, drops_lo: 0, pure_drops_lo: 0, pure_orders_lo: 0 };
      var b = byOrgTemp[key];
      var isHigh = g.value >= threshold;
      if (isHigh) { b.orders_hi += g.orders.length; b.value_hi += g.value; b.drops_hi++; total.orders_hi += g.orders.length; total.value_hi += g.value; total.drops_hi++; }
      else {
        b.orders_lo += g.orders.length; b.value_lo += g.value; b.drops_lo++;
        total.orders_lo += g.orders.length; total.value_lo += g.value; total.drops_lo++;
        // "Pure" = a low-value drop that was NEVER combined with another order — it went
        // out as its own dedicated truck stop, exactly as if there were no consolidation
        // logic at all. This is the actual current spend on standalone low-value drops,
        // separate from the ones that already got folded into someone else's stop.
        if (g.orders.length === 1) { b.pure_drops_lo++; b.pure_orders_lo++; total.pure_drops_lo++; total.pure_orders_lo++; }
      }
    });
    var out = Object.keys(byOrgTemp).sort().map(function (k) {
      var b = byOrgTemp[k];
      return { org: b.org, type: b.type, orders_ge1k: b.orders_hi, orders_lt1k: b.orders_lo, drops_ge1k: b.drops_hi, drops_lt1k: b.drops_lo, value_ge1k: Math.round(b.value_hi), value_lt1k: Math.round(b.value_lo), pure_drops_lt1k: b.pure_drops_lo, pure_orders_lt1k: b.pure_orders_lo };
    });
    return {
      rows: out, threshold: threshold,
      combined: { org: 'TOTAL (All BU)', type: 'All Types', orders_ge1k: total.orders_hi, orders_lt1k: total.orders_lo, drops_ge1k: total.drops_hi, drops_lt1k: total.drops_lo, value_ge1k: Math.round(total.value_hi), value_lt1k: Math.round(total.value_lo), pure_drops_lt1k: total.pure_drops_lo, pure_orders_lt1k: total.pure_orders_lo }
    };
  }

  // ── 3) Low-value order tracker — same numbers as buildValueTier, reshaped to answer
  // "how many low-value orders go out compared with [total] value", BU-wise + combined.
  // "Low-value" here means: part of a drop whose COMBINED value still didn't reach
  // AED 1,000, not that the individual order itself was under AED 1,000. ──
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

  // ── 3a) Low-value order CONSOLIDATION SAVINGS — directly answers "how many drops
  // have been implemented for low-value orders, and what's the saving on transport
  // cost." A low-value order isn't automatically a separate truck stop: if 2+ of them
  // share the same date/route/location, they already ride on ONE drop together (see
  // buildValueTier/buildCombinedDrops above). So low_value_orders minus low_value_drops
  // is exactly the count of orders that got folded into a drop with another order,
  // instead of needing a truck stop of their own — that's the real consolidation
  // number, and each one saved is one fewer drop charge. Cost uses the same per-drop
  // rate card already used for Re-attempt Cost (temperature-based Multi-route rate),
  // so it's consistent with every other estimate on this page — not a new made-up rate. ──
  function buildLowValueConsolidation(valueTier) {
    function shape(row) {
      var ordersSaved = Math.max(0, row.orders_lt1k - row.drops_lt1k);
      var rate = estimateRedeliveryRate(row.type, row.type); // row.type is 'Ambient'/'Frozen' already
      return {
        org: row.org, type: row.type,
        low_value_orders: row.orders_lt1k, low_value_drops: row.drops_lt1k,
        orders_saved_by_consolidation: ordersSaved,
        estimated_cost_saved_aed: Math.round(ordersSaved * rate),
        // "Pure" low-value orders — sent as their own dedicated drop, never combined
        // with anything else. This is the current, ongoing transport cost being spent
        // on standalone low-value orders — separate from what's already being saved
        // by consolidation above.
        pure_low_value_orders: row.pure_orders_lt1k || 0,
        pure_low_value_drops: row.pure_drops_lt1k || 0,
        pure_drop_current_cost_aed: Math.round((row.pure_drops_lt1k || 0) * rate)
      };
    }
    return { rows: valueTier.rows.map(shape), combined: shape(valueTier.combined) };
  }

  // ── 3b) Low-value orders RANKED BY CUSTOMER, BU-wise — the BU/type roll-up above
  // tells you a BU has a low-value problem; this names exactly which customer/branch is
  // causing it, so a specific account can be approached about a minimum order value
  // instead of chasing the number in the abstract. An order counts as "low-value" here
  // if the COMBINED drop it was part of stayed under AED 1,000 (same rule as above).
  // Ranked by low-value order COUNT first (the thing driving avoidable drop cost), then
  // by low-value AED as a tiebreak. Only customers with at least one low-value order are
  // included. ──
  function buildLowValueByCustomer(rows) {
    var drops = buildCombinedDrops(rows);
    var byKey = {};
    drops.forEach(function (g) {
      var isLow = g.value < LOW_VALUE_THRESHOLD;
      g.orders.forEach(function (r) {
        var org = r.org || '(Unknown)';
        var customer = r.customer || '(No customer name on file)';
        var key = org + '||' + customer;
        if (!byKey[key]) byKey[key] = { org: org, customer: customer, low: 0, total: 0, lowValueSum: 0, totalValueSum: 0 };
        var b = byKey[key];
        var val = +r.value || 0;
        b.total++; b.totalValueSum += val;
        if (isLow) { b.low++; b.lowValueSum += val; }
      });
    });
    return Object.keys(byKey).map(function (k) {
      var b = byKey[k];
      return {
        org: b.org, customer: b.customer, low_value_orders: b.low, total_orders: b.total,
        low_value_pct: b.total ? +((b.low / b.total) * 100).toFixed(1) : 0,
        low_value_total_aed: Math.round(b.lowValueSum), total_value_aed: Math.round(b.totalValueSum)
      };
    }).filter(function (r) { return r.low_value_orders > 0; })
      .sort(function (a, b) { return b.low_value_orders - a.low_value_orders || b.low_value_total_aed - a.low_value_total_aed; });
  }

  // ── 4) Re-delivery tracking — attempt number = position of this occurrence in that
  // order's full date history up to date_to. attempt_number > 1 means this occurrence IS
  // a re-delivery (the earlier attempt(s) failed / weren't collected as delivered).
  // COST IS PER DROP, NOT PER ORDER: if 10 re-delivered orders all land on the same
  // location on the same day, that's one truck stop and one drop charge — not ten. Cost
  // is charged exactly once per distinct (date, location) among the re-delivered rows. ──
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
    var costedDrops = new Set(); // "date::location" already charged — one charge per drop
    rows.forEach(function (r) {
      if (isExcludedRow(r)) return;
      var dates = historyByCode[r.order_code] || [];
      var attemptNumber = dates.indexOf(r.date_key) + 1;
      if (attemptNumber <= 1) return; // first-ever attempt — not a re-delivery
      var dedupeKey = r.order_code + '::' + r.date_key;
      if (seenOrderAttempt.has(dedupeKey)) return;
      seenOrderAttempt.add(dedupeKey);
      var dropKey = r.date_key + '::' + (r.location_id || r.customer || r.order_code);
      var cost = 0;
      if (!costedDrops.has(dropKey)) { costedDrops.add(dropKey); cost = estimateRedeliveryRate(r.truck_type, r.temperature); }
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
    return { daily: dailyRows, by_bu: buRows, by_attempt: byAttempt, total_orders_redelivered: totalOrders.size, total_estimated_cost: Math.round(totalCost), note: 'This counts RE-ATTEMPTS, not confirmed re-deliveries — an order appearing in dispatch again on a later date, since the dispatch file doesn\'t carry a delivered-successfully confirmation. Cost is billed PER DROP, not per order — 10 re-attempted orders to the same stop on the same day is one drop charge, not ten. Rate is an ESTIMATE (temperature-based Multi-route rate) — no per-order truck-type was on file for every row.' };
  }

  // ── 5) Full drops — day-wise and BU-wise. A "drop" is a COMBINED DROP (see above):
  // one truck stop = one (date, route, location) combo, regardless of how many orders
  // shared that stop — the same definition already validated against transport's own
  // reported drop counts in Transport Cost Reconciliation. ──
  function buildDrops(rows) {
    var drops = buildCombinedDrops(rows);
    var dailyBu = {}; // "date||org" -> count
    var totalsByBu = {}; // org -> count
    drops.forEach(function (g) {
      var dbKey = g.date + '||' + g.org;
      dailyBu[dbKey] = (dailyBu[dbKey] || 0) + 1;
      totalsByBu[g.org] = (totalsByBu[g.org] || 0) + 1;
    });
    var dailyRows = Object.keys(dailyBu).sort().map(function (k) {
      var parts = k.split('||'); return { date: parts[0], org: parts[1], drops: dailyBu[k] };
    });
    var buRows = Object.keys(totalsByBu).sort().map(function (org) { return { org: org, drops: totalsByBu[org] }; });
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
    // Same DGC->DCV remap and DPS exclusion as everywhere else on this page (see
    // buildSummary) — this section queries rejection_rows directly, so it needs the
    // same correction applied separately rather than inheriting it from `rows`.
    rejR.rows = rejR.rows
      .filter(function (r) { return VALID_ORGS.indexOf(String(r.org || '').toUpperCase().trim()) !== -1; })
      .filter(function (r) { return String(r.org || '').toUpperCase().trim() !== 'DPS'; })
      .map(function (r) {
        if (String(r.org || '').toUpperCase().trim() === 'DGC') return Object.assign({}, r, { org: 'DCV' });
        return r;
      });
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
      if (attempts === 0) { byBu[org].not_redelivered++; status = 'Not yet re-attempted'; }
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
    // Business rule: DGC is not tracked as its own BU on this page — anything that
    // comes in tagged DGC gets folded into DCV's Frozen bucket instead (org forced to
    // DCV, temperature forced to Frozen, regardless of what the row's own temperature/
    // truck_type field says). DPS is dropped entirely, not just hidden — it never
    // contributes to any total, drop count, or value figure anywhere on this page.
    // Applied once, right here, before any section is built, so every downstream
    // number (Value Tier, Low-Value, Consolidation, Re-delivery, Drops, Rejection
    // Re-delivery) reflects this consistently — nowhere on the page can show raw DGC
    // or any DPS data by only fixing one section and missing another.
    rows = rows.filter(function (r) { return String(r.org || '').toUpperCase().trim() !== 'DPS'; })
      .map(function (r) {
        if (String(r.org || '').toUpperCase().trim() === 'DGC') {
          return Object.assign({}, r, { org: 'DCV', temperature: 'Frozen', truck_type: 'Frozen' });
        }
        return r;
      });
    // Optional Food / Non-Food filter — applied once, up front, so every section below
    // (value tier, low-value, re-delivery, drops, rejection re-delivery) sees only the
    // rows that match, with no risk of some sections honoring it and others not.
    if (q.order_type === 'food' || q.order_type === 'nonfood') {
      rows = rows.filter(function (r) { return String(r.drop_type || '').toLowerCase() === q.order_type; });
    }
    // Diagnostic: how many distinct dates actually have data in this period, vs how many
    // calendar days the period spans — a low "days with data" number explains a low
    // Total Drops figure for a month far more often than a calculation bug does (e.g. if
    // dispatch was only uploaded for 2 of that month's days, the drop count for that
    // month IS roughly 2 days' worth, correctly).
    var distinctDates = Array.from(new Set(rows.map(function (r) { return r.date_key; }))).sort();
    var calendarDays = Math.round((new Date(period.date_to) - new Date(period.date_from)) / 86400000) + 1;
    var historyRows = await fetchHistoryUpTo(period.date_to);
    // Optional what-if: a custom minimum-order-value threshold, so "what would we save
    // if we raised the minimum to AED 1,500" can be tested directly instead of guessed.
    // Defaults to the standard AED 1,000 (LOW_VALUE_THRESHOLD) when not provided.
    var customThreshold = q.threshold ? Math.max(0, parseInt(q.threshold)) : LOW_VALUE_THRESHOLD;
    var valueTier = buildValueTier(rows, customThreshold);
    var lowValue = buildLowValue(valueTier);
    var lowValueConsolidation = buildLowValueConsolidation(valueTier);
    var lowValueByCustomer = buildLowValueByCustomer(rows);
    var redelivery = buildRedelivery(rows, historyRows, period.date_from, period.date_to);
    var drops = buildDrops(rows);
    var rejectionRedelivery = await buildRejectionRedelivery(period.date_from, period.date_to);
    var orgs = Array.from(new Set(rows.filter(function (r) { return !isExcludedRow(r); }).map(function (r) { return r.org || '(Unknown)'; }))).sort();
    var averages = buildAverages(valueTier, drops, redelivery, distinctDates);
    return {
      period: period, orgs: orgs, value_tier: valueTier, low_value: lowValue, low_value_consolidation: lowValueConsolidation, low_value_by_customer: lowValueByCustomer,
      redelivery: redelivery, drops: drops, rejection_redelivery: rejectionRedelivery, averages: averages,
      data_coverage: { days_with_data: distinctDates.length, calendar_days: calendarDays, dates_with_data: distinctDates }
    };
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
      var covRow = es.addRow(['Days With Data', data.data_coverage.days_with_data + ' of ' + data.data_coverage.calendar_days + ' calendar days in this period']);
      if (data.data_coverage.days_with_data < data.data_coverage.calendar_days) covRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } };
      es.addRow([]);
      headerRow(es, ['Metric', 'Value']);
      es.addRow(['Orders \u2265 AED 1,000 (All BU)', data.value_tier.combined.orders_ge1k]);
      es.addRow(['Orders < AED 1,000 (All BU)', data.value_tier.combined.orders_lt1k]);
      es.addRow(['Drops \u2265 AED 1,000 (All BU)', data.value_tier.combined.drops_ge1k]);
      es.addRow(['Drops < AED 1,000 (All BU)', data.value_tier.combined.drops_lt1k]);
      var lowRow = es.addRow(['Low-Value Orders % of All Orders (All BU)', data.low_value.combined.low_value_pct + '%']);
      if (data.low_value.combined.low_value_pct > 40) lowRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      var savedRow = es.addRow(['Orders Saved by Consolidation (All BU)', data.low_value_consolidation.combined.orders_saved_by_consolidation]);
      if (data.low_value_consolidation.combined.orders_saved_by_consolidation > 0) savedRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
      var savedCostRow = es.addRow(['Estimated Cost Saved by Consolidation (AED)', data.low_value_consolidation.combined.estimated_cost_saved_aed]);
      if (data.low_value_consolidation.combined.estimated_cost_saved_aed > 0) savedCostRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
      var pureRow = es.addRow(['Pure Orders \u2014 Sent Standalone, Never Combined (All BU)', data.low_value_consolidation.combined.pure_low_value_orders]);
      if (data.low_value_consolidation.combined.pure_low_value_orders > 0) pureRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } };
      var pureCostRow = es.addRow(['Current Cost of Standalone Low-Value Drops (AED)', data.low_value_consolidation.combined.pure_drop_current_cost_aed]);
      if (data.low_value_consolidation.combined.pure_drop_current_cost_aed > 0) pureCostRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      es.addRow(['Total Drops (All BU)', data.drops.combined_total_drops]);
      es.addRow(['Orders Re-attempted', data.redelivery.total_orders_redelivered]);
      var costRow = es.addRow(['Estimated Re-delivery Cost (AED)', data.redelivery.total_estimated_cost]);
      if (data.redelivery.total_estimated_cost > 0) costRow.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      es.addRow(['Rejected Orders in Period', data.rejection_redelivery.total_rejected]);
      es.addRow(['...of which Re-attempted', data.rejection_redelivery.total_redelivered]);
      es.addRow([]);
      var noteRow = es.addRow([data.redelivery.note]); es.mergeCells('A' + noteRow.number + ':B' + noteRow.number); noteRow.font = { italic: true, size: 10, color: { argb: 'FF888888' } };

      // Sheet 2: Value Tier (Ambient/Frozen x >=1k/<1k, BU-wise + Total)
      var vt = wb.addWorksheet('Value Tier');
      vt.columns = [{ width: 16 }, { width: 14 }, { width: 14 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 16 }];
      titleRow(vt, 'ORDER VALUE TIER \u2014 ' + data.period.label, 'H');
      vt.addRow([]);
      headerRow(vt, ['BU', 'Type', 'Orders \u22651k', 'Orders <1k', 'Drops \u22651k', 'Drops <1k', 'Value \u22651k (AED)', 'Value <1k (AED)']);
      data.value_tier.rows.concat([data.value_tier.combined]).forEach(function (r) {
        var isTotal = r.org.indexOf('TOTAL') !== -1;
        var row = vt.addRow([r.org, r.type, r.orders_ge1k, r.orders_lt1k, r.drops_ge1k, r.drops_lt1k, r.value_ge1k, r.value_lt1k]);
        row.getCell(7).numFmt = '#,##0'; row.getCell(8).numFmt = '#,##0';
        if (isTotal) row.eachCell(function (c) { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; });
      });

      // Sheet 3: Low-Value Orders
      var lv = wb.addWorksheet('Low-Value Orders');
      lv.columns = [{ width: 16 }, { width: 14 }, { width: 16 }, { width: 14 }, { width: 14 }, { width: 18 }, { width: 16 }, { width: 20 }];
      titleRow(lv, 'LOW-VALUE ORDERS (< AED 1,000) \u2014 ' + data.period.label, 'H');
      lv.addRow([]);
      headerRow(lv, ['BU', 'Type', 'Low-Value Orders', 'Total Orders', 'Low-Value %', 'Low-Value Total (AED)', 'Total Value (AED)', 'Low-Value Share of Value %']);
      data.low_value.rows.concat([data.low_value.combined]).forEach(function (r) {
        var isTotal = r.org.indexOf('TOTAL') !== -1;
        var row = lv.addRow([r.org, r.type, r.low_value_orders, r.total_orders, r.low_value_pct + '%', r.low_value_total_aed, r.total_value_aed, r.low_value_share_of_value_pct + '%']);
        row.getCell(6).numFmt = '#,##0'; row.getCell(7).numFmt = '#,##0';
        var fill = r.low_value_pct >= 50 ? RED : (r.low_value_pct >= 25 ? AMBER : GREEN);
        row.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
        if (isTotal) row.eachCell(function (c) { c.font = { bold: true }; });
      });

      // Sheet 3a: Low-Value Order Consolidation Savings — how many low-value orders
      // got folded into a shared drop instead of needing their own truck stop, and
      // the estimated cost that saved.
      var lvs = wb.addWorksheet('Consolidation Savings');
      lvs.columns = [{ width: 16 }, { width: 14 }, { width: 18 }, { width: 16 }, { width: 22 }, { width: 20 }, { width: 20 }, { width: 24 }];
      titleRow(lvs, 'LOW-VALUE ORDER CONSOLIDATION SAVINGS \u2014 ' + data.period.label + ' (threshold: AED ' + data.value_tier.threshold.toLocaleString() + ')', 'H');
      lvs.addRow([]);
      var lvsNote = lvs.addRow(['A low-value order (below the threshold above) isn\'t automatically a separate truck stop \u2014 if two or more share the same date/route/location, they already ride on ONE drop together. "Saved by Consolidation" = orders that got folded into a shared drop. "Pure" = orders that went out as their OWN dedicated drop, never combined \u2014 that\'s the current, ongoing cost still being spent on standalone low-value orders. Same per-drop rate card used for Re-attempt Cost throughout \u2014 not a new made-up rate.']);
      lvs.mergeCells('A' + lvsNote.number + ':H' + lvsNote.number); lvsNote.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
      lvs.addRow([]);
      headerRow(lvs, ['BU', 'Type', 'Low-Value Orders', 'Low-Value Drops', 'Orders Saved by Consolidation', 'Estimated Cost Saved (AED)', 'Pure (Non-Combined) Orders', 'Current Cost of Pure Drops (AED)']);
      data.low_value_consolidation.rows.concat([data.low_value_consolidation.combined]).forEach(function (r) {
        var isTotal = r.org.indexOf('TOTAL') !== -1;
        var row = lvs.addRow([r.org, r.type, r.low_value_orders, r.low_value_drops, r.orders_saved_by_consolidation, r.estimated_cost_saved_aed, r.pure_low_value_orders, r.pure_drop_current_cost_aed]);
        row.getCell(6).numFmt = '#,##0'; row.getCell(8).numFmt = '#,##0';
        if (r.orders_saved_by_consolidation > 0) row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GREEN } };
        if (r.pure_low_value_orders > 0) row.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } };
        if (isTotal) row.eachCell(function (c) { c.font = { bold: true }; });
      });
      lvs.addRow([]);
      // Sheet 3b: Low-Value Orders by Customer — ranked, BU-wise, full list (not capped
      // to the top 50 shown on screen).
      var lvc = wb.addWorksheet('Low-Value by Customer');
      lvc.columns = [{ width: 8 }, { width: 34 }, { width: 12 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 18 }, { width: 16 }];
      titleRow(lvc, 'LOW-VALUE ORDERS RANKED BY CUSTOMER \u2014 ' + data.period.label, 'H');
      lvc.addRow([]);
      var lvcNote = lvc.addRow(['Ranked by number of low-value orders (the driver of avoidable drop cost), highest first. Only customers with at least one low-value order are listed.']);
      lvc.mergeCells('A' + lvcNote.number + ':H' + lvcNote.number); lvcNote.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
      lvc.addRow([]);
      headerRow(lvc, ['Rank', 'Customer', 'BU', 'Low-Value Orders', 'Total Orders', 'Low-Value %', 'Low-Value Total (AED)', 'Total Value (AED)']);
      data.low_value_by_customer.forEach(function (r, i) {
        var row = lvc.addRow([i + 1, r.customer, r.org, r.low_value_orders, r.total_orders, r.low_value_pct + '%', r.low_value_total_aed, r.total_value_aed]);
        row.getCell(7).numFmt = '#,##0'; row.getCell(8).numFmt = '#,##0';
        var fill = i < 10 ? RED : (i < 30 ? AMBER : null);
        if (fill) row.getCell(4).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      });
      if (!data.low_value_by_customer.length) lvc.addRow(['No low-value orders in this period.']);

      // Sheet 4: Re-delivery
      var rd = wb.addWorksheet('Re-attempts');
      rd.columns = [{ width: 14 }, { width: 20 }, { width: 18 }];
      titleRow(rd, 'RE-ATTEMPTS \u2014 ' + data.period.label, 'C');
      rd.addRow([]);
      rd.addRow(['Attempt Distribution', '', '']);
      headerRow(rd, ['Attempt #', 'Count', '']);
      rd.addRow(['2nd attempt', data.redelivery.by_attempt.second, '']);
      rd.addRow(['3rd attempt', data.redelivery.by_attempt.third, '']);
      rd.addRow(['4th+ attempt', data.redelivery.by_attempt.fourth_plus, '']);
      rd.addRow([]);
      headerRow(rd, ['Date', 'Orders Re-attempted', 'Estimated Cost (AED)']);
      data.redelivery.daily.forEach(function (r) {
        var row = rd.addRow([r.date, r.orders_redelivered, r.estimated_cost]);
        row.getCell(3).numFmt = '#,##0';
        if (r.estimated_cost > 0) row.getCell(3).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      });
      rd.addRow([]);
      headerRow(rd, ['BU', 'Orders Re-attempted', 'Estimated Cost (AED)']);
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
      var combinedDropsRow = dr.addRow(['TOTAL (All BU)', '', data.drops.combined_total_drops]); combinedDropsRow.eachCell(function (c) { c.font = { bold: true }; c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: AMBER } }; });
      dr.addRow([]);
      headerRow(dr, ['Date', 'BU', 'Drops']);
      data.drops.daily_by_bu.forEach(function (r) { dr.addRow([r.date, r.org, r.drops]); });

      // Sheet 5b: Averages — most useful for a multi-day/multi-month Custom range, where
      // a raw total (e.g. 3 months combined) is hard to compare against a single day.
      var av = wb.addWorksheet('Averages');
      av.columns = [{ width: 34 }, { width: 22 }];
      titleRow(av, 'AVERAGES \u2014 ' + data.period.label, 'B');
      av.addRow([]);
      var avNote = av.addRow(['Based on ' + data.averages.days_with_data + ' day(s) with data, spanning ' + data.averages.months_spanned + ' calendar month(s).']);
      av.mergeCells('A' + avNote.number + ':B' + avNote.number); avNote.font = { italic: true, size: 10, color: { argb: 'FF888888' } };
      av.addRow([]);
      headerRow(av, ['Metric', 'Value']);
      av.addRow(['Avg Orders / Day', data.averages.avg_orders_per_day]);
      av.addRow(['Avg Drops / Day', data.averages.avg_drops_per_day]);
      var avVpd = av.addRow(['Avg Order Value / Day (AED)', data.averages.avg_value_per_day_aed]); avVpd.getCell(2).numFmt = '#,##0';
      av.addRow([]);
      av.addRow(['Avg Orders / Month', data.averages.avg_orders_per_month]);
      av.addRow(['Avg Drops / Month', data.averages.avg_drops_per_month]);
      var avVpm = av.addRow(['Avg Order Value / Month (AED)', data.averages.avg_value_per_month_aed]); avVpm.getCell(2).numFmt = '#,##0';
      var avRpm = av.addRow(['Avg Re-attempt Cost / Month (AED)', data.averages.avg_redelivery_cost_per_month_aed]); avRpm.getCell(2).numFmt = '#,##0';
      if (data.averages.avg_redelivery_cost_per_month_aed > 0) avRpm.getCell(2).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: RED } };
      av.addRow(['Avg Low-Value Orders %', data.averages.avg_low_value_pct + '%']);

      // Sheet 6: Rejection -> Re-attempt
      var rr = wb.addWorksheet('Rejection to Re-attempt');
      rr.columns = [{ width: 16 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 12 }];
      titleRow(rr, 'REJECTED ORDERS \u2014 RE-ATTEMPT STATUS \u2014 ' + data.period.label, 'F');
      rr.addRow([]);
      headerRow(rr, ['BU', '1 Attempt', '2 Attempts', '3+ Attempts', 'Not Yet Re-attempted', 'Total Rejected']);
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
