// ══ TRIP EXPENSE SPLITTER ════════════════════════════════════════════════
// Friends' trip spend tracker. Create a trip, list who's joining (plain
// names — friends don't need app logins), log every spend against a
// member's name with a category (Hotel, Food, Entertainment, Car/Petrol,
// Other) and up to 3 receipt photos, then get a live gross total, an equal
// split, and — at the end of the trip — who overspent (gets money back)
// and who underspent (needs to pay in), plus a shortest "who pays whom"
// settlement list.
module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog) {

  (async function initTables() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS trip_trips (
        id SERIAL PRIMARY KEY,
        name TEXT NOT NULL,
        members JSONB NOT NULL DEFAULT '[]',
        status TEXT NOT NULL DEFAULT 'active',
        created_by_username TEXT NOT NULL,
        created_by_name TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE TABLE IF NOT EXISTS trip_expenses (
        id SERIAL PRIMARY KEY,
        trip_id INTEGER NOT NULL REFERENCES trip_trips(id) ON DELETE CASCADE,
        spender_name TEXT NOT NULL,
        category TEXT NOT NULL,
        description TEXT,
        amount NUMERIC(12,2) NOT NULL,
        expense_date DATE NOT NULL DEFAULT CURRENT_DATE,
        added_by_username TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE TABLE IF NOT EXISTS trip_expense_photos (
        id SERIAL PRIMARY KEY,
        expense_id INTEGER NOT NULL REFERENCES trip_expenses(id) ON DELETE CASCADE,
        photo_data BYTEA NOT NULL,
        photo_mime TEXT,
        position INTEGER NOT NULL DEFAULT 1,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_expenses_trip ON trip_expenses (trip_id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_photos_expense ON trip_expense_photos (expense_id)`);
      console.log('Trip Expense Splitter module DB ready');
    } catch (e) { console.error('Trip module init error:', e.message); }
  })();

  function toStr(v) { return String(v == null ? '' : v).trim(); }
  function toNum(v) { var n = parseFloat(v); return isNaN(n) ? 0 : n; }
  var CATEGORIES = ['Hotel', 'Food', 'Entertainment', 'Car/Petrol', 'Other'];

  async function loadTrip(id) {
    var r = await pool.query('SELECT * FROM trip_trips WHERE id=$1', [id]);
    return r.rows[0] || null;
  }

  // Anyone who set up or was added to a trip's data-entry crew can see/edit it.
  // Keep it simple: any authenticated user with trip_expense access can manage
  // any trip — this is a small-group friends feature, not per-user isolated data.
  async function requireTripAccess(req, res, next) {
    try {
      var trip = await loadTrip(req.params.tripId || req.params.id);
      if (!trip) return res.status(404).json({ error: 'Trip not found' });
      req.trip = trip;
      next();
    } catch (e) { res.status(500).json({ error: e.message }); }
  }

  // ── Create a trip ──
  app.post('/api/trips', requireAuth, async function (req, res) {
    try {
      var name = toStr(req.body.name);
      var members = Array.isArray(req.body.members) ? req.body.members.map(toStr).filter(Boolean) : [];
      if (!name) return res.status(400).json({ error: 'Trip name is required.' });
      if (members.length < 2) return res.status(400).json({ error: 'Add at least 2 members who joined the trip.' });
      // de-dupe, keep order
      var seen = {}; members = members.filter(function (m) { var k = m.toLowerCase(); if (seen[k]) return false; seen[k] = true; return true; });
      var r = await pool.query(
        `INSERT INTO trip_trips (name, members, created_by_username, created_by_name) VALUES ($1,$2,$3,$4) RETURNING id`,
        [name, JSON.stringify(members), req.user.username, req.user.full_name || req.user.username]
      );
      auditLog(null, req.user.username, 'TRIP_CREATED', 'Created trip "' + name + '" with ' + members.length + ' members', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true, id: r.rows[0].id });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── List trips ──
  app.get('/api/trips', requireAuth, async function (req, res) {
    try {
      var r = await pool.query(`
        SELECT t.*, COALESCE(e.total, 0) AS gross_total, COALESCE(e.cnt, 0) AS expense_count
        FROM trip_trips t
        LEFT JOIN (SELECT trip_id, SUM(amount) AS total, COUNT(*) AS cnt FROM trip_expenses GROUP BY trip_id) e ON e.trip_id = t.id
        ORDER BY (t.status = 'active') DESC, t.created_at DESC
      `);
      res.json({ trips: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Trip detail ──
  app.get('/api/trips/:id', requireAuth, requireTripAccess, async function (req, res) {
    try { res.json({ trip: req.trip }); } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Edit trip (name / members / status) ──
  app.put('/api/trips/:id', requireAuth, requireTripAccess, async function (req, res) {
    try {
      var name = toStr(req.body.name) || req.trip.name;
      var members = Array.isArray(req.body.members) ? req.body.members.map(toStr).filter(Boolean) : req.trip.members;
      var status = ['active', 'closed'].includes(req.body.status) ? req.body.status : req.trip.status;
      await pool.query('UPDATE trip_trips SET name=$1, members=$2, status=$3, updated_at=NOW() WHERE id=$4',
        [name, JSON.stringify(members), status, req.trip.id]);
      auditLog(null, req.user.username, 'TRIP_UPDATED', 'Updated trip "' + name + '" (' + status + ')', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Delete trip ──
  app.delete('/api/trips/:id', requireAuth, requireTripAccess, async function (req, res) {
    try {
      await pool.query('DELETE FROM trip_trips WHERE id=$1', [req.trip.id]);
      auditLog(null, req.user.username, 'TRIP_DELETED', 'Deleted trip "' + req.trip.name + '"', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Add an expense — up to 3 receipt photos ──
  app.post('/api/trips/:tripId/expenses', requireAuth, requireTripAccess, upload.array('photos', 3), async function (req, res) {
    try {
      var spender = toStr(req.body.spenderName);
      var category = toStr(req.body.category) || 'Other';
      var description = toStr(req.body.description);
      var amount = toNum(req.body.amount);
      var date = req.body.expenseDate || null;
      if (!spender) return res.status(400).json({ error: 'Pick who spent this.' });
      var memberNames = (req.trip.members || []).map(function (m) { return m.toLowerCase(); });
      if (memberNames.indexOf(spender.toLowerCase()) === -1) return res.status(400).json({ error: 'That name is not in this trip\'s member list.' });
      if (!amount || amount <= 0) return res.status(400).json({ error: 'Enter a valid amount.' });
      if (!CATEGORIES.includes(category)) category = 'Other';

      var r = await pool.query(
        `INSERT INTO trip_expenses (trip_id, spender_name, category, description, amount, expense_date, added_by_username)
         VALUES ($1,$2,$3,$4,$5,COALESCE($6,CURRENT_DATE),$7) RETURNING id`,
        [req.trip.id, spender, category, description || null, amount, date, req.user.username]
      );
      var expenseId = r.rows[0].id;

      var files = req.files || [];
      for (var i = 0; i < Math.min(files.length, 3); i++) {
        await pool.query('INSERT INTO trip_expense_photos (expense_id, photo_data, photo_mime, position) VALUES ($1,$2,$3,$4)',
          [expenseId, files[i].buffer, files[i].mimetype, i + 1]);
      }

      await pool.query('UPDATE trip_trips SET updated_at=NOW() WHERE id=$1', [req.trip.id]);
      auditLog(null, req.user.username, 'TRIP_EXPENSE_ADDED', spender + ' spent ' + amount + ' on ' + category + ' (' + req.trip.name + ')', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true, id: expenseId });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── List expenses for a trip ──
  app.get('/api/trips/:tripId/expenses', requireAuth, requireTripAccess, async function (req, res) {
    try {
      var r = await pool.query(`
        SELECT e.*, COALESCE(p.photo_count, 0) AS photo_count
        FROM trip_expenses e
        LEFT JOIN (SELECT expense_id, COUNT(*) AS photo_count FROM trip_expense_photos GROUP BY expense_id) p ON p.expense_id = e.id
        WHERE e.trip_id=$1
        ORDER BY e.expense_date DESC, e.created_at DESC
      `, [req.trip.id]);
      res.json({ expenses: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Delete an expense (and its photos, via cascade) ──
  app.delete('/api/trips/:tripId/expenses/:expenseId', requireAuth, requireTripAccess, async function (req, res) {
    try {
      var r = await pool.query('DELETE FROM trip_expenses WHERE id=$1 AND trip_id=$2 RETURNING id', [req.params.expenseId, req.trip.id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Expense not found' });
      auditLog(null, req.user.username, 'TRIP_EXPENSE_DELETED', 'Deleted expense #' + req.params.expenseId + ' (' + req.trip.name + ')', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Serve a receipt photo — same auth-via-query-token trust boundary used
  // for brand logos, since <img src> can't carry a custom header. ──
  app.get('/api/trips/:tripId/expenses/:expenseId/photos/:position', async function (req, res) {
    try {
      var token = req.headers['x-auth-token'] || req.query.token;
      if (!token) return res.status(401).end();
      var sess = await pool.query('SELECT user_id FROM sessions WHERE token=$1 AND expires_at>NOW()', [token]);
      if (!sess.rows[0]) return res.status(401).end();
      var r = await pool.query('SELECT photo_data, photo_mime FROM trip_expense_photos WHERE expense_id=$1 AND position=$2', [req.params.expenseId, req.params.position]);
      if (!r.rows[0]) return res.status(404).end();
      res.setHeader('Content-Type', r.rows[0].photo_mime || 'image/jpeg');
      res.setHeader('Cache-Control', 'private, max-age=300');
      res.send(r.rows[0].photo_data);
    } catch (e) { res.status(500).end(); }
  });

  // ── Summary: gross total, equal split, per-person balance, settlement plan ──
  app.get('/api/trips/:tripId/summary', requireAuth, requireTripAccess, async function (req, res) {
    try {
      var members = req.trip.members || [];
      var er = await pool.query('SELECT spender_name, category, amount FROM trip_expenses WHERE trip_id=$1', [req.trip.id]);
      var expenses = er.rows;

      var grossTotal = 0;
      var byPerson = {}, byCategory = {};
      members.forEach(function (m) { byPerson[m] = 0; });
      expenses.forEach(function (e) {
        var amt = toNum(e.amount);
        grossTotal += amt;
        byPerson[e.spender_name] = (byPerson[e.spender_name] || 0) + amt;
        byCategory[e.category] = (byCategory[e.category] || 0) + amt;
      });

      var n = members.length || 1;
      var equalShare = grossTotal / n;

      var people = members.map(function (m) {
        var spent = Math.round((byPerson[m] || 0) * 100) / 100;
        var balance = Math.round((spent - equalShare) * 100) / 100; // + = overspent, gets paid back; - = owes
        return { name: m, spent: spent, share: Math.round(equalShare * 100) / 100, balance: balance };
      });

      // Greedy min-transaction settlement: largest creditor paid by largest debtor, repeat.
      var creditors = people.filter(function (p) { return p.balance > 0.01; }).map(function (p) { return { name: p.name, amt: p.balance }; }).sort(function (a, b) { return b.amt - a.amt; });
      var debtors = people.filter(function (p) { return p.balance < -0.01; }).map(function (p) { return { name: p.name, amt: -p.balance }; }).sort(function (a, b) { return b.amt - a.amt; });
      var settlements = [];
      var ci = 0, di = 0;
      while (ci < creditors.length && di < debtors.length) {
        var pay = Math.min(creditors[ci].amt, debtors[di].amt);
        pay = Math.round(pay * 100) / 100;
        if (pay > 0.01) settlements.push({ from: debtors[di].name, to: creditors[ci].name, amount: pay });
        creditors[ci].amt -= pay; debtors[di].amt -= pay;
        if (creditors[ci].amt <= 0.01) ci++;
        if (debtors[di].amt <= 0.01) di++;
      }

      res.json({
        tripName: req.trip.name,
        status: req.trip.status,
        memberCount: n,
        grossTotal: Math.round(grossTotal * 100) / 100,
        equalShare: Math.round(equalShare * 100) / 100,
        byCategory: byCategory,
        people: people.sort(function (a, b) { return b.balance - a.balance; }),
        settlements: settlements,
        expenseCount: expenses.length
      });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
};
