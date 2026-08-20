// ══ TRIP EXPENSE SPLITTER ════════════════════════════════════════════════
// Friends' trip spend tracker. Create a trip, list who's joining (plain
// names — friends don't need app logins), log every spend against a
// member's name with a category (Hotel, Food, Entertainment, Car/Petrol,
// Other) and up to 3 receipt photos, then get a live gross total, an equal
// split, and — at the end of the trip — who overspent (gets money back)
// and who underspent (needs to pay in), plus a shortest "who pays whom"
// settlement list.
module.exports = function (app, pool, requireAuth, requireRole, upload, auditLog, bcrypt) {

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
      // ── Global Guest Accounts — completely independent of the office
      // `users` table / requireAuth. Azhar creates ONE (or a few) guest
      // logins here, then grants each guest account access to whichever
      // trips it should see via trip_guest_access. A guest logs in once at
      // /trip-login and sees only the trips they've been granted — they can
      // add expenses and open/close those trips, nothing else. ──
      await pool.query(`CREATE TABLE IF NOT EXISTS trip_guest_accounts (
        id SERIAL PRIMARY KEY,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        display_name TEXT NOT NULL DEFAULT 'Guest',
        active BOOLEAN NOT NULL DEFAULT true,
        created_by_username TEXT,
        created_at TIMESTAMPTZ DEFAULT NOW(),
        updated_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE TABLE IF NOT EXISTS trip_guest_access (
        id SERIAL PRIMARY KEY,
        guest_id INTEGER NOT NULL REFERENCES trip_guest_accounts(id) ON DELETE CASCADE,
        trip_id INTEGER NOT NULL REFERENCES trip_trips(id) ON DELETE CASCADE,
        granted_at TIMESTAMPTZ DEFAULT NOW(),
        UNIQUE(guest_id, trip_id)
      )`);
      await pool.query(`CREATE TABLE IF NOT EXISTS trip_guest_sessions (
        token TEXT PRIMARY KEY,
        guest_id INTEGER NOT NULL REFERENCES trip_guest_accounts(id) ON DELETE CASCADE,
        expires_at TIMESTAMPTZ NOT NULL,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_expenses_trip ON trip_expenses (trip_id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_photos_expense ON trip_expense_photos (expense_id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_guest_access_guest ON trip_guest_access (guest_id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_guest_access_trip ON trip_guest_access (trip_id)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_trip_guest_sessions_guest ON trip_guest_sessions (guest_id)`);
      console.log('Trip Expense Splitter module DB ready');
    } catch (e) { console.error('Trip module init error:', e.message); }
  })();

  var crypto = require('crypto');
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

  // ══ GUEST ACCOUNTS (admin-managed, global — not tied to one trip) ══════
  // Azhar creates guest accounts here (e.g. one shared "Friends" login used
  // across every trip). Access to individual trips is granted separately
  // below, so the same guest login can be reused trip after trip.

  app.get('/api/trip-guests', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query(`
        SELECT g.id, g.username, g.display_name, g.active, g.created_at,
               COALESCE(a.trip_count, 0) AS trip_count
        FROM trip_guest_accounts g
        LEFT JOIN (SELECT guest_id, COUNT(*) AS trip_count FROM trip_guest_access GROUP BY guest_id) a ON a.guest_id = g.id
        ORDER BY g.created_at DESC
      `);
      res.json({ guests: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // Create or reset a guest account — upsert on username. Returns the
  // plaintext password once so it can be shared; never stored/retrievable again.
  app.post('/api/trip-guests', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var username = toStr(req.body.username).toLowerCase().replace(/\s+/g, '');
      var password = toStr(req.body.password);
      var displayName = toStr(req.body.displayName) || 'Friends';
      if (!username || username.length < 3) return res.status(400).json({ error: 'Username must be at least 3 characters.' });
      if (!password || password.length < 4) return res.status(400).json({ error: 'Password must be at least 4 characters.' });

      var existing = await pool.query('SELECT id FROM trip_guest_accounts WHERE username=$1', [username]);
      var hash = await bcrypt.hash(password, 10);
      var guestId;
      if (existing.rows[0]) {
        guestId = existing.rows[0].id;
        await pool.query('UPDATE trip_guest_accounts SET password_hash=$1, display_name=$2, active=true, updated_at=NOW() WHERE id=$3', [hash, displayName, guestId]);
        await pool.query('DELETE FROM trip_guest_sessions WHERE guest_id=$1', [guestId]); // old password stops working immediately
      } else {
        var r = await pool.query(
          `INSERT INTO trip_guest_accounts (username, password_hash, display_name, created_by_username) VALUES ($1,$2,$3,$4) RETURNING id`,
          [username, hash, displayName, req.user.username]
        );
        guestId = r.rows[0].id;
      }
      auditLog(null, req.user.username, 'TRIP_GUEST_SET', 'Set guest account "' + username + '" (' + displayName + ')', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true, id: guestId, username: username, password: password, displayName: displayName, loginUrl: '/trip-login' });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.put('/api/trip-guests/:guestId/toggle', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query('UPDATE trip_guest_accounts SET active = NOT active, updated_at=NOW() WHERE id=$1 RETURNING active', [req.params.guestId]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Guest account not found' });
      if (!r.rows[0].active) await pool.query('DELETE FROM trip_guest_sessions WHERE guest_id=$1', [req.params.guestId]);
      res.json({ success: true, active: r.rows[0].active });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.delete('/api/trip-guests/:guestId', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query('DELETE FROM trip_guest_accounts WHERE id=$1 RETURNING id', [req.params.guestId]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Guest account not found' });
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Per-trip access grants — which guest accounts can see THIS trip ──
  app.get('/api/trips/:id/guest-access', requireAuth, requireTripAccess, async function (req, res) {
    try {
      var r = await pool.query(`
        SELECT g.id, g.username, g.display_name, g.active,
               (a.trip_id IS NOT NULL) AS has_access
        FROM trip_guest_accounts g
        LEFT JOIN trip_guest_access a ON a.guest_id = g.id AND a.trip_id = $1
        ORDER BY g.display_name
      `, [req.trip.id]);
      res.json({ guests: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.post('/api/trips/:id/guest-access/:guestId', requireAuth, requireTripAccess, async function (req, res) {
    try {
      var g = await pool.query('SELECT id, username FROM trip_guest_accounts WHERE id=$1', [req.params.guestId]);
      if (!g.rows[0]) return res.status(404).json({ error: 'Guest account not found' });
      await pool.query('INSERT INTO trip_guest_access (guest_id, trip_id) VALUES ($1,$2) ON CONFLICT (guest_id, trip_id) DO NOTHING', [req.params.guestId, req.trip.id]);
      auditLog(null, req.user.username, 'TRIP_GUEST_ACCESS_GRANTED', 'Granted ' + g.rows[0].username + ' access to "' + req.trip.name + '"', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  app.delete('/api/trips/:id/guest-access/:guestId', requireAuth, requireTripAccess, async function (req, res) {
    try {
      await pool.query('DELETE FROM trip_guest_access WHERE guest_id=$1 AND trip_id=$2', [req.params.guestId, req.trip.id]);
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

  // Shared summary calculator so both the admin route and the member-portal
  // route return the identical numbers.
  async function computeSummary(trip) {
    var members = trip.members || [];
    var er = await pool.query('SELECT spender_name, category, amount FROM trip_expenses WHERE trip_id=$1', [trip.id]);
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
      var balance = Math.round((spent - equalShare) * 100) / 100;
      return { name: m, spent: spent, share: Math.round(equalShare * 100) / 100, balance: balance };
    });
    var creditors = people.filter(function (p) { return p.balance > 0.01; }).map(function (p) { return { name: p.name, amt: p.balance }; }).sort(function (a, b) { return b.amt - a.amt; });
    var debtors = people.filter(function (p) { return p.balance < -0.01; }).map(function (p) { return { name: p.name, amt: -p.balance }; }).sort(function (a, b) { return b.amt - a.amt; });
    var settlements = [];
    var ci = 0, di = 0;
    while (ci < creditors.length && di < debtors.length) {
      var pay = Math.round(Math.min(creditors[ci].amt, debtors[di].amt) * 100) / 100;
      if (pay > 0.01) settlements.push({ from: debtors[di].name, to: creditors[ci].name, amount: pay });
      creditors[ci].amt -= pay; debtors[di].amt -= pay;
      if (creditors[ci].amt <= 0.01) ci++;
      if (debtors[di].amt <= 0.01) di++;
    }
    return {
      tripName: trip.name, status: trip.status, memberCount: n,
      grossTotal: Math.round(grossTotal * 100) / 100, equalShare: Math.round(equalShare * 100) / 100,
      byCategory: byCategory, people: people.sort(function (a, b) { return b.balance - a.balance; }),
      settlements: settlements, expenseCount: expenses.length
    };
  }

  // ══ GUEST PORTAL — separate login, separate URL, works across many trips ══
  // Guests never touch the office `users` table, requireAuth, or
  // azr-operations.com's main login. They authenticate with the account
  // Azhar created and get a token scoped to their guest_id; every endpoint
  // below re-checks trip_guest_access before touching a trip's data, so a
  // guest only ever sees trips they've explicitly been granted.

  var GUEST_SESSION_DAYS = 30;

  async function requireGuestAuth(req, res, next) {
    var token = req.headers['x-trip-token'] || req.query.token;
    if (!token) return res.status(401).json({ error: 'Not authenticated' });
    try {
      var r = await pool.query(
        `SELECT s.token, s.expires_at, g.id AS guest_id, g.username, g.display_name, g.active
         FROM trip_guest_sessions s JOIN trip_guest_accounts g ON g.id = s.guest_id
         WHERE s.token=$1 AND s.expires_at>NOW()`, [token]
      );
      if (!r.rows[0]) return res.status(401).json({ error: 'Session expired — please log in again.' });
      if (!r.rows[0].active) return res.status(403).json({ error: 'This login has been disabled.' });
      req.guest = r.rows[0];
      next();
    } catch (e) { res.status(500).json({ error: e.message }); }
  }

  // Verifies the logged-in guest has been granted this specific trip, then
  // attaches req.trip — used on every trip-scoped guest-portal route.
  async function requireGuestTripAccess(req, res, next) {
    try {
      var tripId = req.params.tripId;
      var access = await pool.query('SELECT 1 FROM trip_guest_access WHERE guest_id=$1 AND trip_id=$2', [req.guest.guest_id, tripId]);
      if (!access.rows[0]) return res.status(403).json({ error: 'You don\'t have access to this trip.' });
      var trip = await loadTrip(tripId);
      if (!trip) return res.status(404).json({ error: 'Trip not found' });
      req.trip = trip;
      next();
    } catch (e) { res.status(500).json({ error: e.message }); }
  }

  // ── Login ──
  app.post('/api/trip-guest-auth/login', async function (req, res) {
    try {
      var username = toStr(req.body.username).toLowerCase();
      var password = toStr(req.body.password);
      if (!username || !password) return res.status(400).json({ error: 'Username and password required' });
      var r = await pool.query('SELECT * FROM trip_guest_accounts WHERE username=$1', [username]);
      var guest = r.rows[0];
      if (!guest) return res.status(401).json({ error: 'Invalid username or password' });
      if (!guest.active) return res.status(403).json({ error: 'This login has been disabled.' });
      var match = await bcrypt.compare(password, guest.password_hash);
      if (!match) return res.status(401).json({ error: 'Invalid username or password' });

      var token = crypto.randomBytes(32).toString('hex');
      var expires = new Date(Date.now() + GUEST_SESSION_DAYS * 24 * 60 * 60 * 1000);
      await pool.query('INSERT INTO trip_guest_sessions (token, guest_id, expires_at) VALUES ($1,$2,$3)', [token, guest.id, expires]);

      res.json({ success: true, token: token, displayName: guest.display_name, username: guest.username });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Logout ──
  app.post('/api/trip-guest-auth/logout', requireGuestAuth, async function (req, res) {
    try {
      var token = req.headers['x-trip-token'] || req.query.token;
      await pool.query('DELETE FROM trip_guest_sessions WHERE token=$1', [token]);
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Who am I ──
  app.get('/api/trip-guest-auth/me', requireGuestAuth, function (req, res) {
    res.json({ displayName: req.guest.display_name, username: req.guest.username });
  });

  // ── List only the trips this guest has been granted ──
  app.get('/api/trip-guest/trips', requireGuestAuth, async function (req, res) {
    try {
      var r = await pool.query(`
        SELECT t.*, COALESCE(e.total, 0) AS gross_total, COALESCE(e.cnt, 0) AS expense_count
        FROM trip_trips t
        JOIN trip_guest_access a ON a.trip_id = t.id AND a.guest_id = $1
        LEFT JOIN (SELECT trip_id, SUM(amount) AS total, COUNT(*) AS cnt FROM trip_expenses GROUP BY trip_id) e ON e.trip_id = t.id
        ORDER BY (t.status = 'active') DESC, t.updated_at DESC
      `, [req.guest.guest_id]);
      res.json({ trips: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Trip detail, scoped to a trip the guest has access to ──
  app.get('/api/trip-guest/trips/:tripId', requireGuestAuth, requireGuestTripAccess, function (req, res) {
    res.json({ trip: req.trip });
  });

  // ── Open / close a trip (guest-allowed; no delete, no member edits) ──
  app.put('/api/trip-guest/trips/:tripId/status', requireGuestAuth, requireGuestTripAccess, async function (req, res) {
    try {
      var status = ['active', 'closed'].includes(req.body.status) ? req.body.status : null;
      if (!status) return res.status(400).json({ error: 'Invalid status' });
      await pool.query('UPDATE trip_trips SET status=$1, updated_at=NOW() WHERE id=$2', [status, req.trip.id]);
      res.json({ success: true, status: status });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Add an expense (guest portal — up to 3 receipt photos) ──
  app.post('/api/trip-guest/trips/:tripId/expenses', requireGuestAuth, requireGuestTripAccess, upload.array('photos', 3), async function (req, res) {
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
        [req.trip.id, spender, category, description || null, amount, date, 'guest:' + req.guest.username]
      );
      var expenseId = r.rows[0].id;
      var files = req.files || [];
      for (var i = 0; i < Math.min(files.length, 3); i++) {
        await pool.query('INSERT INTO trip_expense_photos (expense_id, photo_data, photo_mime, position) VALUES ($1,$2,$3,$4)',
          [expenseId, files[i].buffer, files[i].mimetype, i + 1]);
      }
      await pool.query('UPDATE trip_trips SET updated_at=NOW() WHERE id=$1', [req.trip.id]);
      res.json({ success: true, id: expenseId });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── List expenses (guest portal) ──
  app.get('/api/trip-guest/trips/:tripId/expenses', requireGuestAuth, requireGuestTripAccess, async function (req, res) {
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

  // ── Delete an expense (guest portal) ──
  app.delete('/api/trip-guest/trips/:tripId/expenses/:expenseId', requireGuestAuth, requireGuestTripAccess, async function (req, res) {
    try {
      var r = await pool.query('DELETE FROM trip_expenses WHERE id=$1 AND trip_id=$2 RETURNING id', [req.params.expenseId, req.trip.id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Expense not found' });
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Serve a receipt photo (guest portal — token via query param for <img>) ──
  app.get('/api/trip-guest/trips/:tripId/expenses/:expenseId/photos/:position', async function (req, res) {
    try {
      var token = req.headers['x-trip-token'] || req.query.token;
      if (!token) return res.status(401).end();
      var sess = await pool.query(
        `SELECT s.guest_id FROM trip_guest_sessions s WHERE s.token=$1 AND s.expires_at>NOW()`, [token]
      );
      if (!sess.rows[0]) return res.status(401).end();
      var access = await pool.query('SELECT 1 FROM trip_guest_access WHERE guest_id=$1 AND trip_id=$2', [sess.rows[0].guest_id, req.params.tripId]);
      if (!access.rows[0]) return res.status(403).end();
      var r = await pool.query('SELECT photo_data, photo_mime FROM trip_expense_photos WHERE expense_id=$1 AND position=$2', [req.params.expenseId, req.params.position]);
      if (!r.rows[0]) return res.status(404).end();
      res.setHeader('Content-Type', r.rows[0].photo_mime || 'image/jpeg');
      res.setHeader('Cache-Control', 'private, max-age=300');
      res.send(r.rows[0].photo_data);
    } catch (e) { res.status(500).end(); }
  });

  // ── Summary (guest portal) ──
  app.get('/api/trip-guest/trips/:tripId/summary', requireGuestAuth, requireGuestTripAccess, async function (req, res) {
    try { res.json(await computeSummary(req.trip)); } catch (e) { res.status(500).json({ error: e.message }); }
  });
};
