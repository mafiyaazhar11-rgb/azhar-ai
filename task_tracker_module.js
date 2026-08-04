// ══ TASK TRACKER ═════════════════════════════════════════════════════════
// Super admin / sub admin assign a task to a specific user by username —
// a title, a comment/instructions, and a deadline. That's it, no file
// upload on either side. The assigned user's "submission" is just a text
// message describing what they did, written when they mark it done.
module.exports = function (app, pool, requireAuth, requireRole, auditLog) {

  (async function initTable() {
    try {
      await pool.query(`CREATE TABLE IF NOT EXISTS tasks (
        id SERIAL PRIMARY KEY,
        title TEXT NOT NULL,
        comments TEXT,
        assigned_to_username TEXT NOT NULL,
        assigned_to_full_name TEXT,
        assigned_by_username TEXT NOT NULL,
        assigned_by_full_name TEXT,
        deadline DATE,
        status TEXT NOT NULL DEFAULT 'pending',
        submission_text TEXT,
        submitted_at TIMESTAMPTZ,
        created_at TIMESTAMPTZ DEFAULT NOW()
      )`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_tasks_assignee ON tasks (assigned_to_username)`);
      await pool.query(`CREATE INDEX IF NOT EXISTS idx_tasks_status ON tasks (status)`);
      console.log('Task Tracker module DB ready');
    } catch (e) { console.error('Task Tracker init error:', e.message); }
  })();

  function toStr(v) { return String(v == null ? '' : v).trim(); }

  // ── Lightweight assignee picker — superadmin/subadmin only, no passwords
  // or other sensitive fields, just who's available to assign a task to. ──
  app.get('/api/tasks/assignable-users', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var r = await pool.query('SELECT username, full_name, role FROM users WHERE active = true ORDER BY full_name, username');
      res.json({ users: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Create / assign a new task ──
  app.post('/api/tasks', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var title = toStr(req.body.title);
      var comments = toStr(req.body.comments);
      var assignedTo = toStr(req.body.assignedTo);
      var deadline = req.body.deadline || null;
      if (!title) return res.status(400).json({ error: 'Task name is required.' });
      if (!assignedTo) return res.status(400).json({ error: 'Please pick who this task is assigned to.' });

      var u = await pool.query('SELECT username, full_name FROM users WHERE username=$1', [assignedTo]);
      if (!u.rows[0]) return res.status(400).json({ error: 'That user was not found.' });

      var r = await pool.query(
        `INSERT INTO tasks (title, comments, assigned_to_username, assigned_to_full_name, assigned_by_username, assigned_by_full_name, deadline)
         VALUES ($1,$2,$3,$4,$5,$6,$7) RETURNING id`,
        [title, comments || null, u.rows[0].username, u.rows[0].full_name || u.rows[0].username,
         req.user.username, req.user.full_name || req.user.username, deadline]
      );
      auditLog(null, req.user.username, 'TASK_ASSIGNED', 'Assigned "' + title + '" to ' + (u.rows[0].full_name || u.rows[0].username) + (deadline ? ' — due ' + deadline : ''), req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true, id: r.rows[0].id });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Super admin / sub admin: full list, optional filters ──
  app.get('/api/tasks', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var params = [], where = [];
      if (req.query.status) { where.push('status=$' + (params.length + 1)); params.push(req.query.status); }
      if (req.query.username) { where.push('assigned_to_username=$' + (params.length + 1)); params.push(req.query.username); }
      var q = 'SELECT * FROM tasks' + (where.length ? ' WHERE ' + where.join(' AND ') : '') + ' ORDER BY (status=\'pending\') DESC, deadline ASC NULLS LAST, created_at DESC';
      var r = await pool.query(q, params);
      res.json({ tasks: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Any logged-in user: just their own tasks ──
  app.get('/api/tasks/mine', requireAuth, async function (req, res) {
    try {
      var r = await pool.query(
        `SELECT * FROM tasks WHERE assigned_to_username=$1 ORDER BY (status='pending') DESC, deadline ASC NULLS LAST, created_at DESC`,
        [req.user.username]
      );
      res.json({ tasks: r.rows });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── The assigned user (or an admin) submits their completion text ──
  app.put('/api/tasks/:id/submit', requireAuth, async function (req, res) {
    try {
      var id = parseInt(req.params.id);
      var submissionText = toStr(req.body.submissionText);
      if (!submissionText) return res.status(400).json({ error: 'Please write a short message describing what you did.' });

      var t = await pool.query('SELECT * FROM tasks WHERE id=$1', [id]);
      if (!t.rows[0]) return res.status(404).json({ error: 'Task not found.' });
      var isOwner = t.rows[0].assigned_to_username === req.user.username;
      var isAdmin = req.user.role === 'superadmin' || req.user.role === 'subadmin';
      if (!isOwner && !isAdmin) return res.status(403).json({ error: 'This task isn\'t assigned to you.' });

      await pool.query(
        `UPDATE tasks SET status='submitted', submission_text=$1, submitted_at=NOW() WHERE id=$2`,
        [submissionText, id]
      );
      auditLog(null, req.user.username, 'TASK_SUBMITTED', 'Submitted "' + t.rows[0].title + '"', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Admin: reopen a submitted task (send it back if the work wasn't actually done) ──
  app.put('/api/tasks/:id/reopen', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var id = parseInt(req.params.id);
      var r = await pool.query(`UPDATE tasks SET status='pending' WHERE id=$1 RETURNING title`, [id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Task not found.' });
      auditLog(null, req.user.username, 'TASK_REOPENED', 'Reopened "' + r.rows[0].title + '"', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Admin: edit a task's title/comments/deadline ──
  app.put('/api/tasks/:id', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var id = parseInt(req.params.id);
      var title = toStr(req.body.title);
      var comments = toStr(req.body.comments);
      var deadline = req.body.deadline || null;
      if (!title) return res.status(400).json({ error: 'Task name is required.' });
      var r = await pool.query(
        `UPDATE tasks SET title=$1, comments=$2, deadline=$3 WHERE id=$4 RETURNING id`,
        [title, comments || null, deadline, id]
      );
      if (!r.rows[0]) return res.status(404).json({ error: 'Task not found.' });
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });

  // ── Admin: delete a task ──
  app.delete('/api/tasks/:id', requireAuth, requireRole('superadmin', 'subadmin'), async function (req, res) {
    try {
      var id = parseInt(req.params.id);
      var r = await pool.query('DELETE FROM tasks WHERE id=$1 RETURNING title, assigned_to_username', [id]);
      if (!r.rows[0]) return res.status(404).json({ error: 'Task not found.' });
      auditLog(null, req.user.username, 'TASK_DELETED', 'Deleted "' + r.rows[0].title + '" (was assigned to ' + r.rows[0].assigned_to_username + ')', req.headers['x-forwarded-for'] || req.ip || '');
      res.json({ success: true });
    } catch (e) { res.status(500).json({ error: e.message }); }
  });
};
