/* ═══════════════════════════════════════════════════════════════
   NPPS WCR Review — app.js
   Flow: login → home/history → paste Doc link → clean preview
         → DWR upload + Parse (coverage) → paste points in → OK
         → Grammar check (inline accept/reject) → Download revised .docx
   Original Google Doc is never modified.
   ═══════════════════════════════════════════════════════════════ */
const CFG = window.WCRR_CONFIG || { DEMO_MODE: true };

const State = {
  user: null,
  docId: null, docTitle: '', origDocxBlob: null,
  dwrs: [],            // {name, b64}
  coverage: [],        // {section, point}  (uncovered only)
  coveragePoints: [],  // {point, covered, section}  (full list)
  grammar: [],         // {id, original, suggestion, status:'open'|'accepted'|'rejected'}
  coverageDone: false,
  docLoaded: false,    // true ONLY after a document actually renders
};

/* ───────────── Toast ───────────── */
const Toast = {
  _t: null,
  show(msg, type) {
    const el = document.getElementById('toast');
    el.className = 'toast' + (type ? ' ' + type : '');
    el.innerHTML = msg;
    el.classList.remove('hidden');
    clearTimeout(Toast._t);
    Toast._t = setTimeout(() => el.classList.add('hidden'), 3200);
  }
};

/* ───────────── Screen manager ───────────── */
const Screen = {
  show(name) {
    ['auth', 'home', 'review'].forEach(s =>
      document.getElementById('screen-' + s).classList.toggle('hidden', s !== name));
    document.getElementById('topbar').classList.toggle('hidden', name === 'auth');
  }
};

/* ───────────── Auth (name + Employee ID) ───────────── */
const Auth = {
  KEY: 'wcrr_user',
  employees: [],

  async loadEmployees() {
    const err = document.getElementById('auth-err');
    try {
      const r = await fetch(CFG.EMPLOYEES_URL + '?t=' + Date.now());
      if (!r.ok) throw new Error('HTTP ' + r.status);
      Auth.employees = await r.json();
      if (!Array.isArray(Auth.employees) || !Auth.employees.length) throw new Error('empty list');
      if (err) err.textContent = '';
    } catch (e) {
      Auth.employees = [];
      if (err) err.textContent = 'Could not load the employee list (' + e.message + '). Check that employees.json is in the repo.';
    }
    Auth.populate();
  },
  populate() {
    const sel = document.getElementById('auth-name-select');
    if (!sel) return;
    sel.innerHTML = '<option value="">— Select your name —</option>';
    [...Auth.employees].sort((a,b) => a.name.localeCompare(b.name)).forEach(emp => {
      const o = document.createElement('option');
      o.value = emp.empNo; o.textContent = emp.name;
      sel.appendChild(o);
    });
  },
  onName() { document.getElementById('auth-emp').value = ''; document.getElementById('auth-err').textContent = ''; Auth._refreshBtn(); },
  onEmp() { document.getElementById('auth-err').textContent = ''; Auth._refreshBtn(); },
  _refreshBtn() {
    const sel = document.getElementById('auth-name-select').value;
    const typed = document.getElementById('auth-emp').value.trim();
    const match = sel && typed && sel.toLowerCase() === typed.toLowerCase();
    document.getElementById('auth-continue').disabled = !match;
  },

  restore() {
    try {
      const u = JSON.parse(localStorage.getItem(Auth.KEY) || 'null');
      if (u && u.name && u.emp) { State.user = u; Auth._enter(); return true; }
    } catch (e) {}
    return false;
  },
  login() {
    const sel = document.getElementById('auth-name-select');
    const empNo = sel.value;
    const typed = document.getElementById('auth-emp').value.trim();
    const err = document.getElementById('auth-err');
    if (!empNo || !typed) { err.textContent = 'Select your name and enter your employee number.'; return; }
    if (empNo.toLowerCase() !== typed.toLowerCase()) { err.textContent = 'Employee number does not match the selected name.'; return; }
    const emp = Auth.employees.find(e => e.empNo.toLowerCase() === typed.toLowerCase());
    if (!emp) { err.textContent = 'Employee not found.'; return; }
    State.user = { name: emp.name, emp: emp.empNo };
    localStorage.setItem(Auth.KEY, JSON.stringify(State.user));
    err.textContent = '';
    Auth._enter();
  },
  _enter() {
    document.getElementById('who').textContent = `${State.user.name} · ${State.user.emp}`;
    document.getElementById('demo-badge').classList.toggle('hidden', !CFG.DEMO_MODE);
    History.render();
    Screen.show('home');
  },
  logout() {
    localStorage.removeItem(Auth.KEY);
    State.user = null;
    Screen.show('auth');
  }
};

/* ───────────── History (review sessions) ───────────── */
const History = {
  key() { return 'wcrr_hist_' + (State.user ? State.user.emp : 'anon'); },
  load() { try { return JSON.parse(localStorage.getItem(History.key()) || '[]'); } catch (e) { return []; } },
  save(list) { localStorage.setItem(History.key(), JSON.stringify(list)); },
  add(entry) {
    const list = History.load();
    list.unshift(entry);
    History.save(list.slice(0, 50));
  },
  render() {
    const list = History.load();
    const wrap = document.getElementById('history-list');
    if (!list.length) {
      wrap.innerHTML = `<div class="hist-empty">No reviews yet. Start one with “Create WCR Review”.</div>`;
      return;
    }
    wrap.innerHTML = list.map(h => `
      <div class="hist-row" onclick="Review.openHistory('${h.id}')">
        <div class="hist-main">
          <span class="hist-title">${h.title || 'Untitled WCR'}</span>
          <span class="hist-meta">${new Date(h.at).toLocaleString()} · ${h.dwrCount||0} DWR(s) · ${h.fixCount||0} fix(es)</span>
        </div>
        <span class="ghost-btn">Open</span>
      </div>`).join('');
  }
};

/* ───────────── Google (read-only) — GIS token, Docs read, Drive export ─────────────
   Mirrors the current tool's "Connect Google" practice; adds the Docs read scope.
   In DEMO_MODE these are bypassed in favour of a built-in sample. */
const GoogleAPI = {
  token: null,
  _pending: null,   // shared in-flight token request (prevents double popups)
  SCOPES: 'https://www.googleapis.com/auth/documents.readonly https://www.googleapis.com/auth/drive.readonly',
  ensureToken() {
    if (CFG.DEMO_MODE) return Promise.reject(new Error('demo'));
    if (GoogleAPI.token) return Promise.resolve(GoogleAPI.token);
    // If a request is already in flight, reuse it — do NOT open a second popup.
    if (GoogleAPI._pending) return GoogleAPI._pending;
    GoogleAPI._pending = new Promise((resolve, reject) => {
      if (!window.google || !google.accounts) { reject(new Error('Google library not loaded')); return; }
      const client = google.accounts.oauth2.initTokenClient({
        client_id: CFG.GOOGLE_CLIENT_ID,
        scope: GoogleAPI.SCOPES,
        callback: (resp) => {
          if (resp && resp.access_token) { GoogleAPI.token = resp.access_token; resolve(resp.access_token); }
          else reject(new Error('Authorization failed or was cancelled.'));
        },
        error_callback: (err) => reject(new Error('Google sign-in was blocked or closed (' + (err && err.type || 'popup') + ').'))
      });
      client.requestAccessToken();
    }).finally(() => { GoogleAPI._pending = null; });
    return GoogleAPI._pending;
  },
  async readDoc(docId) {
    const token = await GoogleAPI.ensureToken();
    const r = await fetch(`https://docs.googleapis.com/v1/documents/${docId}`, {
      headers: { Authorization: 'Bearer ' + token }
    });
    if (!r.ok) throw new Error('Read failed (' + r.status + '): ' + (await GoogleAPI._reason(r)));
    return r.json();
  },
  async exportDocx(docId) {
    const token = await GoogleAPI.ensureToken();
    const url = `https://www.googleapis.com/drive/v3/files/${docId}/export?mimeType=application/vnd.openxmlformats-officedocument.wordprocessingml.document`;
    const r = await fetch(url, { headers: { Authorization: 'Bearer ' + token } });
    if (!r.ok) throw new Error('Export failed (' + r.status + '): ' + (await GoogleAPI._reason(r)));
    return r.blob();
  },
  async _reason(r) {
    try {
      const j = await r.json();
      return (j && j.error && j.error.message) ? j.error.message : 'unknown error';
    } catch (e) { return 'unknown error'; }
  }
};

/* ───────────── Worker (Gemini) — DWR coverage + grammar ───────────── */
const Worker = {
  // Maps the tool's logical modes to your worker's routes.
  ROUTES: { 'dwr-coverage': '/coverage-check', 'grammar': '/sentence-grammar' },
  async call(mode, payload) {
    if (CFG.DEMO_MODE) return Demo.worker(mode, payload);
    const route = Worker.ROUTES[mode];
    if (!route) throw new Error('Unknown AI mode: ' + mode);
    const r = await fetch(CFG.WORKER_URL + route, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!r.ok) throw new Error('AI service error (' + r.status + ')');
    return r.json();
  }
};

/* ═════════════ Review session lifecycle ═════════════ */
const Review = {
  startNew() {
    State.docId = null; State.docTitle = ''; State.origDocxBlob = null;
    State.dwrs = []; State.coverage = []; State.grammar = []; State.coverageDone = false;
    State.docLoaded = false;
    document.getElementById('doc-link').value = '';
    document.getElementById('doc-body').classList.add('hidden');
    document.getElementById('doc-body').innerHTML = '';
    document.getElementById('preview-empty').classList.remove('hidden');
    document.getElementById('doc-title').textContent = 'No document loaded';
    document.getElementById('grammar-count').classList.add('hidden');
    DWR.refresh(); DWR.clearCoverage();
    Grammar.clear();
    Review._lock('card-grammar', true);
    Review._lock('card-export', true);
    Screen.show('review');
  },
  backHome() { History.render(); Screen.show('home'); },
  openHistory(id) {
    // sessions are summaries only (we never store the customer's doc); reopen starts fresh
    Toast.show('Past reviews are summaries only — start a fresh review to load the document again.');
    Review.startNew();
  },
  _lock(cardId, locked) { document.getElementById(cardId).classList.toggle('disabled', locked); }
};

/* ═════════════ Preview ═════════════ */
const Preview = {
  extractDocId(url) {
    if (!url) return null;
    url = url.trim();
    // published links:  /document/d/e/<id>/pub
    let m = url.match(/\/document\/d\/e\/([a-zA-Z0-9_-]+)/);
    if (m) return m[1];
    // standard:  /document/d/<id>/edit?tab=...   (stops at the next slash/?/#)
    m = url.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
    if (m) return m[1];
    // open?id= or ?id=
    m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (m) return m[1];
    // bare id pasted on its own
    if (/^[a-zA-Z0-9_-]{20,}$/.test(url)) return url;
    return null;
  },
  async loadFromLink() {
    const link = document.getElementById('doc-link').value.trim();
    const btn = document.getElementById('load-doc-btn');
    const empty = document.getElementById('preview-empty');

    // reset load state — a previous hung attempt must not count as loaded
    State.docLoaded = false;
    State.docId = null;
    Review._lock('card-grammar', true);
    Review._lock('card-export', true);
    DWR.refresh();

    const LIMIT = 60; // seconds before we give up
    let remaining = LIMIT;
    const setLabel = () => { btn.innerHTML = '<span class="spin"></span>Loading… ' + remaining + 's'; };
    setLabel();
    btn.disabled = true;
    // live countdown in the empty-state area so the wait is visible
    empty.classList.remove('hidden');
    const empMsg = empty.querySelector('p');
    const empOrig = empMsg ? empMsg.innerHTML : '';
    if (empMsg) empMsg.innerHTML = 'Fetching the document from Google… <strong>' + remaining + 's</strong>';
    const ticker = setInterval(() => {
      remaining--;
      if (remaining < 0) return;
      setLabel();
      if (empMsg) empMsg.innerHTML = 'Fetching the document from Google… <strong>' + remaining + 's</strong>'
        + (remaining < 40 ? '<br><span class="muted" style="font-size:12px">Taking a while — if this is the first load, Google may be asking you to grant access in a popup.</span>' : '');
    }, 1000);

    // hard timeout so it can never spin forever
    const timeout = new Promise((_, rej) =>
      setTimeout(() => rej(new Error('Timed out after ' + LIMIT + 's. The document may be very large, not shared with your account, or Google access wasn’t granted.')), LIMIT * 1000));

    try {
      if (CFG.DEMO_MODE) {
        await Promise.race([Demo.delay(500), timeout]);
        State.docId = 'demo'; State.docTitle = 'WCR — MV Reliance Star (Demo)';
        Preview.renderModel(Demo.sampleModel());
      } else {
        const docId = Preview.extractDocId(link);
        if (!docId) throw new Error('That doesn’t look like a Google Doc link.');
        // One popup, tied to the click, before any data fetch.
        await Promise.race([GoogleAPI.ensureToken(), timeout]);
        const [doc, blob] = await Promise.race([
          Promise.all([GoogleAPI.readDoc(docId), GoogleAPI.exportDocx(docId)]),
          timeout
        ]);
        State.docId = docId;
        State.origDocxBlob = blob;
        State.docTitle = doc.title || 'WCR';
        Preview.renderModel(Preview.docToModel(doc));
      }
      // success — NOW the document is genuinely loaded
      State.docLoaded = true;
      Review._lock('card-grammar', false); // grammar available as soon as a doc is loaded
      document.getElementById('doc-title').textContent = State.docTitle;
      empty.classList.add('hidden');
      document.getElementById('doc-body').classList.remove('hidden');
      Toast.show('Document loaded. Upload DWRs to check coverage.', 'ok');
    } catch (e) {
      State.docLoaded = false;
      State.docId = null;
      if (empMsg) empMsg.innerHTML = empOrig;
      Toast.show(e.message || 'Could not load the document.', 'err');
    } finally {
      clearInterval(ticker);
      btn.disabled = false; btn.textContent = 'Load document';
      DWR.refresh(); // re-evaluate Parse gating against the real load state
    }
  },
  // Google Docs API JSON → our simple block model
  docToModel(doc) {
    const blocks = [];
    const c = (doc.body && doc.body.content) || [];
    const inlineObjs = doc.inlineObjects || {};
    // resolve an inline image element to its image URL
    const imgUrl = (objId) => {
      try {
        const eo = inlineObjs[objId].inlineObjectProperties.embeddedObject;
        return eo.imageProperties?.contentUri || (eo.embeddedDrawingProperties ? null : null);
      } catch (e) { return null; }
    };
    const runText = (el) => (el.paragraph?.elements || [])
      .map(e => e.textRun ? e.textRun.content : '').join('').replace(/\n$/, '');
    // collect any inline images inside a paragraph's elements
    const paraImages = (el) => (el.paragraph?.elements || [])
      .filter(e => e.inlineObjectElement)
      .map(e => imgUrl(e.inlineObjectElement.inlineObjectId))
      .filter(Boolean);

    c.forEach(el => {
      if (el.paragraph) {
        const imgs = paraImages(el);
        imgs.forEach(src => blocks.push({ t: 'img', src }));
        const style = el.paragraph.paragraphStyle?.namedStyleType || 'NORMAL_TEXT';
        const text = runText(el);
        if (!text.trim()) return;
        if (style === 'TITLE') blocks.push({ t: 'h1', text });
        else if (style === 'HEADING_1' || style === 'HEADING_2') blocks.push({ t: 'h2', text });
        else if (style === 'HEADING_3') blocks.push({ t: 'h3', text });
        else if (el.paragraph.bullet) blocks.push({ t: 'li', text });
        else blocks.push({ t: 'p', text });
      } else if (el.table) {
        const rows = (el.table.tableRows || []).map(r =>
          (r.tableCells || []).map(cell => {
            // pull text AND any images from the cell
            let cellText = '';
            const cellImgs = [];
            (cell.content || []).forEach(cc => {
              if (cc.paragraph) {
                cellText += (cc.paragraph.elements || []).map(e => e.textRun ? e.textRun.content : '').join('');
                (cc.paragraph.elements || []).forEach(e => {
                  if (e.inlineObjectElement) {
                    const u = imgUrl(e.inlineObjectElement.inlineObjectId);
                    if (u) cellImgs.push(u);
                  }
                });
              }
            });
            return { text: cellText.trim(), imgs: cellImgs };
          }));
        blocks.push({ t: 'table', rows });
      }
    });
    return blocks;
  },
  renderModel(blocks) {
    const body = document.getElementById('doc-body');
    let html = '', openList = false;
    const closeList = () => { if (openList) { html += '</ul>'; openList = false; } };
    blocks.forEach(b => {
      if (b.t === 'li') { if (!openList) { html += '<ul>'; openList = true; } html += `<li>${esc(b.text)}</li>`; return; }
      closeList();
      if (b.t === 'h1') html += `<h1>${esc(b.text)}</h1>`;
      else if (b.t === 'h2') html += `<h2>${esc(b.text)}</h2>`;
      else if (b.t === 'h3') html += `<h3>${esc(b.text)}</h3>`;
      else if (b.t === 'p') html += `<p>${esc(b.text)}</p>`;
      else if (b.t === 'img') html += `<img class="doc-img" src="${b.src}" loading="lazy" alt="figure"/>`;
      else if (b.t === 'table') {
        const multiCol = b.rows.length > 1 && (b.rows[0] || []).length > 1;
        html += '<table>' + b.rows.map((r, ri) =>
          '<tr>' + r.map(cell => {
            const txt = (cell && cell.text) || '';
            const imgs = (cell && cell.imgs) || [];
            const inner = esc(txt) + imgs.map(s => `<img class="doc-img" src="${s}" loading="lazy" alt="figure"/>`).join('');
            return (multiCol && ri === 0) ? `<th>${inner}</th>` : `<td>${inner}</td>`;
          }).join('') + '</tr>'
        ).join('') + '</table>';
      }
    });
    closeList();
    body.innerHTML = html;
    body.setAttribute('contenteditable', 'true');
    body.setAttribute('spellcheck', 'false');
  }
};

/* ═════════════ DWR upload + Parse (coverage) ═════════════ */
const DWR = {
  pick() {
    const inp = document.createElement('input');
    inp.type = 'file'; inp.accept = 'application/pdf'; inp.multiple = true;
    inp.onchange = e => Array.from(e.target.files).forEach(DWR.add);
    inp.click();
  },
  add(file) {
    if (file.type !== 'application/pdf') { Toast.show('DWRs must be PDF.', 'err'); return; }
    const r = new FileReader();
    r.onload = ev => {
      State.dwrs.push({ name: file.name, b64: String(ev.target.result).split(',')[1] });
      DWR.refresh();
    };
    r.readAsDataURL(file);
  },
  remove(i) { State.dwrs.splice(i, 1); DWR.refresh(); },
  refresh() {
    const list = document.getElementById('dwr-list');
    list.innerHTML = State.dwrs.map((d, i) =>
      `<div class="dwr-item"><span class="doc-ico">📕</span><span class="nm">${esc(d.name)}</span>
        <button class="rm" onclick="DWR.remove(${i})" title="Remove">×</button></div>`).join('');
    document.getElementById('dwr-parse-btn').disabled = State.dwrs.length === 0 || !State.docLoaded;
  },
  async parse() {
    if (!State.docLoaded) { Toast.show('Load the document first — coverage needs the report text to compare against.', 'err'); return; }
    const reportText = DWR._currentReportText().trim();
    if (reportText.length < 50) {
      Toast.show('The loaded report looks empty, so coverage can’t be checked. Re-load the document.', 'err');
      return;
    }
    const btn = document.getElementById('dwr-parse-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span>Parsing…';
    try {
      const res = await Worker.call('dwr-coverage', {
        report: reportText,
        dwrs: State.dwrs.map(d => ({ name: d.name, b64: d.b64 }))
      });
      State.coveragePoints = (res && res.points) || [];
      State.coverage = (res && res.uncovered) || [];
      DWR.renderCoverage();
      const missing = State.coverage.length;
      if (!State.coveragePoints.length) Toast.show('Couldn’t extract points from those DWRs — check they’re the right PDFs.', 'err');
      else if (!missing) Toast.show('All ' + State.coveragePoints.length + ' DWR points appear covered.', 'ok');
      else Toast.show(missing + ' of ' + State.coveragePoints.length + ' DWR points are NOT covered — see the list.', 'err');
    } catch (e) {
      Toast.show(e.message || 'Parse failed.', 'err');
    } finally {
      btn.disabled = false; btn.textContent = 'Parse';
    }
  },
  _currentReportText() {
    return document.getElementById('doc-body').innerText || '';
  },
  renderCoverage() {
    const out = document.getElementById('coverage-out');
    const items = document.getElementById('coverage-items');
    const pts = State.coveragePoints;
    if (!pts.length) { out.classList.add('hidden'); return; }
    out.classList.remove('hidden');
    const covered = pts.filter(p => p.covered).length;
    const missing = pts.length - covered;
    let html = `<div class="cov-summary"><strong>${covered}</strong> covered · <strong>${missing}</strong> not covered · ${pts.length} total</div>`;
    // show NOT covered first (actionable), each with copy + the section to paste into
    pts.filter(p => !p.covered).forEach((p) => {
      const idx = State.coverage.findIndex(c => c.point === p.point);
      html += `<div class="cov-item">
        ${idx >= 0 ? `<button class="copy" onclick="DWR.copyPoint(${idx})">copy</button>` : ''}
        <span class="cov-flag miss">NOT COVERED</span>
        <span class="cov-sec">${esc(p.section || 'General')}</span>${esc(p.point)}
      </div>`;
    });
    // then covered (collapsed-feel, muted)
    pts.filter(p => p.covered).forEach((p) => {
      html += `<div class="cov-item ok">
        <span class="cov-flag good">COVERED</span>${esc(p.point)}
      </div>`;
    });
    items.innerHTML = html;
  },
  copyPoint(i) {
    const c = State.coverage[i];
    const text = c.point;
    navigator.clipboard?.writeText(text).then(
      () => Toast.show('Copied — paste it into the “' + (c.section || 'relevant') + '” section.', 'ok'),
      () => Toast.show('Copy this: ' + text)
    );
  },
  acceptCoverage() {
    State.coverageDone = true;
    Review._lock('card-grammar', false);
    Toast.show('Coverage confirmed. Run the grammar check next.', 'ok');
    document.getElementById('coverage-out').classList.add('hidden');
  },
  clearCoverage() {
    State.coverage = [];
    State.coveragePoints = [];
    document.getElementById('coverage-out').classList.add('hidden');
    document.getElementById('coverage-items').innerHTML = '';
  }
};

/* ═════════════ Grammar (inline highlight, accept/reject) ═════════════ */
const Grammar = {
  async run() {
    const btn = document.getElementById('grammar-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span>Checking…';
    try {
      const text = document.getElementById('doc-body').innerText || '';
      const res = await Worker.call('grammar', { report: text });
      const issues = (res && res.issues) || [];
      // keep only issues whose original sentence is actually present in the preview
      State.grammar = issues
        .filter(it => it.original && document.getElementById('doc-body').innerText.includes(it.original))
        .map((it, i) => ({ id: 'g' + i, original: it.original, suggestion: it.suggestion, status: 'open' }));
      Grammar.wrapAll();
      Grammar.render();
      const n = State.grammar.length;
      const chip = document.getElementById('grammar-count');
      chip.textContent = n + ' sentence' + (n === 1 ? '' : 's') + ' flagged';
      chip.classList.toggle('hidden', n === 0);
      Review._lock('card-export', false);
      Toast.show(n ? n + ' sentence(s) flagged. Click each to review.' : 'No problem sentences found.', n ? '' : 'ok');
    } catch (e) {
      Toast.show(e.message || 'Grammar check failed.', 'err');
    } finally {
      btn.disabled = false; btn.textContent = 'Run grammar check';
    }
  },
  // wrap each flagged sentence in the (editable) preview with a clickable span
  wrapAll() {
    const body = document.getElementById('doc-body');
    State.grammar.forEach(g => {
      const node = Grammar._findTextNode(body, g.original);
      if (!node) return;
      const idx = node.nodeValue.indexOf(g.original);
      if (idx < 0) return;
      const range = document.createRange();
      range.setStart(node, idx);
      range.setEnd(node, idx + g.original.length);
      const span = document.createElement('span');
      span.className = 'gx'; span.dataset.gid = g.id;
      span.onclick = () => Grammar.locate(g.id);
      try { range.surroundContents(span); } catch (e) {}
    });
  },
  _findTextNode(root, needle) {
    const walk = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
    let n;
    while ((n = walk.nextNode())) { if (n.nodeValue.includes(needle)) return n; }
    return null;
  },
  render() {
    const wrap = document.getElementById('grammar-list');
    if (!State.grammar.length) { wrap.innerHTML = '<p class="muted" style="font-size:12.5px;margin-top:8px">No sentence-level issues.</p>'; return; }
    wrap.innerHTML = State.grammar.map(g => `
      <div class="gitem ${g.status !== 'open' ? 'done' : ''}" id="item-${g.id}" onclick="Grammar.locate('${g.id}')">
        <span class="g-was">${esc(g.original)}</span>
        <span class="g-now">${esc(g.suggestion)}</span>
        ${g.status === 'open' ? `<div class="g-btns">
          <button class="g-accept" onclick="event.stopPropagation();Grammar.accept('${g.id}')">Accept</button>
          <button class="g-reject" onclick="event.stopPropagation();Grammar.reject('${g.id}')">Reject</button>
        </div>` : `<span class="muted" style="font-size:11.5px">${g.status === 'accepted' ? '✓ accepted' : 'rejected'}</span>`}
      </div>`).join('');
  },
  locate(gid) {
    document.querySelectorAll('.gx.active').forEach(e => e.classList.remove('active'));
    document.querySelectorAll('.gitem.active').forEach(e => e.classList.remove('active'));
    const span = document.querySelector(`.gx[data-gid="${gid}"]`);
    if (span) { span.classList.add('active'); span.scrollIntoView({ behavior: 'smooth', block: 'center' }); }
    document.getElementById('item-' + gid)?.classList.add('active');
  },
  accept(gid) {
    const g = State.grammar.find(x => x.id === gid); if (!g) return;
    const span = document.querySelector(`.gx[data-gid="${gid}"]`);
    if (span) { span.textContent = g.suggestion; span.classList.remove('active'); span.classList.add('fixed'); }
    g.status = 'accepted';
    Grammar.render();
  },
  reject(gid) {
    const g = State.grammar.find(x => x.id === gid); if (!g) return;
    const span = document.querySelector(`.gx[data-gid="${gid}"]`);
    if (span) { const t = document.createTextNode(span.textContent); span.replaceWith(t); }
    g.status = 'rejected';
    Grammar.render();
  },
  clear() {
    State.grammar = [];
    document.getElementById('grammar-list').innerHTML = '';
    document.getElementById('grammar-count').classList.add('hidden');
  }
};

/* ═════════════ Export — revised .docx ═════════════
   LIVE: take the original exported .docx and apply accepted text
   replacements (layout preserved). DEMO / no-original: build a clean
   .docx from the current preview. Either way the original Doc is safe. */
const Export = {
  async ensureJSZip() {
    if (window.JSZip) return;
    await new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js';
      s.onload = res; s.onerror = () => rej(new Error('Could not load export library'));
      document.head.appendChild(s);
    });
  },
  async run() {
    const btn = document.getElementById('export-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span>Building…';
    try {
      await Export.ensureJSZip();
      const accepted = State.grammar.filter(g => g.status === 'accepted');
      let blob;
      if (State.origDocxBlob && !CFG.DEMO_MODE) {
        blob = await Export.patchOriginal(State.origDocxBlob, accepted);
      } else {
        blob = await Export.buildFromPreview();
      }
      const fname = (State.docTitle || 'WCR').replace(/[^\w\- ]+/g, '').replace(/\s+/g, '_') + '_revised.docx';
      Export._download(blob, fname);
      History.add({
        id: 'r' + Date.now(), title: State.docTitle, at: Date.now(),
        dwrCount: State.dwrs.length, fixCount: accepted.length
      });
      Toast.show('Revised .docx downloaded. Original Doc untouched.', 'ok');
    } catch (e) {
      Toast.show(e.message || 'Export failed.', 'err');
    } finally {
      btn.disabled = false; btn.textContent = 'Download revised .docx';
    }
  },
  // surgical text replacement inside the original docx (word/document.xml)
  async patchOriginal(blob, accepted) {
    const zip = await JSZip.loadAsync(blob);
    let xml = await zip.file('word/document.xml').async('string');
    accepted.forEach(g => {
      // replace across the plain-text projection; works when a sentence
      // sits within a single run (the common case for body prose).
      const from = Export._xmlEscape(g.original);
      const to = Export._xmlEscape(g.suggestion);
      xml = xml.split(from).join(to);
    });
    zip.file('word/document.xml', xml);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  },
  // build a clean docx from the current preview DOM
  async buildFromPreview() {
    const body = document.getElementById('doc-body');
    const paras = [];
    body.childNodes.forEach(node => Export._nodeToParas(node, paras));
    const docXml = Export._docXml(paras);
    const zip = new JSZip();
    zip.file('[Content_Types].xml', Export.CT);
    zip.folder('_rels').file('.rels', Export.RELS);
    zip.folder('word').file('document.xml', docXml);
    zip.folder('word').folder('_rels').file('document.xml.rels', Export.DOCRELS);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  },
  _nodeToParas(node, out) {
    if (node.nodeType === 3) { if (node.nodeValue.trim()) out.push({ style: 'Normal', text: node.nodeValue }); return; }
    if (node.nodeType !== 1) return;
    const tag = node.tagName.toLowerCase();
    if (tag === 'h1') out.push({ style: 'Title', text: node.innerText });
    else if (tag === 'h2') out.push({ style: 'Heading1', text: node.innerText });
    else if (tag === 'h3') out.push({ style: 'Heading2', text: node.innerText });
    else if (tag === 'p') out.push({ style: 'Normal', text: node.innerText });
    else if (tag === 'ul' || tag === 'ol') node.querySelectorAll('li').forEach(li => out.push({ style: 'Normal', text: '• ' + li.innerText }));
    else if (tag === 'table') {
      node.querySelectorAll('tr').forEach(tr => {
        const cells = Array.from(tr.children).map(td => td.innerText.trim());
        out.push({ style: 'Normal', text: cells.join('   |   ') });
      });
    } else if (node.innerText && node.innerText.trim()) out.push({ style: 'Normal', text: node.innerText });
  },
  _docXml(paras) {
    const body = paras.map(p => {
      const st = p.style === 'Normal' ? '' : `<w:pPr><w:pStyle w:val="${p.style}"/></w:pPr>`;
      return `<w:p>${st}<w:r><w:t xml:space="preserve">${Export._xmlEscape(p.text)}</w:t></w:r></w:p>`;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134"/></w:sectPr></w:body></w:document>`;
  },
  _xmlEscape(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); },
  _download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(a.href); }, 2000);
  },
  CT: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  RELS: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  DOCRELS: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`
};

/* ───────────── util ───────────── */
function esc(s) { return (s == null ? '' : String(s)).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

/* ───────────── Demo data ───────────── */
const Demo = {
  delay(ms) { return new Promise(r => setTimeout(r, ms)); },
  sampleModel() {
    return [
      { t: 'h1', text: 'Work Completion Report' },
      { t: 'p', text: 'Customer: Reliance Shipping  ·  Vessel: MV Reliance Star  ·  Engine: Niigata 6L28HX' },
      { t: 'h2', text: 'History' },
      { t: 'p', text: 'The engine was running since last 12000 hours without any major overhaul being carried out on it.' },
      { t: 'h2', text: 'Scope of Work' },
      { t: 'p', text: 'Complete top overhaul of all six units including decarbonisation of cylinder heads, calibration of liners, and inspection of the running gear.' },
      { t: 'li', text: 'Cylinder heads overhauled and pressure tested.' },
      { t: 'li', text: 'Fuel injectors tested and reconditioned.' },
      { t: 'h2', text: 'Maintenance Summary' },
      { t: 'p', text: 'All the units was opened up and the components were measured for wear as per the maker manual.' },
      { t: 'p', text: 'The crankshaft deflections were taken and found within the limit specified by manufacturer.' },
      { t: 'h2', text: 'Recommendations' },
      { t: 'p', text: 'The engine post overhaul must be closely monitored for any abnormalities which could cause serious breakdowns.' },
    ];
  },
  worker(mode, payload) {
    if (mode === 'dwr-coverage') {
      return Demo.delay(700).then(() => ({
        uncovered: [
          { section: 'Maintenance Summary', point: 'The fuel injection pump was calibrated on the test bench, but this is not mentioned anywhere in the report.' },
          { section: 'Scope of Work', point: 'Turbocharger cartridge replacement appears in the DWR dated 12 Mar but is missing from the scope.' },
          { section: 'History', point: 'DWR notes a prior crankcase relief valve incident — worth recording under engine history.' },
        ]
      }));
    }
    if (mode === 'grammar') {
      return Demo.delay(700).then(() => ({
        issues: [
          { original: 'The engine was running since last 12000 hours without any major overhaul being carried out on it.',
            suggestion: 'The engine had run for 12,000 hours since its last major overhaul.' },
          { original: 'All the units was opened up and the components were measured for wear as per the maker manual.',
            suggestion: 'All units were opened up and their components measured for wear as per the maker’s manual.' },
        ]
      }));
    }
    return Promise.resolve({});
  }
};

/* ───────────── boot ───────────── */
(function init() {
  document.getElementById('demo-badge').classList.toggle('hidden', !CFG.DEMO_MODE);
  Auth.loadEmployees();
  if (!Auth.restore()) Screen.show('auth');
  // load Google Identity Services if configured
  if (!CFG.DEMO_MODE && CFG.GOOGLE_CLIENT_ID) {
    const s = document.createElement('script');
    s.src = 'https://accounts.google.com/gsi/client'; s.async = true; s.defer = true;
    document.head.appendChild(s);
  }
})();
