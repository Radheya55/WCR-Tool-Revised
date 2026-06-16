/* ═══════════════════════════════════════════════════════════════
   NPPS WCR Review — app.js
   Flow: login → home/history → paste Doc link → clean preview
         → DWR upload + Parse (coverage) → paste points in → OK
         → Grammar check (inline accept/reject) → Download revised .docx
   Original Google Doc is never modified.
   ═══════════════════════════════════════════════════════════════ */
const CFG = window.WCRR_CONFIG || { DEMO_MODE: true };
const BUILD = 'v15';

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
    ['auth', 'home', 'review', 'confirm'].forEach(s =>
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
  SCOPES: 'https://www.googleapis.com/auth/documents https://www.googleapis.com/auth/drive',
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
  // Make a copy of the Doc in the user's Drive. Returns the new file id.
  async copyDoc(docId, newName) {
    const token = await GoogleAPI.ensureToken();
    const r = await fetch(`https://www.googleapis.com/drive/v3/files/${docId}/copy`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ name: newName })
    });
    if (!r.ok) throw new Error('Copy failed (' + r.status + '): ' + (await GoogleAPI._reason(r)));
    const j = await r.json();
    return j.id;
  },
  // Apply text replacements to a Doc via batchUpdate. reps = [{from, to}].
  // Each becomes a replaceAllText request (matchCase true, exact text).
  async applyReplacements(docId, reps) {
    const token = await GoogleAPI.ensureToken();
    const requests = reps.map(rep => ({
      replaceAllText: {
        containsText: { text: rep.from, matchCase: true },
        replaceText: rep.to
      }
    }));
    const r = await fetch(`https://docs.googleapis.com/v1/documents/${docId}:batchUpdate`, {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ requests })
    });
    if (!r.ok) throw new Error('Apply failed (' + r.status + '): ' + (await GoogleAPI._reason(r)));
    return r.json(); // contains replies with occurrencesChanged counts
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
  _lock(cardId, locked) {
    const card = document.getElementById(cardId);
    card.classList.toggle('disabled', locked);
    // also flip the actual button's disabled attribute, not just the card class
    const btn = card.querySelector('button');
    if (btn) btn.disabled = locked;
  },
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
        // Read ONLY the document text for the preview. The .docx export is
        // heavy (Google refuses it on large image-rich docs: "too heavy to
        // import"), so we defer it until the user actually clicks Download.
        const doc = await Promise.race([GoogleAPI.readDoc(docId), timeout]);
        State.docId = docId;
        State.origDocxBlob = null;   // fetched lazily at export time
        State.docTitle = doc.title || 'WCR';
        Preview.renderModel(Preview.docToModel(doc));
      }
      // success — NOW the document is genuinely loaded
      State.docLoaded = true;
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
  // Google Docs API JSON → structured block model that reads like the original:
  // section titles, two-column label/value tables, sub-headed bullet lists,
  // and a generated table of contents. Number-grid tables are skipped.
  docToModel(doc) {
    const blocks = [];
    const c = (doc.body && doc.body.content) || [];
    const runText = (el) => (el.paragraph?.elements || [])
      .map(e => e.textRun ? e.textRun.content : '').join('').replace(/\n$/, '');
    const cellText = (cell) => (cell.content || [])
      .map(cc => cc.paragraph ? (cc.paragraph.elements || [])
        .map(e => e.textRun ? e.textRun.content : '').join('') : '')
      .join(' ').replace(/\s+/g, ' ').trim();
    const isBoldOnly = (el) => {
      const els = (el.paragraph?.elements || []).filter(e => e.textRun && e.textRun.content.trim());
      return els.length > 0 && els.every(e => e.textRun.textStyle && e.textRun.textStyle.bold);
    };

    const sections = [];
    let skipNextTable = false;

    c.forEach(el => {
      if (el.paragraph) {
        const style = el.paragraph.paragraphStyle?.namedStyleType || 'NORMAL_TEXT';
        const text = runText(el).trim();
        if (!text) return;
        const isTitle = style === 'TITLE';
        const isHeading = style === 'HEADING_1' || style === 'HEADING_2';
        // a short, fully-bold normal paragraph is a section title in these docs
        const isSectionByBold = style === 'NORMAL_TEXT' && isBoldOnly(el) && text.length < 70 && !el.paragraph.bullet;

        if (isTitle) { blocks.push({ t: 'h1', text }); return; }
        if (isHeading || isSectionByBold) {
          if (/^contents$/i.test(text)) { skipNextTable = true; return; } // we build our own TOC
          sections.push(text);
          blocks.push({ t: 'section', text });
          return;
        }
        if (style === 'HEADING_3' || style === 'HEADING_4') blocks.push({ t: 'h3', text });
        else if (el.paragraph.bullet) blocks.push({ t: 'li', text });
        else blocks.push({ t: 'p', text });
        return;
      }

      if (el.table) {
        const rows = (el.table.tableRows || []).map(r =>
          (r.tableCells || []).map(cell => cellText(cell)));
        if (skipNextTable) { skipNextTable = false; return; }   // the CONTENTS table
        const flat = rows.flat().filter(Boolean);
        if (!flat.length) return;
        const wordy = flat.filter(t => Preview._isProse(t)).length;
        if ((flat.length - wordy) > wordy) return;              // number grid → skip

        // Shape detection: is this a 2-column label/value table?
        const twoColRows = rows.filter(r => r.filter(x => x && x.trim()).length === 2);
        const isKV = twoColRows.length >= Math.max(2, rows.length * 0.5);

        if (isKV) {
          const kv = [];
          rows.forEach(r => {
            const ne = r.filter(x => x && x.trim());
            if (ne.length === 2) kv.push([ne[0].trim(), ne[1].trim()]);
            else if (ne.length === 1 && Preview._isProse(ne[0])) kv.push([ne[0].trim(), '']);
          });
          if (kv.length) blocks.push({ t: 'kvtable', rows: kv });
          return;
        }

        // Otherwise: prose table (Maintenance Summary, Scope) → subheads+bullets
        rows.forEach(cells => {
          cells.filter(t => t && t.trim()).forEach(t => {
            if (!Preview._isProse(t)) return;
            if (/^(ok|nil|none|-|\u2013|\u2014)$/i.test(t.trim())) return;
            const parts = Preview._splitSentences(t);
            if (parts.length > 1) Preview._emitWithSubheads(parts, blocks);
            else blocks.push({ t: 'p', text: t });
          });
        });
      }
    });

    // insert a generated Table of Contents after the title
    if (sections.length >= 3) {
      const toc = { t: 'toc', items: sections };
      const at = (blocks[0] && blocks[0].t === 'h1') ? 1 : 0;
      blocks.splice(at, 0, toc);
    }
    return blocks;
  },
  _isProse(t) { return !!t && /[A-Za-z]{2,}/.test(t); },

  // sentence split that does NOT break on abbreviations (Mr. Josko stays whole)
  _splitSentences(t) {
    t = (t || '').replace(/\s+/g, ' ').trim();
    const ABBR = /\b(Mr|Mrs|Ms|Dr|Sr|Jr|St|No|Nos|Fig|Ref|Sl|vs|etc|approx|Rev|Sec|Sr\.No|DR|Dia|min|max|temp|Qty)\.\s/gi;
    t = t.replace(ABBR, m => m.replace('. ', '.\u0001'));      // protect the space
    let parts = t.split(/(?<=[.!?])\s+(?=[A-Z0-9])/);
    return parts.map(s => s.replace(/\u0001/g, ' ').trim()).filter(s => s.length > 1);
  },

  // turn a flat sentence list into {subhead}+{bullets}. A short fragment that
  // ends with ':' (e.g. "Cylinder Heads:") becomes a bold sub-heading; the
  // sentences after it become bullets under it.
  _emitWithSubheads(parts, blocks) {
    let bucket = [];
    const flush = () => { if (bucket.length) { blocks.push({ t: 'bullets', items: bucket }); bucket = []; } };
    parts.forEach(s => {
      const heady = /:$/.test(s) && s.length < 60;          // "Cylinder Heads:"
      // also catch "Word Word:" prefix inside a longer sentence
      const m = !heady && s.match(/^([A-Z][A-Za-z ,/&-]{2,40}):\s+(.*)$/);
      if (heady) { flush(); blocks.push({ t: 'subhead', text: s.replace(/:$/, '') }); }
      else if (m) { flush(); blocks.push({ t: 'subhead', text: m[1] }); bucket.push(m[2]); }
      else bucket.push(s);
    });
    flush();
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
      else if (b.t === 'section') html += `<h2 class="sec-title">${esc(b.text)}</h2>`;
      else if (b.t === 'toc') html += '<div class="toc"><div class="toc-h">Contents</div><ol>' +
        b.items.map(i => `<li>${esc(i)}</li>`).join('') + '</ol></div>';
      else if (b.t === 'subhead') html += `<p class="sub-head">${esc(b.text)}</p>`;
      else if (b.t === 'bullets') html += '<ul class="pt-list">' + b.items.map(i => `<li>${esc(i)}</li>`).join('') + '</ul>';
      else if (b.t === 'kvtable') html += '<table class="kv">' + b.rows.map(r =>
        `<tr><td class="kv-k">${esc(r[0])}</td><td class="kv-v">${esc(r[1])}</td></tr>`).join('') + '</table>';
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
    html += `<div class="cov-confidence">This is an AI comparison of your DWRs against the report — not a guarantee. Scan the “covered” list below to confirm each was really written up before you rely on it.</div>`;
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
      const rawCount = issues.length;
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

      // Honesty guard: differentiate "clean" from "the check didn't land".
      const note = document.getElementById('grammar-note');
      if (note) {
        if (rawCount > 0 && n === 0) {
          note.className = 'grammar-note warn';
          note.textContent = 'The checker returned suggestions, but none matched the text in the preview — it likely didn’t parse the report cleanly. Re-run it, or review the wording manually before trusting this.';
        } else if (rawCount === 0) {
          note.className = 'grammar-note';
          note.textContent = 'No sentences were flagged. This is not a guarantee the writing is perfect — only that the automatic check found nothing obvious. A quick human skim is still worth it.';
        } else {
          note.className = 'grammar-note';
          note.textContent = '';
        }
      }
      Toast.show(n ? n + ' sentence(s) flagged. Click each to review.' : 'Check complete — see the note below.', n ? '' : 'ok');
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
  // Build the full-screen review: corrected preview + list of changes.
  review() {
    const accepted = State.grammar.filter(g => g.status === 'accepted');
    // clone the (already corrected) preview so the user sees the final text
    document.getElementById('confirm-doc-body').innerHTML =
      document.getElementById('doc-body').innerHTML;
    // change list
    const list = document.getElementById('confirm-change-list');
    if (!accepted.length) {
      list.innerHTML = '<p class="muted" style="font-size:12.5px">No grammar fixes accepted — the copy will be identical to the original.</p>';
    } else {
      list.innerHTML = accepted.map(g => `
        <div class="chg">
          <span class="chg-was">${esc(g.original)}</span>
          <span class="chg-now">${esc(g.suggestion)}</span>
        </div>`).join('');
    }
    document.getElementById('confirm-sub').textContent =
      accepted.length + ' fix(es) will be written into a new copy named “' +
      (State.docTitle || 'WCR') + ' — Revised ' + new Date().toISOString().slice(0,10) + '”. The original Doc is never changed.';
    document.getElementById('confirm-result').classList.add('hidden');
    document.getElementById('confirm-create-btn').classList.remove('hidden');
    Screen.show('confirm');
  },
  cancelReview() { Screen.show('review'); },

  // Actually copy the Doc in Drive and apply the accepted fixes.
  async commit() {
    const btn = document.getElementById('confirm-create-btn');
    btn.disabled = true; btn.innerHTML = '<span class="spin"></span>Working…';
    try {
      const accepted = State.grammar.filter(g => g.status === 'accepted');

      if (CFG.DEMO_MODE || !State.docId) {
        await Export.ensureJSZip();
        const blob = await Export.buildFromPreview();
        Export._download(blob, (State.docTitle || 'WCR').replace(/[^\w\- ]+/g, '').replace(/\s+/g, '_') + '_revised.docx');
        Toast.show('Demo: built a clean revised .docx.', 'ok');
        return;
      }

      const date = new Date().toISOString().slice(0, 10);
      const copyName = (State.docTitle || 'WCR') + ' — Revised ' + date;
      btn.innerHTML = '<span class="spin"></span>Copying…';
      const copyId = await GoogleAPI.copyDoc(State.docId, copyName);

      let changed = 0;
      if (accepted.length) {
        btn.innerHTML = '<span class="spin"></span>Applying ' + accepted.length + ' fix(es)…';
        const reps = accepted.map(g => ({ from: g.original, to: g.suggestion }));
        const res = await GoogleAPI.applyReplacements(copyId, reps);
        (res.replies || []).forEach(rep => { changed += (rep.replaceAllText?.occurrencesChanged || 0); });
      }

      const url = 'https://docs.google.com/document/d/' + copyId + '/edit';
      Export._showCopyLink(url, accepted.length, changed);
      document.getElementById('confirm-create-btn').classList.add('hidden');
      History.add({
        id: 'r' + Date.now(), title: State.docTitle, at: Date.now(),
        dwrCount: State.dwrs.length, fixCount: accepted.length, copyUrl: url
      });
      Toast.show('Revised copy created in your Drive. Original untouched.', 'ok');
    } catch (e) {
      Toast.show(e.message || 'Could not create the revised copy.', 'err');
    } finally {
      btn.disabled = false; btn.textContent = 'Create revised copy in Drive';
    }
  },
  _showCopyLink(url, total, changed) {
    const note = document.getElementById('confirm-result');
    if (!note) { window.open(url, '_blank'); return; }
    note.classList.remove('hidden');
    const warn = (total && changed < total)
      ? `<div class="muted" style="margin-top:6px;font-size:11.5px">${changed} of ${total} fixes applied. A few sentences span Google’s internal formatting and couldn’t be matched exactly — open the copy to check those.</div>`
      : '';
    note.innerHTML = `<a class="primary-btn full" href="${url}" target="_blank" rel="noopener" style="display:block;text-align:center;text-decoration:none">Open revised copy in Google Docs</a>${warn}`;
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
  var bt = document.getElementById('build-tag'); if (bt) bt.textContent = 'build ' + BUILD;
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
