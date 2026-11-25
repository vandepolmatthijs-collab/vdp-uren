/* VdP Consulting - Tijdregistratie (web-app voor iPhone)
   Lokaal-eerst: alles wordt meteen in het toestel bewaard en daarna gesynchroniseerd met uren.json in OneDrive. */
(() => {
'use strict';

const CFG = window.VDP_CONFIG || {};
const DEMO = !CFG.clientId || new URLSearchParams(location.search).has('demo');
const SCOPES = ['User.Read', 'Files.ReadWrite', 'Mail.Send'];
const GRAPH = 'https://graph.microsoft.com/v1.0';
const LS_KEY = DEMO ? 'vdp-uren-demo' : 'vdp-uren-v1';
const DATA_PATH = (CFG.dataFolder || 'Apps/VdP Tijdregistratie') + '/uren.json';
const ROOT_RE = new RegExp(CFG.rootPattern || '^Projecten 20\\d\\d$');
const MIN_ENTRY_MINUTES = 1;

// ------------------------------------------------------------------ hulpfuncties
const $ = sel => document.querySelector(sel);
const pad = n => String(n).padStart(2, '0');
const iso = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
const parse = s => new Date(s);                       // lokale tijd (geen tijdzone in de string)
const monthKey = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}`;
const addMonths = (key, n) => { const [y, m] = key.split('-').map(Number); return monthKey(new Date(y, m - 1 + n, 1)); };
const monthLabel = key => { const [y, m] = key.split('-').map(Number); return new Date(y, m - 1, 1).toLocaleDateString('nl-BE', { month: 'long', year: 'numeric' }); };
const fmtDur = min => { const t = Math.round(min); return `${Math.floor(t / 60)}:${pad(t % 60)}`; };
const fmtHours = min => (Math.round(min / 60 * 100) / 100).toLocaleString('nl-BE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtClock = sec => `${pad(Math.floor(sec / 3600))}:${pad(Math.floor(sec % 3600 / 60))}:${pad(Math.floor(sec % 60))}`;
const fmtDate = d => d.toLocaleDateString('nl-BE', { day: '2-digit', month: '2-digit', year: 'numeric' });
const fmtDay = d => d.toLocaleDateString('nl-BE', { weekday: 'long', day: 'numeric', month: 'long' });
const fmtTime = d => `${pad(d.getHours())}:${pad(d.getMinutes())}`;
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const nowIso = () => iso(new Date());

function toast(msg, ms = 2600) {
  const t = $('#toast'); t.textContent = msg; t.hidden = false;
  clearTimeout(toast._t); toast._t = setTimeout(() => { t.hidden = true; }, ms);
}

// ------------------------------------------------------------------ projectnamen
// "26003_Ravels_Kerkstraat_RDK" -> nummer 2026003 (zoals in het urenblad), kort 26003, naam Ravels_Kerkstraat_RDK
function projectFromFolder(folder, root) {
  const m = folder.match(/^(\d{2})(\d{3})[\s_\-]+(.+)$/);
  const rawName = m ? m[3] : folder;
  return {
    id: `${root}/${folder}`, folder, root,
    short: m ? m[1] + m[2] : '',
    number: m ? '20' + m[1] + m[2] : '',
    name: rawName.replace(/\s+-\s+/g, '_').trim(),
    label: rawName.replace(/\s+-\s+|_/g, ' · ').trim()
  };
}

// ------------------------------------------------------------------ toestand
const defaultSettings = () => ({
  roots: [...(CFG.defaultRoots || ['Projecten 2026'])],
  recipient: '',
  rounding: 0,
  lastReportSent: addMonths(monthKey(new Date()), -1),  // eerste automatische rapport = huidige maand
  lastReportSentAt: '',
  dismissedRoots: [],
  updated: ''                                           // leeg = OneDrive wint bij de eerste synchronisatie
});

let S = load();
function load() {
  try {
    const s = JSON.parse(localStorage.getItem(LS_KEY));
    if (s && s.entries) return Object.assign({ projects: [], running: null, runningUpdated: '', lastSync: '' }, s,
      { settings: Object.assign(defaultSettings(), s.settings) });
  } catch (e) { /* leeg of geblokkeerd */ }
  return { entries: [], settings: defaultSettings(), running: null, runningUpdated: '', projects: [], lastSync: '' };
}
function persist() { try { localStorage.setItem(LS_KEY, JSON.stringify(S)); } catch (e) { /* vol of privémodus */ } }
function changed() { persist(); scheduleSync(); render(); }
const liveEntries = () => S.entries.filter(e => !e.deleted);

// ------------------------------------------------------------------ Microsoft 365 (MSAL + Graph)
let msalApp = null, account = null;

async function initAuth() {
  if (DEMO) return true;
  if (!window.msal) { setSync('err', 'offline'); return !!S.lastSync; }
  msalApp = new msal.PublicClientApplication({
    auth: { clientId: CFG.clientId, authority: `https://login.microsoftonline.com/${CFG.tenantId || 'organizations'}`,
            redirectUri: location.origin + location.pathname.replace(/index\.html$/, ''), navigateToLoginRequestUrl: false },
    cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
  });
  await msalApp.initialize();
  try {
    const res = await msalApp.handleRedirectPromise();
    if (res && res.account) account = res.account;
  } catch (e) { console.warn(e); toast('Aanmelden mislukt: ' + e.message, 5000); }
  account = account || msalApp.getAllAccounts()[0] || null;
  if (account) msalApp.setActiveAccount(account);
  return !!account;
}

function login() { msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' }); }
function logout() { msalApp.logoutRedirect({ account }); }

async function token() {
  try {
    const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account });
    return r.accessToken;
  } catch (e) {
    if (e instanceof msal.InteractionRequiredAuthError || /interaction_required|login_required|consent_required/.test(e.errorCode || '')) {
      await msalApp.acquireTokenRedirect({ scopes: SCOPES, account });
    }
    throw e;
  }
}

async function graph(path, opts = {}) {
  const t = await token();
  const res = await fetch(path.startsWith('http') ? path : GRAPH + path, {
    ...opts, headers: { Authorization: 'Bearer ' + t, ...(opts.headers || {}) }
  });
  return res;
}
const drivePath = p => p.split('/').map(encodeURIComponent).join('/');

// ------------------------------------------------------------------ projecten ophalen
const DEMO_FOLDERS = {
  'Projecten 2026': ['26001_Turnhout_Voorbeeldstraat 1_Janssens', '26002_Ravels_Dorpsstraat 12_Peeters', '26003_Retie_Testlaan_Maes',
    '26004_Beerse_Industrieweg_Demo NV', '26005_Oud-Turnhout_Kerkstraat 5_Claes', 'Arendonk - Voorbeeldweg 3 - Willems'],
  'Projecten 2027': ['27001_Mol_Nieuwstraat 8_Verhoeven']
};

async function listFolder(path) {
  const out = []; let url = `/me/drive/root:/${drivePath(path)}:/children?$select=name,folder&$top=999`;
  while (url) {
    const r = await graph(url);
    if (r.status === 404) return null;
    if (!r.ok) throw new Error('OneDrive ' + r.status);
    const j = await r.json();
    j.value.filter(i => i.folder).forEach(i => out.push(i.name));
    url = j['@odata.nextLink'] || null;
  }
  return out;
}

let rootSuggestions = [];
async function refreshProjects() {
  const projects = [];
  for (const root of S.settings.roots) {
    const names = DEMO ? (DEMO_FOLDERS[root] || null) : await listFolder(root);
    if (!names) continue;
    names.filter(n => !/^[.~]/.test(n)).forEach(n => projects.push(projectFromFolder(n, root)));
  }
  projects.sort((a, b) => b.root.localeCompare(a.root) || (!!b.number - !!a.number) || b.folder.localeCompare(a.folder, 'nl'));
  S.projects = projects;
  // nieuwe jaarmappen (bv. Projecten 2027) voorstellen
  let rootNames = [];
  if (DEMO) rootNames = Object.keys(DEMO_FOLDERS);
  else {
    const r = await graph('/me/drive/root/children?$select=name,folder&$top=999');
    if (r.ok) rootNames = (await r.json()).value.filter(i => i.folder).map(i => i.name);
  }
  rootSuggestions = rootNames.filter(n => ROOT_RE.test(n) && !S.settings.roots.includes(n) && !S.settings.dismissedRoots.includes(n));
  persist();
}

// ------------------------------------------------------------------ synchronisatie met OneDrive
let syncTimer = null, syncing = false, syncAgain = false;
function scheduleSync() { clearTimeout(syncTimer); syncTimer = setTimeout(syncNow, 1500); }

function setSync(cls, text) { const el = $('#syncState'); el.className = 'sync ' + cls; el.textContent = text; }

function merge(local, remote) {
  if (!remote) return local;
  const map = new Map();
  for (const e of [...(remote.entries || []), ...local.entries]) {
    const cur = map.get(e.id);
    if (!cur || (e.updated || '') > (cur.updated || '')) map.set(e.id, e);
  }
  const rs = remote.settings || {};
  const settings = (local.settings.updated || '') >= (rs.updated || '') ? local.settings : Object.assign(defaultSettings(), rs);
  const localRunNewer = (local.runningUpdated || '') >= (remote.runningUpdated || '');
  return {
    ...local, settings,
    entries: [...map.values()].sort((a, b) => a.start.localeCompare(b.start)),
    running: localRunNewer ? local.running : (remote.running || null),
    runningUpdated: localRunNewer ? local.runningUpdated : (remote.runningUpdated || '')
  };
}
const payload = s => ({ version: 1, settings: s.settings, entries: s.entries, running: s.running, runningUpdated: s.runningUpdated });

async function syncNow() {
  if (DEMO) { S.lastSync = nowIso(); persist(); setSync('ok', 'demo'); return true; }
  if (!account || !navigator.onLine) { setSync('err', 'offline'); return false; }
  if (syncing) { syncAgain = true; return false; }
  syncing = true; setSync('busy', 'synchroniseren…');
  try {
    for (let attempt = 0; attempt < 4; attempt++) {
      let remote = null, eTag = null;
      const meta = await graph(`/me/drive/root:/${drivePath(DATA_PATH)}?$select=eTag,@microsoft.graph.downloadUrl`);
      if (meta.ok) {
        const m = await meta.json(); eTag = m.eTag;
        const dl = await fetch(m['@microsoft.graph.downloadUrl'], { cache: 'no-store' });
        if (dl.ok) remote = JSON.parse(await dl.text() || 'null');
      } else if (meta.status !== 404) throw new Error('OneDrive ' + meta.status);

      const merged = merge(S, remote);
      const body = JSON.stringify(payload(merged), null, 1);
      const same = remote && body === JSON.stringify(payload(remote), null, 1);
      if (!same) {
        const put = await graph(`/me/drive/root:/${drivePath(DATA_PATH)}:/content`, {
          method: 'PUT', body, headers: { 'Content-Type': 'application/json', ...(eTag ? { 'If-Match': eTag } : {}) }
        });
        if (put.status === 412 || put.status === 409) continue;   // ander toestel schreef tussendoor: opnieuw samenvoegen
        if (!put.ok) throw new Error('Opslaan in OneDrive mislukt (' + put.status + ')');
      }
      S = Object.assign(merged, { lastSync: nowIso() }); persist();
      setSync('ok', 'bewaard ' + fmtTime(new Date()));
      if (tab !== 'projects' || !document.querySelector('#q:focus')) render();
      return true;
    }
    throw new Error('Synchronisatieconflict');
  } catch (e) {
    console.warn(e); setSync('err', 'niet gesynchroniseerd');
    return false;
  } finally {
    syncing = false;
    if (syncAgain) { syncAgain = false; scheduleSync(); }
  }
}

// ------------------------------------------------------------------ timer
function projectById(id) { return S.projects.find(p => p.id === id); }

function startProject(p) {
  if (S.running && S.running.projectId === p.id) { toast('Deze timer loopt al'); return; }
  const stopped = S.running ? stopRunning(true) : null;
  S.running = { projectId: p.id, number: p.number, short: p.short, name: p.name, label: p.label, folder: p.folder, root: p.root, start: nowIso(), note: '' };
  S.runningUpdated = nowIso();
  changed();
  toast(stopped ? `Gewisseld naar ${p.short || p.label}` : `Timer gestart: ${p.short || p.label}`);
}

function stopRunning(silent) {
  const r = S.running; if (!r) return null;
  const end = new Date(); const min = (end - parse(r.start)) / 60000;
  S.running = null; S.runningUpdated = nowIso();
  let entry = null;
  if (min >= MIN_ENTRY_MINUTES) {
    entry = { id: uid(), projectId: r.projectId, number: r.number, short: r.short, name: r.name, label: r.label, folder: r.folder, root: r.root,
              start: r.start, end: iso(end), note: (r.note || '').trim(), updated: nowIso() };
    S.entries.push(entry);
    if (!silent) toast(`${fmtDur(min)} geregistreerd op ${r.short || r.label}`);
  } else if (!silent) toast('Korter dan 1 minuut: niet bewaard');
  if (!silent) changed();
  return entry;
}

function entryMinutes(e, rounding = S.settings.rounding) {
  let m = Math.max(0, (parse(e.end) - parse(e.start)) / 60000);
  if (rounding > 0) m = Math.round(m / rounding) * rounding;
  return m;
}

// ------------------------------------------------------------------ weergave
let tab = 'projects', filterRoot = 'all', query = '', viewMonth = monthKey(new Date());

function render() { renderTimer(); renderView(); }

function renderTimer() {
  const el = $('#timerCard');
  const r = S.running;
  if (!r) { el.className = 'timer idle'; el.innerHTML = 'Geen timer actief. Tik op een project om te starten.'; return; }
  el.className = 'timer running';
  const start = parse(r.start);
  const long = (Date.now() - start) / 3600000 >= 10;
  if (!el.querySelector('.clock')) {
    el.innerHTML = `
      <div class="row">
        <div class="grow">
          <div class="label">Bezig sinds <span data-k="since"></span></div>
          <div class="proj" data-k="proj"></div>
          <div class="clock" data-k="clock"></div>
        </div>
        <button class="btn-stop" id="btnStop" aria-label="Timer stoppen"><span></span></button>
      </div>
      <input id="runNote" placeholder="Omschrijving werkzaamheden (optioneel)" enterkeyhint="done" autocomplete="off">
      <div class="warn" data-k="warn" hidden></div>`;
    $('#btnStop').onclick = () => stopRunning(false);
    const inp = $('#runNote');
    inp.oninput = () => { if (S.running) { S.running.note = inp.value; S.runningUpdated = nowIso(); persist(); scheduleSync(); } };
    inp.onkeydown = e => { if (e.key === 'Enter') inp.blur(); };
  }
  el.querySelector('[data-k=since]').textContent = (start.toDateString() === new Date().toDateString() ? '' : start.toLocaleDateString('nl-BE', { weekday: 'short', day: 'numeric', month: 'numeric' }) + ' ') + fmtTime(start);
  el.querySelector('[data-k=proj]').textContent = (r.short ? r.short + '  ' : '') + r.label;
  el.querySelector('[data-k=clock]').textContent = fmtClock((Date.now() - start) / 1000);
  const inp = $('#runNote'); if (document.activeElement !== inp && inp.value !== (r.note || '')) inp.value = r.note || '';
  const w = el.querySelector('[data-k=warn]');
  if (long) {
    w.hidden = false;
    if (!w.dataset.set) {
      w.dataset.set = 1;
      w.innerHTML = 'Deze timer loopt al meer dan 10 uur. Vergeten stoppen? <button class="btn small ghost" id="btnFix" style="margin-top:6px">Stoppen en eindtijd aanpassen</button>';
      $('#btnFix').onclick = () => { const e = stopRunning(true); changed(); if (e) openEditor(e); };
    }
  } else w.hidden = true;
}
setInterval(() => { if (S.running) renderTimer(); }, 1000);

function monthMinutesByProject(month) {
  const m = new Map();
  for (const e of liveEntries()) if (e.start.startsWith(month)) m.set(e.projectId, (m.get(e.projectId) || 0) + entryMinutes(e));
  return m;
}

function projectRow(p, monthMin) {
  const active = S.running && S.running.projectId === p.id;
  const min = monthMin.get(p.id) || 0;
  return `<button class="item ${active ? 'active' : ''}" data-start="${esc(p.id)}">
    <span class="nr ${p.short ? '' : 'none'}">${esc(p.short || '—')}</span>
    <span class="main"><span class="title">${esc(p.label)}</span>
      <span class="sub">${active ? 'loopt nu' : min ? 'deze maand ' + fmtDur(min) : esc(p.root)}</span></span>
    <span class="play"><svg viewBox="0 0 10 12"><path d="M0 0l10 6-10 6z"/></svg></span></button>`;
}

function renderProjects(v) {
  const monthMin = monthMinutesByProject(monthKey(new Date()));
  const roots = S.settings.roots;
  const q = query.trim().toLowerCase();
  let list = S.projects.filter(p => (filterRoot === 'all' || p.root === filterRoot) &&
    (!q || (p.folder + ' ' + p.number).toLowerCase().includes(q)));

  const recentIds = [];
  if (!q) for (const e of [...liveEntries()].sort((a, b) => b.start.localeCompare(a.start))) {
    if (!recentIds.includes(e.projectId) && projectById(e.projectId)) recentIds.push(e.projectId);
    if (recentIds.length >= 4) break;
  }
  const sugg = rootSuggestions.map(n => `<div class="suggest"><span class="grow">Nieuwe projectmap gevonden: <strong>${esc(n)}</strong></span>
    <button class="btn small" data-addroot="${esc(n)}">Toevoegen</button><button class="x" data-dismissroot="${esc(n)}" aria-label="Negeren">✕</button></div>`).join('');

  const html = `${sugg}
    <div class="search"><input id="q" type="search" placeholder="Zoek op nummer, gemeente, straat of klant" value="${esc(query)}" autocomplete="off"></div>
    ${roots.length > 1 ? `<div class="chips"><button class="chip ${filterRoot === 'all' ? 'on' : ''}" data-root="all">Alle</button>${roots.map(r =>
      `<button class="chip ${filterRoot === r ? 'on' : ''}" data-root="${esc(r)}">${esc(r.replace(/^Projecten\s*/, ''))}</button>`).join('')}</div>` : ''}
    <div id="plist"></div>`;
  if (!v.querySelector('#q')) v.innerHTML = html;
  else {   // zoekveld laten staan zodat het toetsenbord open blijft
    v.querySelectorAll('.suggest, .chips').forEach(n => n.remove());
    v.querySelector('.search').insertAdjacentHTML('beforebegin', sugg);
    if (roots.length > 1) v.querySelector('.search').insertAdjacentHTML('afterend', `<div class="chips"><button class="chip ${filterRoot === 'all' ? 'on' : ''}" data-root="all">Alle</button>${roots.map(r =>
      `<button class="chip ${filterRoot === r ? 'on' : ''}" data-root="${esc(r)}">${esc(r.replace(/^Projecten\s*/, ''))}</button>`).join('')}</div>`);
  }
  const recent = recentIds.map(projectById).filter(p => filterRoot === 'all' || p.root === filterRoot);
  v.querySelector('#plist').innerHTML =
    (recent.length ? `<h3>Recent</h3><div class="card">${recent.map(p => projectRow(p, monthMin)).join('')}</div>` : '') +
    (list.length ? `<h3>${q ? list.length + ' resultaten' : 'Alle projecten (' + list.length + ')'}</h3><div class="card">${list.map(p => projectRow(p, monthMin)).join('')}</div>`
      : `<div class="empty">${S.projects.length ? 'Geen project gevonden.' : 'Nog geen projecten geladen. Controleer de projectmappen bij Instellingen.'}</div>`);
  const qi = v.querySelector('#q');
  qi.oninput = () => { query = qi.value; renderView(); };
}

function monthNav() {
  return `<div class="monthnav"><button data-month="-1" aria-label="Vorige maand">‹</button><strong>${esc(monthLabel(viewMonth))}</strong>
    <button data-month="1" aria-label="Volgende maand">›</button></div>`;
}

function renderEntries(v) {
  const list = liveEntries().filter(e => e.start.startsWith(viewMonth)).sort((a, b) => b.start.localeCompare(a.start));
  const days = new Map();
  list.forEach(e => { const k = e.start.slice(0, 10); if (!days.has(k)) days.set(k, []); days.get(k).push(e); });
  v.innerHTML = monthNav() + `<button class="btn ghost" id="btnAdd">+ Uren handmatig toevoegen</button>` +
    (list.length ? [...days].map(([k, es]) => {
      const tot = es.reduce((s, e) => s + entryMinutes(e), 0);
      return `<h3>${esc(fmtDay(parse(k + 'T12:00:00')))} · ${fmtDur(tot)}</h3><div class="card">${es.map(e => `
        <button class="item" data-edit="${esc(e.id)}">
          <span class="nr ${e.short ? '' : 'none'}">${esc(e.short || '—')}</span>
          <span class="main"><span class="title">${esc(e.label || e.name)}</span>
            <span class="sub">${fmtTime(parse(e.start))}–${fmtTime(parse(e.end))}${e.note ? ' · ' + esc(e.note) : ''}</span></span>
          <span class="right">${fmtDur(entryMinutes(e))}</span></button>`).join('')}</div>`;
    }).join('') : `<div class="empty">Geen registraties in ${esc(monthLabel(viewMonth))}.</div>`);
  $('#btnAdd').onclick = () => openEditor(null);
}

function summary(month) {
  const rows = new Map();
  for (const e of liveEntries().filter(e => e.start.startsWith(month)).sort((a, b) => a.start.localeCompare(b.start))) {
    if (!rows.has(e.projectId)) rows.set(e.projectId, { projectId: e.projectId, number: e.number, short: e.short, name: e.name, label: e.label || e.name, minutes: 0, entries: [] });
    const r = rows.get(e.projectId); r.minutes += entryMinutes(e); r.entries.push(e);
  }
  return [...rows.values()].sort((a, b) => (!a.number - !b.number) || a.number.localeCompare(b.number) || a.name.localeCompare(b.name, 'nl'));
}

function renderReport(v) {
  const rows = summary(viewMonth);
  const total = rows.reduce((s, r) => s + r.minutes, 0);
  const st = S.settings;
  v.innerHTML = monthNav() + `
    <div class="card total"><span>Totaal gepresteerd</span><span><span class="big">${fmtHours(total)}</span> u</span></div>
    ${rows.length ? `<h3>Per project</h3><div class="card">${rows.map(r => `
      <div class="item"><span class="nr ${r.short ? '' : 'none'}">${esc(r.short || '—')}</span>
        <span class="main"><span class="title">${esc(r.label)}</span><span class="sub">${r.entries.length} registratie${r.entries.length > 1 ? 's' : ''} · ${fmtDur(r.minutes)}</span></span>
        <span class="right"><strong>${fmtHours(r.minutes)}</strong> u</span></div>`).join('')}</div>` : `<div class="empty">Geen uren in deze maand.</div>`}
    <button class="btn" id="btnMail" ${rows.length ? '' : 'disabled'}>Rapport nu mailen</button>
    <button class="btn ghost" id="btnXlsx" ${rows.length ? '' : 'disabled'}>Excel downloaden</button>
    <p class="muted">Het rapport van de vorige maand wordt automatisch gemaild vanaf de 1e van de maand, de eerste keer dat je de app opent.
      Laatst automatisch verstuurd: ${st.lastReportSentAt ? esc(monthLabel(st.lastReportSent)) + ' (op ' + esc(fmtDate(parse(st.lastReportSentAt))) + ')' : 'nog niet'}.
      ${st.rounding ? `Afronding: elke registratie op ${st.rounding} min.` : ''}</p>`;
  $('#btnMail').onclick = async () => {
    const to = st.recipient || (account && account.username) || '';
    if (!confirm(`Urenrapport ${monthLabel(viewMonth)} mailen naar ${to || 'jezelf'}?`)) return;
    try { await sendReport(viewMonth); toast('Rapport verstuurd'); } catch (e) { toast('Versturen mislukt: ' + e.message, 5000); }
  };
  $('#btnXlsx').onclick = () => { const wb = buildWorkbook(viewMonth); if (wb) XLSX.writeFile(wb, `Uren VdP Consulting ${viewMonth}.xlsx`); };
}

function renderSettings(v) {
  const st = S.settings;
  v.innerHTML = `
    <h3>Account</h3>
    <div class="card"><div class="rootrow"><span>${DEMO ? 'Demomodus' : esc(account ? (account.name || '') + ' · ' + account.username : 'Niet aangemeld')}</span>
      ${DEMO ? '' : '<button class="x" id="btnLogout">Afmelden</button>'}</div></div>

    <h3>Projectmappen in OneDrive</h3>
    <div class="card">${st.roots.map(r => `<div class="rootrow"><span>${esc(r)} <span class="muted">· ${S.projects.filter(p => p.root === r).length} projecten</span></span>
      <button class="x" data-delroot="${esc(r)}">Verwijderen</button></div>`).join('') || '<div class="rootrow muted">Geen mappen</div>'}</div>
    <div class="search"><input id="newRoot" placeholder="bv. Projecten 2027" autocomplete="off"><button class="btn small" id="btnAddRoot">Toevoegen</button></div>
    <p class="muted">Nieuwe projecten in deze mappen verschijnen vanzelf in de app. Een nieuwe map "Projecten 20xx" in OneDrive wordt automatisch voorgesteld.</p>

    <h3>Maandrapport</h3>
    <div class="card">
      <label class="field"><span>Versturen naar</span><input id="recipient" type="email" placeholder="${esc(account ? account.username : 'info@…')}" value="${esc(st.recipient)}"></label>
      <label class="field"><span>Afronding per registratie</span><select id="rounding">
        ${[[0, 'Geen (exacte minuten)'], [5, 'Op 5 minuten'], [15, 'Op kwartier'], [30, 'Op halfuur']].map(([v, t]) => `<option value="${v}" ${+st.rounding === v ? 'selected' : ''}>${t}</option>`).join('')}
      </select></label>
    </div>

    <h3>Synchronisatie</h3>
    <div class="card"><div class="rootrow"><span>Laatst bewaard in OneDrive</span><span class="muted">${S.lastSync ? esc(fmtDate(parse(S.lastSync)) + ' ' + fmtTime(parse(S.lastSync))) : '—'}</span></div></div>
    <button class="btn ghost" id="btnSync">Nu synchroniseren en projecten vernieuwen</button>
    <p class="muted">Op je iPhone: open deze pagina in Safari, tik op Deel en kies "Zet op beginscherm".</p>`;
  const set = (k, val) => { S.settings[k] = val; S.settings.updated = nowIso(); changed(); };
  $('#recipient').onchange = e => set('recipient', e.target.value.trim());
  $('#rounding').onchange = e => set('rounding', +e.target.value);
  $('#btnAddRoot').onclick = () => addRoot($('#newRoot').value.trim());
  $('#btnSync').onclick = async () => { await refreshAll(); toast('Bijgewerkt'); };
  if ($('#btnLogout')) $('#btnLogout').onclick = () => { if (confirm('Afmelden?')) logout(); };
}

async function addRoot(name) {
  if (!name) return;
  if (S.settings.roots.includes(name)) { toast('Deze map staat er al in'); return; }
  if (!DEMO) {
    try { if (!(await listFolder(name))) { toast(`Map "${name}" niet gevonden in OneDrive`, 4000); return; } }
    catch (e) { toast('Controle mislukt: ' + e.message, 4000); return; }
  }
  S.settings.roots = [...S.settings.roots, name]; S.settings.updated = nowIso();
  rootSuggestions = rootSuggestions.filter(n => n !== name);
  persist(); await refreshProjects().catch(() => {}); changed();
  toast(`${name} toegevoegd`);
}

function renderView() {
  const v = $('#view');
  if (renderView.tab !== tab) { v.innerHTML = ''; renderView.tab = tab; document.body.className = 'tab-' + tab; }
  if (needLogin) return renderLogin(v);
  ({ projects: renderProjects, entries: renderEntries, report: renderReport, settings: renderSettings })[tab](v);
}

function renderLogin(v) {
  v.innerHTML = `<div class="login"><img src="logo.png" alt="VdP Consulting"><h2>Tijdregistratie</h2>
    <p class="muted">Meld je aan met je Microsoft 365-account. De app leest je projectmappen in OneDrive en bewaart je uren daar.</p>
    <button class="btn" id="btnLogin">Aanmelden met Microsoft 365</button></div>`;
  $('#btnLogin').onclick = login;
}

// ------------------------------------------------------------------ bewerken
function openEditor(entry) {
  const isNew = !entry;
  const now = new Date();
  const e = entry || { projectId: S.running ? '' : (S.projects[0] || {}).id, start: iso(new Date(now - 3600000)), end: iso(now), note: '' };
  const s = parse(e.start), en = parse(e.end);
  const roots = [...new Set(S.projects.map(p => p.root))];
  const known = projectById(e.projectId);
  const opts = (known || isNew ? '' : `<option value="${esc(e.projectId)}" selected>${esc((e.short ? e.short + ' ' : '') + (e.label || e.name))} (map niet meer gevonden)</option>`) +
    roots.map(r => `<optgroup label="${esc(r)}">${S.projects.filter(p => p.root === r).map(p =>
      `<option value="${esc(p.id)}" ${p.id === e.projectId ? 'selected' : ''}>${esc((p.short ? p.short + '  ' : '') + p.label)}</option>`).join('')}</optgroup>`).join('');
  const sh = $('#sheet');
  sh.innerHTML = `<div class="panel">
    <h2>${isNew ? 'Uren toevoegen' : 'Registratie bewerken'}</h2>
    <div class="card">
      <label class="field"><span>Project</span><select id="edProj">${opts}</select></label>
      <label class="field"><span>Datum</span><input id="edDate" type="date" value="${e.start.slice(0, 10)}"></label>
      <div class="two"><label class="field"><span>Van</span><input id="edFrom" type="time" value="${fmtTime(s)}"></label>
        <label class="field"><span>Tot</span><input id="edTo" type="time" value="${fmtTime(en)}"></label></div>
      <label class="field"><span>Omschrijving werkzaamheden</span><textarea id="edNote" rows="2">${esc(e.note)}</textarea></label>
    </div>
    <button class="btn" id="edSave">Bewaren</button>
    ${isNew ? '' : '<button class="btn danger" id="edDel">Verwijderen</button>'}
    <button class="btn ghost" id="edCancel">Annuleren</button></div>`;
  sh.hidden = false;
  const close = () => { sh.hidden = true; sh.innerHTML = ''; };
  sh.onclick = ev => { if (ev.target === sh) close(); };
  $('#edCancel').onclick = close;
  $('#edSave').onclick = () => {
    const pid = $('#edProj').value, date = $('#edDate').value, from = $('#edFrom').value, to = $('#edTo').value;
    if (!pid || !date || !from || !to) { toast('Vul project, datum en uren in'); return; }
    const start = new Date(`${date}T${from}:00`); let end = new Date(`${date}T${to}:00`);
    if (end <= start) end = new Date(end.getTime() + 86400000);             // over middernacht
    if ((end - start) / 3600000 > 16) { toast('Meer dan 16 uur: controleer de tijden'); return; }
    const p = projectById(pid) || e;
    const target = isNew ? { id: uid() } : S.entries.find(x => x.id === e.id);
    Object.assign(target, { projectId: pid, number: p.number, short: p.short, name: p.name, label: p.label || p.name, folder: p.folder, root: p.root,
      start: iso(start), end: iso(end), note: $('#edNote').value.trim(), updated: nowIso() });
    if (isNew) S.entries.push(target);
    close(); changed(); toast('Bewaard');
  };
  if (!isNew) $('#edDel').onclick = () => {
    if (!confirm('Deze registratie verwijderen?')) return;
    const t = S.entries.find(x => x.id === e.id); t.deleted = true; t.updated = nowIso();
    close(); changed(); toast('Verwijderd');
  };
}

// ------------------------------------------------------------------ rapport (mail + Excel)
function reportHtml(month) {
  const rows = summary(month); const total = rows.reduce((s, r) => s + r.minutes, 0);
  const rounding = S.settings.rounding;
  let h = `<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000">
<p style="font-size:20pt;color:#C66518;margin:0 0 4px 0">Gepresteerde uren &ndash; ${esc(monthLabel(month))}</p>
<p style="color:#404040;margin:0 0 16px 0">VdP Consulting Ruimtelijke Ordening &middot; overzicht voor facturatie</p>
<table cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt">
<tr style="background:#C66518;color:#ffffff"><th align="left">Projectnr.</th><th align="left">Project</th><th align="right">Uren</th><th align="right">u:mm</th></tr>`;
  rows.forEach((r, i) => {
    h += `<tr style="background:${i % 2 ? '#FBEFE4' : '#FFFFFF'}"><td>${esc(r.number)}</td><td>${esc(r.name)}</td><td align="right">${fmtHours(r.minutes)}</td><td align="right">${fmtDur(r.minutes)}</td></tr>`;
  });
  h += `<tr style="border-top:2px solid #C66518;font-weight:bold"><td></td><td>Totaal</td><td align="right">${fmtHours(total)}</td><td align="right">${fmtDur(total)}</td></tr></table>`;
  if (rounding) h += `<p style="color:#404040;font-size:9pt">Elke registratie afgerond op ${rounding} minuten.</p>`;
  h += `<p style="font-size:14pt;color:#944B12;margin:24px 0 6px 0">Detail per project</p>`;
  rows.forEach(r => {
    h += `<p style="color:#C66518;font-weight:bold;margin:12px 0 2px 0">${esc(r.number)} ${esc(r.name)} &ndash; ${fmtHours(r.minutes)} u</p>
<table cellpadding="3" cellspacing="0" style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#404040">`;
    r.entries.forEach(e => {
      h += `<tr><td>${fmtDate(parse(e.start))}</td><td>${fmtTime(parse(e.start))}&ndash;${fmtTime(parse(e.end))}</td><td align="right">${fmtDur(entryMinutes(e))}</td><td>${esc(e.note)}</td></tr>`;
    });
    h += `</table>`;
  });
  h += `<p style="color:#404040;font-size:9pt;margin-top:24px">Verstuurd door VdP Tijdregistratie. Het detail staat ook in de Excel-bijlage (kolommen zoals in "Overzicht uren").</p></div>`;
  return h;
}

function buildWorkbook(month) {
  if (!window.XLSX) { toast('Excel-module nog niet geladen'); return null; }
  const rows = summary(month);
  const total = rows.reduce((s, r) => s + r.minutes, 0);
  const round2 = m => Math.round(m / 60 * 100) / 100;
  const ov = [[`Gepresteerde uren ${monthLabel(month)}`], [], ['Projectnummer', 'Naam', 'Uren', 'Duur (u:mm)', 'Aantal registraties'],
    ...rows.map(r => [r.number, r.name, round2(r.minutes), fmtDur(r.minutes), r.entries.length]), ['', 'Totaal', round2(total), fmtDur(total), '']];
  const det = [['Projectnummer', 'Naam', 'Omschrijving werkzaamheden', 'Datum', 'Uren', 'Van', 'Tot'],
    ...rows.flatMap(r => r.entries.map(e => [e.number, e.name, e.note, fmtDate(parse(e.start)), round2(entryMinutes(e)), fmtTime(parse(e.start)), fmtTime(parse(e.end))]))];
  const wb = XLSX.utils.book_new();
  const ws1 = XLSX.utils.aoa_to_sheet(ov); ws1['!cols'] = [{ wch: 14 }, { wch: 48 }, { wch: 8 }, { wch: 12 }, { wch: 18 }];
  const ws2 = XLSX.utils.aoa_to_sheet(det); ws2['!cols'] = [{ wch: 14 }, { wch: 40 }, { wch: 50 }, { wch: 11 }, { wch: 7 }, { wch: 6 }, { wch: 6 }];
  XLSX.utils.book_append_sheet(wb, ws1, 'Overzicht');
  XLSX.utils.book_append_sheet(wb, ws2, 'Detail');
  return wb;
}

async function sendReport(month) {
  const html = reportHtml(month);
  if (DEMO) { const w = window.open('', '_blank'); if (w) w.document.write(html); return; }
  const to = S.settings.recipient || account.username;
  const attachments = [];
  if (window.XLSX) {
    const b64 = XLSX.write(buildWorkbook(month), { type: 'base64', bookType: 'xlsx' });
    attachments.push({ '@odata.type': '#microsoft.graph.fileAttachment', name: `Uren VdP Consulting ${month}.xlsx`,
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: b64 });
  }
  const r = await graph('/me/sendMail', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ message: { subject: `Urenoverzicht ${monthLabel(month)} - VdP Consulting`, body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: to } }], attachments }, saveToSentItems: true })
  });
  if (!r.ok) throw new Error('Outlook ' + r.status);
}

// Rapport van de vorige maand automatisch versturen (eenmalig, vanaf de 1e).
async function checkDueReport() {
  const prev = addMonths(monthKey(new Date()), -1);
  const st = S.settings;
  if ((st.lastReportSent || '') >= prev) return;
  await new Promise(r => setTimeout(r, 1500));            // Excel-bibliotheek laten laden
  const hasEntries = liveEntries().some(e => e.start.startsWith(prev));
  try {
    if (hasEntries) await sendReport(prev);
    st.lastReportSent = prev; st.lastReportSentAt = hasEntries ? nowIso() : st.lastReportSentAt; st.updated = nowIso();
    changed(); await syncNow();
    if (hasEntries) toast(`Urenrapport ${monthLabel(prev)} is naar je mailbox verstuurd`, 4500);
  } catch (e) { console.warn(e); toast('Automatisch maandrapport kon niet verstuurd worden; probeer het via Overzicht.', 5000); }
}

// ------------------------------------------------------------------ gebeurtenissen
document.addEventListener('click', ev => {
  const b = ev.target.closest('button'); if (!b) return;
  if (b.dataset.tab) {
    tab = b.dataset.tab; document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('active', x === b));
    if (tab === 'report' || tab === 'entries') viewMonth = viewMonth || monthKey(new Date());
    renderView(); window.scrollTo(0, 0);
  } else if (b.dataset.start) { const p = projectById(b.dataset.start); if (p) startProject(p); }
  else if (b.dataset.edit) { const e = S.entries.find(x => x.id === b.dataset.edit); if (e) openEditor(e); }
  else if (b.dataset.root) { filterRoot = b.dataset.root; renderView(); }
  else if (b.dataset.month) { viewMonth = addMonths(viewMonth, +b.dataset.month); renderView(); }
  else if (b.dataset.addroot) addRoot(b.dataset.addroot);
  else if (b.dataset.dismissroot) {
    S.settings.dismissedRoots = [...S.settings.dismissedRoots, b.dataset.dismissroot]; S.settings.updated = nowIso();
    rootSuggestions = rootSuggestions.filter(n => n !== b.dataset.dismissroot); changed();
  } else if (b.dataset.delroot) {
    if (!confirm(`${b.dataset.delroot} uit de app halen? Je geregistreerde uren blijven bewaard.`)) return;
    S.settings.roots = S.settings.roots.filter(r => r !== b.dataset.delroot); S.settings.updated = nowIso();
    S.projects = S.projects.filter(p => p.root !== b.dataset.delroot); if (filterRoot === b.dataset.delroot) filterRoot = 'all';
    changed();
  }
});

async function refreshAll() {
  await syncNow();
  try { await refreshProjects(); } catch (e) { console.warn(e); }
  render();
}

document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && !needLogin) refreshAll().then(checkDueReport); });
window.addEventListener('online', () => scheduleSync());
setInterval(() => { if (!needLogin) refreshAll(); }, 5 * 60 * 1000);

// ------------------------------------------------------------------ start
let needLogin = false;
(async () => {
  if (DEMO) {
    $('#demoBanner').hidden = false;
    if (!S.entries.length) seedDemo();
  }
  render();
  const ok = await initAuth();
  if (!ok && !DEMO) { needLogin = true; renderView(); return; }
  await refreshAll();
  await checkDueReport();
})();

if ('serviceWorker' in navigator && location.protocol === 'https:') navigator.serviceWorker.register('sw.js').catch(() => {});

function seedDemo() {
  const d = new Date(); const mk = (dayOffset, h1, m1, h2, m2, folder, root, note) => {
    const b = new Date(d.getFullYear(), d.getMonth(), d.getDate() - dayOffset);
    const p = projectFromFolder(folder, root);
    return { id: uid(), projectId: p.id, number: p.number, short: p.short, name: p.name, label: p.label, folder, root,
      start: iso(new Date(b.getFullYear(), b.getMonth(), b.getDate(), h1, m1)), end: iso(new Date(b.getFullYear(), b.getMonth(), b.getDate(), h2, m2)), note, updated: nowIso() };
  };
  S.entries = [
    mk(0, 8, 30, 10, 15, '26002_Ravels_Dorpsstraat 12_Peeters', 'Projecten 2026', 'Opmaak beschrijvende nota'),
    mk(1, 13, 0, 16, 45, '26004_Beerse_Industrieweg_Demo NV', 'Projecten 2026', 'Overleg gemeente + voorbereiding'),
    mk(2, 9, 0, 11, 0, '26002_Ravels_Dorpsstraat 12_Peeters', 'Projecten 2026', 'Plaatsbezoek'),
    mk(35, 9, 0, 12, 30, '26001_Turnhout_Voorbeeldstraat 1_Janssens', 'Projecten 2026', 'Beroepsschrift - versie 1')
  ];
  persist();
}
})();
