/* Pilotage Victor — répartition équitable et suivi d'un fichier Excel en équipe.
   Tout se passe dans le navigateur : le fichier est lu en mémoire, rien n'est envoyé ni enregistré.
   Fermer l'onglet (ou « Fermer le fichier ») efface toutes les données. */
'use strict';

const $ = (s, r = document) => r.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const fmt = n => Number(n).toLocaleString('fr-FR');
const pct = (a, b) => b ? Math.round(100 * a / b) : 0;
const DAY = 86400000;

const S = { wb: null, fileName: '', sheets: [], cur: 0, tab: 'dash', agentFilter: '', statutFilter: '', search: '', motif: null, absents: new Set(), parAgent: 0, toutRedistribuer: false, more: false };

/* ---------- 1. Lecture du fichier ---------- */

const drop = $('#drop'), fileInput = $('#file');
$('#pick').onclick = () => fileInput.click();
fileInput.onchange = e => e.target.files[0] && loadFile(e.target.files[0]);
drop.addEventListener('dragover', e => { e.preventDefault(); drop.classList.add('over'); });
drop.addEventListener('dragleave', () => drop.classList.remove('over'));
drop.addEventListener('drop', e => { e.preventDefault(); drop.classList.remove('over'); e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]); });
$('#close').onclick = () => { S.wb = null; S.sheets = []; S.fileName = ''; fileInput.value = ''; $('#view').innerHTML = ''; $('#app').hidden = true; $('#tabs').hidden = true; $('#body').classList.add('solo'); drop.hidden = false; };
$('#theme').onclick = () => {
  const r = document.documentElement, dark = r.dataset.theme ? r.dataset.theme === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches;
  r.dataset.theme = dark ? 'light' : 'dark';
  render();
};
$('#refdate').onchange = e => { const sh = cur(); if (sh && e.target.value) { sh.ref = new Date(e.target.value + 'T00:00:00'); compute(sh); render(); } };
$('#tabs').onclick = e => { const b = e.target.closest('button'); if (!b) return; S.tab = b.dataset.tab; render(); };

async function loadFile(f) {
  try {
    const wb = XLSX.read(new Uint8Array(await f.arrayBuffer()), { type: 'array', cellDates: true, cellStyles: true });
    S.wb = wb; S.fileName = f.name;
    S.sheets = wb.SheetNames.map(n => parseSheet(n, wb.Sheets[n]));
    S.cur = Math.max(0, S.sheets.findIndex(s => s.ok));
    initPrep();
    S.absents = new Set(); S.parAgent = 0; S.toutRedistribuer = false; S.agentFilter = ''; S.motif = null;
    drop.hidden = true; $('#app').hidden = false; $('#tabs').hidden = false; $('#body').classList.remove('solo');
    render();
  } catch (err) {
    alert('Impossible de lire ce fichier : ' + err.message);
  }
}

function parseSheet(name, ws) {
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const start0 = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s : { r: 0, c: 0 };
  // Un texte dans une colonne au format « date » peut ressortir en date invalide : on reprend le texte d'origine
  aoa.forEach((r, i) => r.forEach((v, j) => { if (v instanceof Date && isNaN(v)) { const c = ws[XLSX.utils.encode_cell({ r: start0.r + i, c: start0.c + j })]; r[j] = c ? (c.w ?? c.v ?? null) : null; if (r[j] instanceof Date) r[j] = null; } }));
  const head = aoa.slice(0, 8).flat().map(norm).join(' | ');
  if (/etiquettes de lignes|total general/.test(head)) return { name, ok: false, why: 'tableau croisé (ignoré)' };
  const h = aoa.findIndex((r, i) => i < 10 && r.filter(c => typeof c === 'string' && c.trim()).length >= 3);
  if (h < 0) return { name, ok: false, why: 'pas de ligne de titres' };
  const start = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s : { r: 0, c: 0 };
  const headers = aoa[h].map(c => String(c ?? '').trim());
  const all = aoa.slice(h + 1).map((r, k) => ({ r, row: start.r + h + 1 + k }));
  const kept = all.filter(x => x.r.some(c => c !== null && String(c).trim() !== ''));
  if (!kept.length) return { name, ok: false, why: 'aucune ligne' };
  const sh = {
    name, ok: true, ws, headers, raw: kept.map(x => x.r), sheetRow: kept.map(x => x.row),
    r0: start.r, c0: start.c, headerIdx: start.r + h, headerRow: start.r + h + 1, empty: all.length - kept.length,
    yellow: yellowCells(ws), t2Label: team2Label(headers),
  };
  sh.map = detect(headers, sh.raw);
  sh.ref = guessRef(name, sh);
  compute(sh);
  return sh;
}

/* Lignes surlignées en jaune (la décision « à envoyer à l'équipe 2 ») */
function isYellow(s) {
  const rgb = s && s.patternType === 'solid' && s.fgColor && s.fgColor.rgb;
  if (!rgb || rgb.length < 6) return false;
  const h = rgb.slice(-6), R = parseInt(h.slice(0, 2), 16), G = parseInt(h.slice(2, 4), 16), B = parseInt(h.slice(4, 6), 16);
  return R >= 200 && G >= 200 && B <= 150;
}
function yellowCells(ws) {
  const n = new Map();
  for (const k in ws) { if (k[0] === '!' || !isYellow(ws[k].s)) continue; const r = XLSX.utils.decode_cell(k).r; n.set(r, (n.get(r) || 0) + 1); }
  return n;
}

/* Nom de la 2e équipe, lu dans les titres du fichier (« traitement par X », « compétence X ») */
function team2Label(headers) {
  for (const h of headers) {
    const m = h.match(/(?:traitement par|comp[ée]tence)\s+(?:les?\s+|l')?([A-Za-zÀ-ÿ0-9][\wÀ-ÿ-]{1,20})/i);
    if (m) return m[1];
  }
  return 'équipe 2';
}

/* ---------- 2. Reconnaissance automatique des colonnes ---------- */

const FIELDS = [
  ['affaire', "N° d'affaire"], ['compte', 'N° de compte'], ['date', 'Date de génération (ancienneté)'], ['dateStock', 'Date de stock'],
  ['circuit', 'Circuit'], ['categorie', 'Catégorie'], ['dr', 'DR / site'], ['unite', 'Unité'],
  ['agent', 'Agent'], ['traite', 'Traité oui/non'], ['comment', 'Commentaire'],
  ['t2Flag', 'Compétence équipe 2 (oui/non)'], ['t2Comment', 'Commentaire équipe 2'], ['t2Agent', 'Agent équipe 2'], ['t2Traite', 'Traité (équipe 2)'],
];

function detect(headers, rows) {
  const H = headers.map(norm);
  const find = (pred, from = 0) => H.findIndex((h, j) => j >= from && pred(h));
  const firstOf = (...preds) => { for (const p of preds) { const i = find(p); if (i >= 0) return i; } return -1; };
  const m = {};
  m.compte = firstOf(h => h.includes('compte') && (h.includes('num') || h.includes('externe')), h => h === 'compte');
  m.affaire = find(h => h.includes('affaire'));
  m.date = firstOf(h => h.includes('generation'), h => h.startsWith('anciennete'), h => h === 'date de creation');
  m.dateStock = find(h => h.replace(/ /g, '') === 'datestock');
  m.circuit = firstOf(h => h.includes('libelle du circuit'), h => h.includes('circuit'));
  m.categorie = firstOf(h => h.includes('libelle categorie'), h => h.includes('categorie'));
  m.dr = firstOf(h => h === 'dr', h => h === 'site' || h === 'region');
  m.unite = firstOf(h => h.includes("niveau d'organisation") && h.includes('libelle'), h => h === 'unite' || h === 'unites');
  m.agent = find(h => /^agents?$/.test(h));
  m.traite = firstOf(h => h.startsWith('traite oui'), h => h.includes('cloturee'), h => h.startsWith('traite'));
  m.comment = firstOf(h => h.startsWith('si non'), h => h.includes('commentaire'), h => h.includes('motif'));
  m.t2Flag = find(h => /^competence \S/.test(h));
  m.t2Comment = m.comment >= 0 ? find(h => h.startsWith('traitement par') || h.startsWith('commentaire'), m.comment + 1) : -1;
  // Agent de l'équipe 2 : colonne sans vrai titre (« Colonne1 »…) qui contient des prénoms
  m.t2Agent = H.findIndex((h, j) => j > Math.max(m.agent, m.comment) && /^(colonne ?\d*|gestionnaire.*|agent .+)$/.test(h) && looksLikeNames(rows.map(r => r[j])));
  m.t2Traite = m.t2Agent >= 0 ? find(h => h.startsWith('traite'), m.t2Agent + 1) : -1;
  return m;
}

function looksLikeNames(vals) {
  const v = vals.filter(x => x != null && String(x).trim()).map(x => String(x).trim());
  if (!v.length) return false;
  const bad = v.filter(x => /\d/.test(x) || x.split(/\s+/).length > 3 || /^(oui|non|traite)/.test(norm(x)));
  return bad.length === 0 && new Set(v.map(norm)).size <= 15;
}

function guessRef(name, sh) {
  const i = sh.map.dateStock;
  if (i >= 0) { const d = sh.raw.map(r => toDate(r[i])).filter(Boolean).sort((a, b) => b - a)[0]; if (d) return d; }
  const m = name.match(/(\d{2})[-/.](\d{2})[-/.](\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  const j = sh.map.date;
  if (j >= 0) { const d = sh.raw.map(r => toDate(r[j])).filter(Boolean).sort((a, b) => b - a)[0]; if (d) return d; }
  return new Date();
}

function toDate(v) {
  if (v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(), v.getMonth(), v.getDate() + (v.getHours() >= 12 ? 1 : 0));
  if (typeof v === 'number' && v > 20000 && v < 80000) return new Date(Math.round((v - 25569) * DAY));
  if (typeof v === 'string') { const m = v.match(/^(\d{2})\/(\d{2})\/(\d{4})/); if (m) return new Date(+m[3], +m[2] - 1, +m[1]); }
  return null;
}

/* ---------- 3. Nettoyage et calculs ---------- */

function statut(v) {
  const n = norm(v);
  if (!n) return 'vide';
  if (/^(oui|traite|cloture|deja traite)/.test(n)) return 'oui';
  if (/^(non|pas)\b/.test(n)) return 'non';
  return 'autre';
}
const STATUT_LABEL = { oui: 'Traité', non: 'Non traité', vide: 'Non renseigné', autre: 'Autre' };

/* Regroupement automatique des commentaires libres : aucun mot-clé écrit à l'avance,
   on cherche les mots (ou paires de mots) qui reviennent le plus, et on nomme chaque groupe
   par le commentaire le plus fréquent qu'il contient. */
const STOP = new Set(('a au aux avec ce ces cet cette dans de des du elle en est et il ils la le les leur lui mais me meme mes mon ne nous on ou par ' +
  'pour qu que qui sa se ses son sur ta te tes ton tu un une vous y ete etre fait faire ai as avons ont sont suite apres avant tout tous toute toutes ' +
  'plus moins tres bien ci ca cela comme jour cours').split(' '));
const tokens = t => norm(t).replace(/['’]/g, ' ').replace(/[^a-z0-9 ]/g, ' ').split(' ')
  .filter(w => w.length >= 3 && !/\d/.test(w) && !STOP.has(w)).map(w => w.replace(/s$/, ''));

function clusterComments(items) {
  let rest = items.map(x => { const t = tokens(x.text), f = new Set(t); for (let i = 0; i < t.length - 1; i++) f.add(t[i] + ' ' + t[i + 1]); return { x, f }; });
  const groups = [];
  while (groups.length < 10 && rest.length) {
    const df = new Map();
    for (const d of rest) for (const f of d.f) df.set(f, (df.get(f) || 0) + 1);
    // d'abord les paires de mots (« en attente »), plus parlantes qu'un mot seul (« pas »)
    let best = null, score = 0;
    for (const [f, n] of df) if (n >= 3 && f.includes(' ') && n > score) { score = n; best = f; }
    if (!best) for (const [f, n] of df) if (n >= 3 && n > score) { score = n; best = f; }
    if (!best) break;
    groups.push(rest.filter(d => d.f.has(best)).map(d => d.x));
    rest = rest.filter(d => !d.f.has(best));
  }
  const out = new Map(), used = new Set();
  for (const g of groups) {
    const freq = new Map();
    for (const x of g) { const k = norm(x.text); const e = freq.get(k) || { n: 0, txt: x.text.trim() }; e.n++; freq.set(k, e); }
    let label = [...freq.values()].sort((a, b) => b.n - a.n)[0].txt
      .replace(/\(?\b(?=[A-Z0-9]*\d)[A-Z0-9]{8,}\b\)?/g, '').replace(/\s+/g, ' ').trim(); // sans les numéros
    if (label.length > 60) label = label.slice(0, 59) + '…';
    while (used.has(label)) label += ' ';
    used.add(label);
    out.set(label, g);
  }
  if (rest.length) out.set('Autres commentaires', rest.map(d => d.x));
  return out;
}

function compute(sh) {
  const m = sh.map, get = (r, k) => m[k] >= 0 ? r[m[k]] : null;
  const names = new Map(), t2Names = new Map(); // clé sans accents -> formes écrites
  const addName = (map, v) => { const k = norm(v); if (!k) return null; const f = map.get(k) || new Map(); const t = String(v).trim().replace(/\s+/g, ' '); f.set(t, (f.get(t) || 0) + 1); map.set(k, f); return k; };
  sh.rows = sh.raw.map((r, i) => {
    const d = toDate(get(r, 'date')) || toDate(get(r, 'dateStock'));
    const filled = r.filter(c => c !== null && String(c).trim() !== '').length;
    const yc = sh.yellow.get(sh.sheetRow[i]) || 0;
    const marque = yc >= 3 && yc >= 0.3 * filled;
    const fv = get(r, 't2Flag'), flag = fv === true || /^(oui|vrai|true)/.test(norm(fv));
    const t2AgentKey = addName(t2Names, get(r, 't2Agent'));
    const t2Comment = String(get(r, 't2Comment') ?? '').trim();
    const gT = norm(get(r, 't2Traite'));
    let t2Stat = null;
    if (t2AgentKey || t2Comment) t2Stat = m.t2Traite >= 0 ? (/^(oui|traite|deja)/.test(gT) ? 'traite' : /^non/.test(gT) ? 'retour' : 'encours') : (t2Comment ? 'traite' : 'encours');
    else if (marque) t2Stat = 'attente';
    else if (flag) t2Stat = 'propose';
    const traiteRaw = get(r, 'traite');
    return {
      i, raw: r, date: d, age: d ? Math.round((sh.ref - d) / DAY) : null,
      affaire: String(get(r, 'affaire') ?? '').trim(), compte: String(get(r, 'compte') ?? '').trim(),
      circuit: String(get(r, 'circuit') ?? '').trim() || sh.name,
      categorie: String(get(r, 'categorie') ?? '').trim(), dr: String(get(r, 'dr') ?? '').trim(), unite: String(get(r, 'unite') ?? '').trim(),
      agentKey: addName(names, get(r, 'agent')),
      traiteRaw: traiteRaw == null ? '' : String(traiteRaw), statut: statut(traiteRaw),
      deja: /deja trait/.test(norm(traiteRaw)) || /deja trait|autre agent/.test(norm(get(r, 'comment'))),
      comment: String(get(r, 'comment') ?? '').trim(),
      marque, t2Stat, t2: !!t2Stat && t2Stat !== 'propose', t2AgentKey, t2Comment,
    };
  });
  const display = map => { const o = {}; for (const [k, forms] of map) o[k] = [...forms].sort((a, b) => b[1] - a[1] || /[^\x00-\x7f]/.test(b[0]) - /[^\x00-\x7f]/.test(a[0]))[0][0]; return o; };
  sh.agentName = display(names); sh.t2Name = display(t2Names);
  sh.agents = Object.keys(sh.agentName).sort((a, b) => count(sh.rows, r => r.agentKey === b) - count(sh.rows, r => r.agentKey === a));
  sh.rows.forEach(r => { r.agent = r.agentKey ? sh.agentName[r.agentKey] : ''; r.t2Agent = r.t2AgentKey ? sh.t2Name[r.t2AgentKey] : ''; r.motif = null; });
  const toGroup = sh.rows.filter(r => r.statut !== 'oui' || r.t2Stat).map(r => ({ r, text: r.comment || (/^(non|oui)?$/.test(norm(r.traiteRaw)) ? '' : r.traiteRaw) })).filter(x => x.text);
  for (const [label, list] of clusterComments(toGroup)) list.forEach(x => { x.r.motif = label; });
}

const count = (arr, f) => arr.reduce((n, x) => n + (f(x) ? 1 : 0), 0);
const groupBy = (arr, f) => { const m = new Map(); for (const x of arr) { const k = f(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); } return m; };
const mean = a => a.length ? a.reduce((s, x) => s + x, 0) / a.length : 0;
const median = a => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y), h = s.length >> 1; return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2; };
const cur = () => S.sheets[S.cur];

/* Répartition 1-1-1-1 : on prend les lignes de la plus ancienne à la plus récente et on les distribue
   une à une, à tour de rôle, en équilibrant chaque circuit. Un compte déjà donné à un agent lui revient
   (doublons). Si un plafond par agent est fixé, les lignes les plus récentes restent non attribuées. */
function roundRobin(rows, agents, cap, seed = new Map()) {
  const res = new Map(agents.map(a => [a, []])), pool = [];
  if (!agents.length) return { res, pool: [...rows] };
  const perCircuit = new Map(), byCompte = new Map(seed), full = a => cap && res.get(a).length >= cap;
  const sorted = [...rows].sort((a, b) => (a.date ?? Infinity) - (b.date ?? Infinity));
  let k = 0;
  for (const r of sorted) {
    const pc = perCircuit.get(r.circuit) || new Map(); perCircuit.set(r.circuit, pc);
    let best = r.compte && byCompte.get(r.compte);
    if (!best || full(best)) {
      best = null;
      for (let t = 0; t < agents.length; t++) {
        const a = agents[(k + t) % agents.length];
        if (full(a)) continue;
        const na = pc.get(a) || 0, nb = best ? pc.get(best) || 0 : Infinity;
        if (na < nb || (na === nb && res.get(a).length < res.get(best).length)) best = a;
      }
      if (best) k = (agents.indexOf(best) + 1) % agents.length;
    }
    if (!best) { pool.push(r); continue; }
    res.get(best).push(r); pc.set(best, (pc.get(best) || 0) + 1);
    if (r.compte && !byCompte.has(r.compte)) byCompte.set(r.compte, best);
  }
  return { res, pool };
}

/* Répartition du jour : si le fichier contient déjà des noms, on ne distribue que les lignes sans nom
   (les autres restent à leur agent) ; « Lignes par agent » = lignes en plus. Sauf si on demande de tout redistribuer. */
function planAttribution(sh) {
  const agents = sh.agents.filter(a => !S.absents.has(a));
  const dejaNommees = sh.rows.some(r => r.agentKey);
  const tout = S.toutRedistribuer || !dejaNommees;
  const rows = tout ? sh.rows : sh.rows.filter(r => !r.agentKey);
  const seed = new Map();
  if (!tout) for (const r of sh.rows) if (r.agentKey && r.compte && agents.includes(r.agentKey) && !seed.has(r.compte)) seed.set(r.compte, r.agentKey);
  return { agents, dejaNommees, tout, rows, ...roundRobin(rows, agents, S.parAgent || 0, seed) };
}

/* ---------- 4. Petits graphiques SVG (aucune librairie) ---------- */

const tip = $('#tip');
function showTip(e, html) { tip.innerHTML = html; tip.hidden = false; const x = Math.min(e.clientX + 14, innerWidth - tip.offsetWidth - 8), y = Math.min(e.clientY + 14, innerHeight - tip.offsetHeight - 8); tip.style.left = x + 'px'; tip.style.top = y + 'px'; }
document.addEventListener('mouseover', e => { const t = e.target.closest('[data-tip]'); if (t) showTip(e, t.dataset.tip); else tip.hidden = true; });
document.addEventListener('mousemove', e => { if (!tip.hidden && e.target.closest('[data-tip]')) showTip(e, e.target.closest('[data-tip]').dataset.tip); });

function legend(series) { return series.length < 2 ? '' : `<div class="legend">${series.map(s => `<span><i style="background:${s.color}"></i>${esc(s.name)}</span>`).join('')}</div>`; }

// Les graphiques sont dessinés à la largeur réelle de leur carte (texte toujours lisible).
const PENDING = [];
function mount(draw, onClick) {
  const id = 'c' + Math.random().toString(36).slice(2, 8);
  PENDING.push({ id, draw, onClick });
  return `<div class="chart" id="${id}"></div>`;
}
function flushCharts() {
  for (const { id, draw, onClick } of PENDING.splice(0)) {
    const el = document.getElementById(id); if (!el) continue;
    el.innerHTML = draw(Math.max(280, el.clientWidth));
    if (onClick) el.addEventListener('click', e => { const r = e.target.closest('[data-key]'); r && onClick(r.dataset.key); });
  }
}
let rz; addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => S.sheets.length && render(), 150); });

// Barres horizontales empilées. items: [{label, values:[...], key}]
function hbars(items, series, { labelW = 170, valFmt = v => fmt(v), total = true, suffix = it => '', onClick } = {}) {
  if (!items.length) return '<p class="muted">Aucune donnée.</p>';
  return legend(series) + mount(W => {
    const lw = Math.min(labelW, W * 0.4);
    const rowH = 30, barH = 18, right = 90, plotW = W - lw - right;
    const maxChars = Math.floor(lw / 6.6);
    const max = Math.max(1, ...items.map(it => it.values.reduce((a, b) => a + b, 0)));
    const H = items.length * rowH + 6;
    let g = '';
    items.forEach((it, i) => {
      const y = i * rowH + (rowH - barH) / 2;
      const sum = it.values.reduce((a, b) => a + b, 0);
      let x = lw, segs = '';
      const nz = it.values.map((v, j) => [v, j]).filter(([v]) => v > 0);
      nz.forEach(([v, j], n) => {
        const w = Math.max(2, (v / max) * plotW) - (n < nz.length - 1 ? 2 : 0);
        const last = n === nz.length - 1, r = Math.min(4, w / 2);
        const d = last ? `M${x},${y}h${w - r}a${r},${r} 0 0 1 ${r},${r}v${barH - 2 * r}a${r},${r} 0 0 1 -${r},${r}h-${w - r}z` : `M${x},${y}h${w}v${barH}h-${w}z`;
        const t = `<b>${esc(it.label)}</b>${esc(series[j].name)} : ${valFmt(v)}${series.length > 1 ? ` (${pct(v, sum)} %)` : ''}`;
        segs += `<path d="${d}" fill="${series[j].color}" data-tip="${esc(t)}"/>`;
        x += w + (n < nz.length - 1 ? 2 : 0);
      });
      const lab = it.label.length > maxChars ? it.label.slice(0, maxChars - 1) + '…' : it.label;
      g += `<g class="row"${onClick ? ` data-key="${esc(it.key ?? it.label)}" style="cursor:pointer"` : ''}>
        <rect class="hit" x="0" y="${i * rowH}" width="${W}" height="${rowH}" data-tip="${esc(`<b>${esc(it.label)}</b>` + series.map((s, j) => `${esc(s.name)} : ${valFmt(it.values[j])}`).join('<br>'))}"/>
        <text class="lbl" x="${lw - 10}" y="${y + barH / 2 + 4}" text-anchor="end">${esc(lab)}</text>${segs}
        ${total ? `<text class="val" x="${x + 6}" y="${y + barH / 2 + 4}">${valFmt(sum)}${esc(suffix(it))}</text>` : ''}</g>`;
    });
    return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img"><line class="axis" x1="${lw}" x2="${lw}" y1="0" y2="${H}"/>${g}</svg>`;
  }, onClick);
}

// Barres groupées (2 séries côte à côte) pour comparer
function groupedBars(items, series, { labelW = 140, valFmt = v => fmt(v) } = {}) {
  return legend(series) + mount(W => {
    const barH = 11, gap = 2, rowH = series.length * (barH + gap) + 14, right = 70, plotW = W - labelW - right;
    const max = Math.max(1, ...items.flatMap(it => it.values));
    const H = items.length * rowH;
    let g = '';
    items.forEach((it, i) => {
      const y0 = i * rowH + 7;
      g += `<text class="lbl" x="${labelW - 10}" y="${y0 + (series.length * (barH + gap)) / 2 + 3}" text-anchor="end">${esc(it.label)}</text>`;
      it.values.forEach((v, j) => {
        const y = y0 + j * (barH + gap), w = Math.max(2, (v / max) * plotW), r = Math.min(4, w / 2);
        g += `<path d="M${labelW},${y}h${w - r}a${r},${r} 0 0 1 ${r},${r}v${barH - 2 * r}a${r},${r} 0 0 1 -${r},${r}h-${w - r}z" fill="${series[j].color}" data-tip="${esc(`<b>${esc(it.label)}</b>${esc(series[j].name)} : ${valFmt(v)}`)}"/>
          <text class="val" x="${labelW + w + 6}" y="${y + barH - 1}">${valFmt(v)}</text>`;
      });
    });
    return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img"><line class="axis" x1="${labelW}" x2="${labelW}" y1="0" y2="${H}"/>${g}</svg>`;
  });
}

// Colonnes empilées (par semaine / mois)
function columns(items, series) {
  if (!items.length) return '<p class="muted">Aucune date.</p>';
  return legend(series) + mount(W => {
    const H = 220, left = 40, bottom = 30, plotW = W - left - 6, plotH = H - bottom - 8;
    const max = Math.max(1, ...items.map(it => it.values.reduce((a, b) => a + b, 0)));
    const step = niceStep(max), top = Math.ceil(max / step) * step;
    const slot = plotW / items.length, bw = Math.min(24, slot * 0.7);
    let g = '';
    for (let v = 0; v <= top; v += step) { const y = 8 + plotH - (v / top) * plotH; g += `<line class="gridl" x1="${left}" x2="${W}" y1="${y}" y2="${y}"/><text x="${left - 6}" y="${y + 4}" text-anchor="end" class="val">${fmt(v)}</text>`; }
    const every = Math.ceil(items.length / Math.max(2, Math.floor(plotW / 60)));
    items.forEach((it, i) => {
      const x = left + i * slot + (slot - bw) / 2; let y = 8 + plotH;
      const sum = it.values.reduce((a, b) => a + b, 0);
      const nz = it.values.map((v, j) => [v, j]).filter(([v]) => v > 0);
      nz.forEach(([v, j], n) => {
        const h = (v / top) * plotH - (n < nz.length - 1 ? 2 : 0), last = n === nz.length - 1, r = Math.min(4, h / 2, bw / 2);
        y -= h;
        const d = last ? `M${x},${y + h}v-${h - r}a${r},${r} 0 0 1 ${r},-${r}h${bw - 2 * r}a${r},${r} 0 0 1 ${r},${r}v${h - r}z` : `M${x},${y}h${bw}v${h}h-${bw}z`;
        g += `<path d="${d}" fill="${series[j].color}" data-tip="${esc(`<b>${esc(it.label)}</b>${esc(series[j].name)} : ${fmt(v)} / ${fmt(sum)}`)}"/>`;
        y -= n < nz.length - 1 ? 2 : 0;
      });
      if (i % every === 0) g += `<text x="${x + bw / 2}" y="${H - bottom + 16}" text-anchor="middle">${esc(it.label)}</text>`;
    });
    g += `<line class="axis" x1="${left}" x2="${W}" y1="${8 + plotH}" y2="${8 + plotH}"/>`;
    return `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img">${g}</svg>`;
  });
}
function niceStep(max) { const raw = max / 4, p = 10 ** Math.floor(Math.log10(raw)), n = raw / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 5 ? 5 : 10) * p; }

const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const SER = () => ({
  oui: { name: 'Traité', color: css('--s1') }, non: { name: 'Non traité', color: css('--s2') },
  autre: { name: 'Autre', color: css('--s7') }, vide: { name: 'Non renseigné', color: css('--none') },
});
const STATS_ORDER = ['oui', 'non', 'autre', 'vide'];
function byStatut(groups) {
  const ser = SER(), used = STATS_ORDER.filter(s => [...groups.values()].some(l => l.some(r => r.statut === s)));
  const items = [...groups].map(([label, l]) => ({ label: label || '(vide)', key: label, values: used.map(s => count(l, r => r.statut === s)), n: l.length }));
  return { items, series: used.map(s => ser[s]) };
}

/* ---------- 5. Vues ---------- */

function render() {
  $('#fname').textContent = S.fileName;
  $('#sheets').innerHTML = S.sheets.map((s, i) => `<button data-i="${i}" class="${i === S.cur ? 'on' : ''}" ${s.ok ? '' : 'disabled'} title="${esc(s.ok ? s.rows.length + ' lignes' : s.why)}">${esc(s.name)}${s.ok ? ` · ${s.rows.length}` : ''}</button>`).join('');
  $('#sheets').onclick = e => { const b = e.target.closest('button'); if (!b || b.disabled) return; S.cur = +b.dataset.i; S.dashAgent = null; S.dashDr = null; S.absents = new Set(); S.parAgent = 0; S.toutRedistribuer = false; S.agentFilter = ''; S.motif = null; render(); };
  document.querySelectorAll('#tabs button').forEach(b => b.classList.toggle('on', b.dataset.tab === S.tab));
  const sh = cur();
  if (!sh || !sh.ok) { $('#view').innerHTML = '<p class="muted">Aucun onglet exploitable dans ce fichier.</p>'; return; }
  $('#tabs [data-tab="equipe2"] .t2').textContent = sh.t2Label === 'équipe 2' ? 'Équipe 2' : sh.t2Label;
  $('#refdate').value = isoDay(sh.ref);
  const T = { preparer: 'Préparer le fichier', dash: 'Tableau de bord', equite: 'Répartition', equipe2: 'Envoyées à ' + sh.t2Label, lignes: 'Lignes', qualite: 'Colonnes et contrôles' };
  $('#view').innerHTML = `<div class="title"><h2>${esc(T[S.tab])}</h2><span class="muted">${esc(sh.name)}</span></div>` +
    ({ preparer: viewPreparer, dash: viewDash, equite: viewEquite, equipe2: viewEquipe2, lignes: viewLignes, qualite: viewQualite })[S.tab](sh);
  flushCharts();
  bindView(sh);
}
const isoWeek = d => { const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate())); const n = t.getUTCDay() || 7; t.setUTCDate(t.getUTCDate() + 4 - n); return Math.ceil(((t - Date.UTC(t.getUTCFullYear(), 0, 1)) / DAY + 1) / 7); };
const isoDay = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
const frDay = d => d ? d.toLocaleDateString('fr-FR') : '';
const short = s => s.replace(/^\d+\s*-\s*/, '');
// Site principal = le plus fréquent sur tous les onglets du fichier
const homeSite = () => { const n = new Map(); for (const s of S.sheets) if (s.ok) for (const r of s.rows) if (r.dr) n.set(r.dr, (n.get(r.dr) || 0) + 1); return [...n].sort((a, b) => b[1] - a[1])[0]?.[0] || ''; };

function kpi(l, v, d = '', hero = false) { return `<div class="kpi${hero ? ' hero' : ''}"><div class="l">${l}</div><div class="v">${v}</div>${d ? `<div class="d">${d}</div>` : ''}</div>`; }

/* ---------- Tableau de bord : les agents d'abord, puis où on travaille (DR / unités) ---------- */
const AGE_BUCKETS = [['< 30 j', 0, 30], ['30-60 j', 30, 60], ['60-90 j', 60, 90], ['90-180 j', 90, 180], ['> 180 j', 180, Infinity]];
const ageBucket = a => AGE_BUCKETS.findIndex(([, lo, hi]) => a >= lo && a < hi);

// Chiffres d'un agent + score « meilleur du fichier »
function agentStats(sh) {
  const by = groupBy(sh.rows.filter(r => r.agentKey), r => r.agentKey);
  const list = [...by].map(([k, l]) => {
    const tr = l.filter(r => r.statut === 'oui'), ages = l.map(r => r.age).filter(a => a != null), agesT = tr.map(r => r.age).filter(a => a != null);
    return { k, name: sh.agentName[k], l, n: l.length, ok: tr.length, ko: count(l, r => r.statut === 'non'), nr: count(l, r => r.statut === 'vide' || r.statut === 'autre'),
      t2: count(l, r => r.t2), age: ages.length ? mean(ages) : 0, ageT: agesT.length ? mean(agesT) : 0, max: ages.length ? Math.max(...ages) : 0 };
  });
  // Score sur 100 : 40 % volume traité, 35 % taux de traitement, 25 % ancienneté des affaires traitées (les plus dures)
  const mOk = Math.max(1, ...list.map(a => a.ok)), mAge = Math.max(1, ...list.map(a => a.ageT));
  list.forEach(a => { a.rate = a.n ? a.ok / a.n : 0; a.score = Math.round(100 * (0.40 * a.ok / mOk + 0.35 * a.rate + 0.25 * a.ageT / mAge)); });
  return list;
}

function viewDash(sh) {
  const R = sh.rows, n = R.length, T2 = esc(sh.t2Label);
  const oui = count(R, r => r.statut === 'oui'), reste = n - oui, vide = count(R, r => r.statut === 'vide');
  const envoyees = R.filter(r => r.t2), t2Faites = count(envoyees, r => r.t2Stat === 'traite');
  const ages = R.map(r => r.age).filter(a => a != null);
  const A = agentStats(sh), RANK = [...A].sort((a, b) => b.score - a.score || b.ok - a.ok);
  const sel = S.dashAgent && A.find(a => a.k === S.dashAgent);
  const best = k => A.length ? [...A].sort((a, b) => b[k] - a[k])[0] : null;
  const color = a => nameColor(a.name)[0];
  // petite carte : un graphique en barres par critère, mêmes agents dans le même ordre partout
  const mini = (title, key, show, max, extra = '') => `<div class="card mini-card"><h3>${title}</h3>${RANK.map(a => `<div class="mini-row clk ${sel && sel.k === a.k ? 'sel' : ''}" data-dagent="${esc(a.k)}"><span>${nameChip(a.name)}</span><b>${show(a[key])}</b><span class="mini-b"><i style="width:${max ? Math.min(100, 100 * a[key] / max) : 0}%;background:${color(a)}"></i></span></div>`).join('')}${extra}</div>`;
  const howScore = `<details class="how"><summary>Comment est calculé le score ?</summary><p>Sur 100 : <b>40</b> pour le nombre d'affaires traitées, <b>35</b> pour le % traité, <b>25</b> pour l'ancienneté des affaires traitées. Chaque critère est comparé au meilleur de l'équipe.</p></details>`;
  const rec = (icon, label, a, val) => a ? `<div class="rec"><span class="rec-i">${icon}</span><span class="rec-l">${label}</span>${nameChip(a.name)}<b>${val}</b></div>` : '';
  // grille agents × ancienneté : case = nombre d'affaires (% traité), plus c'est foncé, plus c'est traité
  const grid = () => `<table class="agt heat"><thead><tr><th>Agent</th>${AGE_BUCKETS.map(([l]) => `<th class="c">${l}</th>`).join('')}<th class="c">Sans date</th><th class="n">Total</th><th class="n">Non traitées</th><th class="n">Non rens.</th><th class="n">Score</th></tr></thead><tbody>
    ${RANK.map(a => { const cells = [...AGE_BUCKETS.map((x, i) => a.l.filter(r => r.age != null && ageBucket(r.age) === i)), a.l.filter(r => r.age == null)];
      const row = `<tr class="clk ${sel && sel.k === a.k ? 'sel' : ''}" data-dagent="${esc(a.k)}"><td>${nameChip(a.name)}</td>${cells.map(l => { if (!l.length) return '<td class="c"></td>'; const p = count(l, r => r.statut === 'oui') / l.length;
        return `<td class="c hc" style="background:rgba(91,95,199,${(0.12 + p * 0.68).toFixed(2)});color:${p > 0.55 ? '#fff' : 'inherit'}" title="${l.length} affaires, ${Math.round(100 * p)} % traitées"><b>${l.length}</b> <small>(${Math.round(100 * p)} %)</small></td>`; }).join('')}
        <td class="n"><b>${a.n}</b></td><td class="n">${a.ko || ''}</td><td class="n">${a.nr || ''}</td><td class="n"><span class="score">${a.score}</span></td></tr>`;
      if (!(sel && sel.k === a.k)) return row;
      const old = a.l.filter(r => r.statut !== 'oui' && r.age != null).sort((x, y) => y.age - x.age).slice(0, 8);
      return row + `<tr class="sub-old"><td colspan="${AGE_BUCKETS.length + 6}"><b>Ses plus anciennes pas encore traitées :</b> ${old.length ? old.map(r => `<span class="old-chip" title="${esc(r.comment)}">${r.age} j · ${esc(short(r.dr) || '—')}${r.affaire ? ' · ' + esc(r.affaire) : ''}</span>`).join('') : '<span class="muted">aucune 👍</span>'}</td></tr>`; }).join('')}
  </tbody></table>`;
  return `
  <div class="kpis">
    ${kpi('Lignes', fmt(n), esc(groupBy(R, r => r.circuit).size > 1 ? groupBy(R, r => r.circuit).size + ' circuits' : [...groupBy(R, r => r.circuit).keys()][0]), true)}
    ${kpi('Traitées', pct(oui, n) + ' %', `${fmt(oui)} lignes`)}
    ${kpi('Reste à faire', fmt(reste), vide ? `dont ${fmt(vide)} sans statut` : '')}
    ${envoyees.length ? kpi('Envoyées à ' + T2, fmt(envoyees.length), t2Faites === envoyees.length ? 'toutes traitées' : `${fmt(t2Faites)} traitées`)
      : kpi('Ancienneté médiane', ages.length ? Math.round(median(ages)) + ' j' : '—', ages.length ? `la plus ancienne : ${Math.max(...ages)} j` : '')}
  </div>
  ${A.length ? `<div class="minis">
    ${mini('📦 Affaires attribuées', 'n', v => v, Math.max(...A.map(a => a.n)))}
    ${mini('✅ Affaires traitées', 'ok', v => v, Math.max(...A.map(a => a.n)))}
    ${mini('📈 % traité', 'rate', v => Math.round(100 * v) + ' %', 1)}
    ${mini('⏳ Ancienneté moyenne', 'age', v => Math.round(v) + ' j', Math.max(...A.map(a => a.age)))}
    ${mini('🏆 Score', 'score', v => v, 100, howScore)}
  </div>
  <div class="grid">
    <div class="card w12"><div class="ph"><span class="num">👥</span>Agents par ancienneté<span class="sum">case = affaires (% traité) · clique un agent</span></div><div class="tbl-wrap" style="max-height:none">${grid()}</div></div>
    <div class="card w6"><div class="ph"><span class="num">📋</span>Récap</div><div class="recs">
      ${rec('🏆', 'Meilleur du fichier', RANK[0], RANK[0].score + '/100')}
      ${rec('📦', 'Le plus d’affaires', best('n'), best('n')?.n)}
      ${rec('✅', 'Le plus de traitées', best('ok'), best('ok')?.ok)}
      ${rec('⏳', 'Traite les plus anciennes', best('ageT'), Math.round(best('ageT')?.ageT || 0) + ' j')}
      ${rec('✕', 'Le plus de non traitées', best('ko'), best('ko')?.ko)}
      ${A.some(a => a.t2) ? rec('↗', 'Envoie le plus à ' + T2, best('t2'), best('t2')?.t2) : ''}
    </div></div>
    ${drBlock(sh).replace('<div class="card w12">', '<div class="card w6 dr-compact">')}
  </div>` : `<div class="grid">${drBlock(sh)}</div>`}
  <details class="more" id="more" ${S.more ? 'open' : ''}><summary>Plus de statistiques</summary>${S.more ? `<div class="inner">${dashMore(sh)}</div>` : ''}</details>`;
}

// Où on travaille : DR en anneau, puis le détail des unités de la DR choisie
const DR_COLORS = ['#5b5fc7', '#e37d39', '#1baf7a', '#c239b3', '#2a78d6', '#b38600', '#038387', '#d13438'];
function drBlock(sh) {
  const R = sh.rows.filter(r => r.dr); if (!R.length) return '';
  const drs = [...groupBy(R, r => r.dr)].sort((a, b) => b[1].length - a[1].length);
  const pick = drs.find(([d]) => d === S.dashDr) || drs[0];
  const n = R.length, ox = 90, oy = 90, ro = 80, ri = 50;
  let a0 = -Math.PI / 2, arcs = '';
  drs.forEach(([d, l], i) => {
    const frac = l.length / n, a1 = a0 + frac * 2 * Math.PI, big = a1 - a0 > Math.PI ? 1 : 0, gap = drs.length > 1 ? 0.012 : 0;
    const p = (r, a) => `${(ox + r * Math.cos(a)).toFixed(2)},${(oy + r * Math.sin(a)).toFixed(2)}`;
    const s0 = a0 + gap, s1 = Math.max(s0, a1 - gap);
    const path = frac >= 0.9999 ? `M${ox - ro},${oy}a${ro},${ro} 0 1,0 ${2 * ro},0a${ro},${ro} 0 1,0 ${-2 * ro},0M${ox - ri},${oy}a${ri},${ri} 0 1,1 ${2 * ri},0a${ri},${ri} 0 1,1 ${-2 * ri},0`
      : `M${p(ro, s0)}A${ro},${ro} 0 ${big} 1 ${p(ro, s1)}L${p(ri, s1)}A${ri},${ri} 0 ${big} 0 ${p(ri, s0)}Z`;
    arcs += `<path d="${path}" fill="${DR_COLORS[i % DR_COLORS.length]}" fill-rule="evenodd" class="arc ${d === pick[0] ? 'on' : ''}" data-ddr="${esc(d)}" data-tip="${esc(`<b>${esc(short(d))}</b>${l.length} lignes · ${pct(l.length, n)} %`)}"/>`;
    a0 = a1;
  });
  const units = [...groupBy(pick[1], r => r.unite || '—')].sort((a, b) => b[1].length - a[1].length);
  const ci = drs.findIndex(([d]) => d === pick[0]), col = DR_COLORS[ci % DR_COLORS.length];
  return `<div class="card w12"><div class="ph"><span class="num">📍</span>Où on travaille<span class="sum">clique une DR pour voir ses unités</span></div>
    <div class="dr-wrap">
      <svg viewBox="0 0 180 180" class="donut" role="img">${arcs}<text x="90" y="86" text-anchor="middle" class="dn-v">${pct(pick[1].length, n)} %</text><text x="90" y="104" text-anchor="middle" class="dn-l">${esc(short(pick[0])).slice(0, 18)}</text></svg>
      <div class="dr-legend">${drs.map(([d, l], i) => `<button type="button" class="dr-item ${d === pick[0] ? 'on' : ''}" data-ddr="${esc(d)}"><i style="background:${DR_COLORS[i % DR_COLORS.length]}"></i><span>${esc(short(d))}</span><b>${pct(l.length, n)} %</b><span class="muted">${l.length}</span></button>`).join('')}</div>
      <div class="dr-units"><h3 style="margin:0 0 10px">${esc(short(pick[0]))} : ${pct(pick[1].length, n)} % des lignes</h3>
        ${units.map(([u, l]) => `<div class="mini-row"><span>${esc(u)}</span><span><b>${pct(l.length, n)} %</b> <span class="muted">du total · ${pct(l.length, pick[1].length)} % de la DR · ${l.length}</span></span><span class="mini-b"><i style="width:${100 * l.length / pick[1].length}%;background:${col}"></i></span></div>`).join('')}
      </div>
    </div></div>`;
}

function motifList(items) {
  if (!items.length) return '<p class="muted">Aucun commentaire.</p>';
  const max = Math.max(...items.map(x => x[1]));
  return `<div class="mlist">${items.map(([t, n]) => `<button type="button" class="mrow" data-motif="${esc(t)}" title="${esc(t)}"><span><span class="t">${esc(t)}</span><span class="track"><span style="width:${(100 * n / max).toFixed(1)}%"></span></span></span><span class="n">${fmt(n)}</span></button>`).join('')}</div>`;
}

function dashMore(sh) {
  const R = sh.rows, n = R.length;
  const deja = count(R, r => r.deja);
  const ages = R.map(r => r.age).filter(a => a != null);
  const home = homeSite(), ailleurs = count(R, r => r.dr && r.dr !== home);
  const doublons = [...groupBy(R.filter(r => r.compte), r => r.compte).values()].filter(l => l.length > 1);
  const drs = byStatut(groupBy(R, r => r.dr)); drs.items.sort((a, b) => b.n - a.n);
  const unites = byStatut(groupBy(R, r => r.unite)); unites.items.sort((a, b) => b.n - a.n);
  const cats = byStatut(groupBy(R, r => r.categorie)); cats.items.sort((a, b) => b.n - a.n);
  const dated = R.filter(r => r.date).sort((a, b) => a.date - b.date);
  const weekly = dated.length && (dated[dated.length - 1].date - dated[0].date) / DAY <= 120;
  const timeline = byStatut(groupBy(dated, r => weekly ? 'S' + isoWeek(r.date) : r.date.toLocaleDateString('fr-FR', { month: 'short', year: '2-digit' })));
  return `
  <div class="kpis">
    ${kpi('Déjà clôturées avant ouverture', fmt(deja), `${pct(deja, n)} % : traitées par un autre agent`)}
    ${ages.length ? kpi('Ancienneté médiane', Math.round(median(ages)) + ' j', `la plus ancienne : ${Math.max(...ages)} j`) : ''}
    ${home && R.some(r => r.dr) ? kpi('Hors site principal', pct(ailleurs, n) + ' %', `${fmt(ailleurs)} lignes hors « ${esc(short(home))} »`) : ''}
    ${R.some(r => r.compte) ? kpi('Comptes en doublon', fmt(doublons.length), `${fmt(doublons.reduce((s, l) => s + l.length, 0))} lignes`) : ''}
  </div>
  <div class="grid">
    ${timeline.items.length > 1 ? `<div class="card w12"><h3>Lignes par ${weekly ? 'semaine' : 'mois'} de génération</h3><p class="hint">Plus c'est à gauche, plus c'est ancien.</p>${columns(timeline.items, timeline.series)}</div>` : ''}
    ${cats.items.length > 1 ? `<div class="card"><h3>Par catégorie</h3>${hbars(cats.items, cats.series, { labelW: 110 })}</div>` : ''}
  </div>`;
}

function viewEquite(sh) {
  const agents = sh.agents.filter(a => !S.absents.has(a));
  const defCap = S.parAgent || 0;
  const sim = planAttribution(sh);
  const simRows = agents.map(a => { const l = sim.res.get(a); const ag = l.map(r => r.age).filter(x => x != null); const avant = sim.tout ? 0 : count(sh.rows, r => r.agentKey === a); return `<tr><td>${esc(sh.agentName[a])}</td>${sim.tout ? '' : `<td class="n">${fmt(avant)}</td>`}<td class="n">${fmt(l.length)}</td><td class="n">${ag.length ? Math.round(mean(ag)) + ' j' : '—'}</td></tr>`; }).join('');

  // Comparaison d'équité : ancienneté moyenne réelle vs 1-1-1-1
  const R = sh.rows.filter(r => r.age != null && r.agentKey);
  const rr = roundRobin(R, agents, 0);
  const items = agents.map(a => ({ label: sh.agentName[a], values: [mean(R.filter(r => r.agentKey === a).map(r => r.age)), mean(rr.res.get(a).map(r => r.age))] }));
  const real = items.map(i => i.values[0]).filter(Boolean), simu = items.map(i => i.values[1]);
  const gapReal = real.length ? Math.max(...real) - Math.min(...real) : 0, gapSim = simu.length ? Math.max(...simu) - Math.min(...simu) : 0;
  const ser = [{ name: 'Répartition actuelle', color: css('--s2') }, { name: 'Répartition 1-1-1-1', color: css('--s1') }];

  return `
  <div class="grid">
    <div class="card w12"><h3>Répartir les lignes</h3>
      <p class="hint">Une ligne chacun à tour de rôle, de la plus ancienne à la plus récente. Le reste reste non attribué (agent vide).
        ${sim.dejaNommees && !sim.tout ? ` Ce fichier a déjà des noms : seules les <strong>${fmt(sim.rows.length)} lignes sans nom</strong> sont distribuées.` : ''}</p>
      <div class="tools">
        <div class="chips">${sh.agents.map(a => `<label class="chip"><input type="checkbox" data-abs="${esc(a)}" ${S.absents.has(a) ? '' : 'checked'}>${esc(sh.agentName[a])}</label>`).join('')}</div>
      </div>
      <div class="tools">
        <label>Lignes par agent <input id="cap" type="number" min="0" max="${sim.rows.length}" value="${defCap}" style="width:80px"></label>
        <span class="muted">0 = tout répartir</span>
        ${sim.dejaNommees ? `<label><input id="tout" type="checkbox" ${S.toutRedistribuer ? 'checked' : ''}> Tout redistribuer</label>` : ''}
        <span style="flex:1"></span>
        <span class="muted">${fmt(sim.rows.length - sim.pool.length)} distribuées · ${fmt(sim.pool.length)} non attribuées</span>
        <button id="dl" class="btn" type="button">Télécharger l'Excel</button>
      </div>
      <div class="tbl-wrap"><table><thead><tr><th>Agent</th>${sim.tout ? '' : '<th class="n">Déjà attribuées</th>'}<th class="n">${sim.tout ? 'Lignes' : 'Nouvelles'}</th><th class="n">Ancienneté moyenne</th></tr></thead><tbody>${simRows}</tbody></table></div>
      <p class="muted">Le fichier téléchargé garde tous les onglets et toutes les lignes, avec la colonne agent remplie. Les couleurs ne sont pas conservées.</p>
    </div>
    ${gapReal ? `<div class="card w12"><h3>Est-ce équitable ?</h3>
      <p class="callout">Écart d'ancienneté entre agents : <strong>${Math.round(gapReal)} jours</strong> aujourd'hui, <strong>${Math.round(gapSim)} jour${Math.round(gapSim) > 1 ? 's' : ''}</strong> en 1-1-1-1.</p>
      ${groupedBars(items, ser, { labelW: 120, valFmt: v => Math.round(v) + ' j' })}</div>` : ''}
  </div>`;
}

function viewEquipe2(sh) {
  const R = sh.rows, T2 = esc(sh.t2Label), G = R.filter(r => r.t2Stat);
  if (!G.length) return `<div class="card w12"><h3>Flux ${T2}</h3><p class="muted">Aucune ligne surlignée en jaune ni envoyée à une 2e équipe dans cet onglet.</p></div>`;
  const st = { propose: "Proposée par l'agent, pas encore surlignée", attente: "Surlignée, en attente d'attribution", encours: 'Attribuée, en cours', traite: `Traitée par ${sh.t2Label}`, retour: "Non traitée : retour à l'agent" };
  const col = { propose: css('--none'), attente: css('--s4'), encours: css('--s7'), traite: css('--s3'), retour: css('--s2') };
  const keys = Object.keys(st).filter(k => G.some(r => r.t2Stat === k));
  const series = keys.map(k => ({ name: st[k], color: col[k] }));
  const mk = groups => [...groups].map(([label, l]) => ({ label: label || '(pas encore attribuée)', values: keys.map(k => count(l, r => r.t2Stat === k)), n: l.length })).sort((a, b) => b.n - a.n);
  const rows = G.map(r => `<tr><td>${nameChip(r.agent)}</td><td>${esc(r.comment)}</td><td>${r.marque ? '🟨' : ''}</td><td>${r.t2Agent ? nameChip(r.t2Agent, true) : '—'}</td><td><span class="pill">${esc(st[r.t2Stat])}</span></td><td>${esc(r.t2Comment)}</td></tr>`).join('');
  return `
  <div class="kpis">
    ${kpi('Envoyées', fmt(count(G, r => r.t2)), `${pct(count(G, r => r.t2), R.length)} % du fichier`, true)}
    ${keys.map(k => kpi(st[k], fmt(count(G, r => r.t2Stat === k)))).join('')}
  </div>
  <div class="grid">
    <div class="card"><h3>Par agent ${T2}</h3>${hbars(mk(groupBy(G, r => r.t2Agent)), series)}</div>
    <div class="card"><h3>Envoyées par agent</h3>${hbars(mk(groupBy(G, r => r.agent)), series)}</div>
    <div class="card w12"><h3>Détail des échanges</h3>
      <div class="tbl-wrap"><table><thead><tr><th>Agent</th><th>Pourquoi non traité</th><th>Jaune</th><th>Agent ${T2}</th><th>Statut</th><th>Retour ${T2}</th></tr></thead><tbody>${rows}</tbody></table></div></div>
  </div>`;
}

function viewLignes(sh) {
  let L = sh.rows;
  if (S.agentFilter) L = L.filter(r => r.agentKey === S.agentFilter);
  if (S.statutFilter) L = L.filter(r => r.statut === S.statutFilter);
  if (S.motif) L = L.filter(r => r.motif === S.motif);
  if (S.search) { const q = norm(S.search); L = L.filter(r => norm([r.affaire, r.compte, r.comment, r.t2Comment, r.dr, r.unite].join(' ')).includes(q)); }
  L = [...L].sort((a, b) => (a.date ?? Infinity) - (b.date ?? Infinity));
  const me = S.agentFilter ? sh.agentName[S.agentFilter] : null;
  const mine = S.agentFilter ? sh.rows.filter(r => r.agentKey === S.agentFilter) : null;
  const shown = L.slice(0, 500);
  return `
  ${me ? `<div class="kpis">${kpi('Mes lignes', fmt(mine.length), esc(me), true)}${kpi('Traitées', pct(count(mine, r => r.statut === 'oui'), mine.length) + ' %')}${kpi('Reste à faire', fmt(count(mine, r => r.statut !== 'oui')))}${kpi('Équipe', pct(count(sh.rows, r => r.statut === 'oui'), sh.rows.length) + ' %', 'traité par toute l\'équipe')}</div>` : ''}
  <div class="card w12">
    <div class="tools">
      <label>Agent <select id="fAgent"><option value="">Tous</option>${sh.agents.map(a => `<option value="${esc(a)}" ${a === S.agentFilter ? 'selected' : ''}>${esc(sh.agentName[a])}</option>`).join('')}</select></label>
      <label>Statut <select id="fStatut"><option value="">Tous</option>${STATS_ORDER.map(s => `<option value="${s}" ${s === S.statutFilter ? 'selected' : ''}>${STATUT_LABEL[s]}</option>`).join('')}</select></label>
      ${S.motif ? `<span class="chip">Motif : ${esc(S.motif)} <button id="clrMotif" class="linkbtn" type="button">✕</button></span>` : ''}
      <label><input id="fSearch" type="search" placeholder="Rechercher (n°, commentaire…)" value="${esc(S.search)}" style="width:240px"></label>
      <span class="muted">${fmt(L.length)} ligne${L.length > 1 ? 's' : ''}${L.length > 500 ? ' (500 premières affichées)' : ''} · triées de la plus ancienne à la plus récente</span>
    </div>
    <div class="tbl-wrap"><table><thead><tr><th>Date</th><th class="n">Anc.</th><th>N° affaire</th><th>N° compte</th><th>Site / unité</th><th>Agent</th><th>Statut</th><th>Commentaire</th><th>${esc(sh.t2Label)}</th></tr></thead>
    <tbody>${shown.map(r => `<tr><td class="num">${frDay(r.date)}</td><td class="n">${r.age ?? ''}</td><td class="num">${esc(r.affaire)}</td><td class="num">${esc(r.compte)}</td><td>${esc(short(r.dr))}<br><span class="muted">${esc(r.unite)}</span></td><td>${r.agent ? nameChip(r.agent) : '—'}</td><td><span class="pill ${r.statut}" title="${esc(r.traiteRaw)}">${STATUT_LABEL[r.statut]}</span></td><td>${esc(r.comment)}</td><td>${r.t2Stat ? (r.marque ? '🟨 ' : '') + (r.t2Agent ? nameChip(r.t2Agent, true) : (r.t2Stat === 'propose' ? 'proposée' : 'à attribuer')) : ''}</td></tr>`).join('')}</tbody></table></div>
  </div>`;
}

function viewQualite(sh) {
  const R = sh.rows, T2 = esc(sh.t2Label);
  const variants = [...groupBy(R.filter(r => r.traiteRaw), r => r.traiteRaw)].map(([k, l]) => [k, l.length]).sort((a, b) => b[1] - a[1]);
  const comptes = [...groupBy(R.filter(r => r.compte), r => r.compte)].filter(([, l]) => l.length > 1);
  const multiAgents = comptes.filter(([, l]) => new Set(l.map(r => r.agentKey).filter(Boolean)).size > 1);
  const fields = FIELDS.map(([k, label]) => `<tr><td>${label}</td><td><select data-map="${k}"><option value="-1">— (absente)</option>${sh.headers.map((h, i) => `<option value="${i}" ${sh.map[k] === i ? 'selected' : ''}>${esc(colLetter(sh.c0 + i))} · ${esc(h || '(sans titre)')}</option>`).join('')}</select></td><td>${sh.map[k] >= 0 ? '<span class="ok">✓ trouvée</span>' : '<span class="muted">non trouvée</span>'}</td></tr>`).join('');
  const nb = f => fmt(count(R, f));
  return `
  <div class="grid">
    <div class="card w8"><h3>Colonnes reconnues automatiquement</h3><p class="hint">L'outil devine les colonnes même si les titres changent d'un fichier à l'autre. Si une colonne est mal reconnue, corrige-la ici : tous les calculs se mettent à jour. 2e équipe détectée : <strong>${T2}</strong>.</p>
      <div class="tbl-wrap"><table><thead><tr><th>Information</th><th>Colonne du fichier</th><th></th></tr></thead><tbody>${fields}</tbody></table></div></div>
    <div class="card w4"><h3>Contrôles</h3>
      <ul class="q">
        <li>Titres en ligne ${sh.headerRow}, ${fmt(R.length)} lignes lues</li>
        <li>${fmt(sh.empty)} ligne${sh.empty > 1 ? 's' : ''} vide${sh.empty > 1 ? 's' : ''} ignorée${sh.empty > 1 ? 's' : ''}</li>
        <li class="${count(R, r => !r.agentKey) ? 'warn' : ''}">${nb(r => !r.agentKey)} ligne(s) sans agent</li>
        <li class="${count(R, r => r.statut === 'vide') ? 'warn' : ''}">${nb(r => r.statut === 'vide')} ligne(s) sans statut</li>
        <li>${nb(r => r.date == null)} ligne(s) sans date</li>
        <li>${nb(r => r.marque)} ligne(s) surlignée(s) en jaune</li>
        <li class="${count(R, r => r.t2Stat === 'attente') ? 'warn' : ''}">${nb(r => r.t2Stat === 'attente')} surlignée(s) mais pas encore attribuée(s) chez ${T2}</li>
        <li class="${count(R, r => r.t2Stat === 'propose') ? 'warn' : ''}">${nb(r => r.t2Stat === 'propose')} proposée(s) pour ${T2} mais pas surlignée(s)</li>
        <li>${nb(r => r.t2AgentKey && !r.marque)} attribuée(s) chez ${T2} sans surlignage</li>
        <li>${fmt(comptes.length)} compte(s) avec plusieurs lignes, dont <strong>${fmt(multiAgents.length)}</strong> répartis sur plusieurs agents</li>
        <li>Prénoms fusionnés : ${sh.agents.map(a => esc(sh.agentName[a])).join(', ')}</li>
      </ul></div>
    <div class="card"><h3>${variants.length} façons d'écrire le statut</h3><p class="hint">Toutes comprises par l'outil.</p>
      <div class="tbl-wrap" style="max-height:320px"><table><thead><tr><th>Écrit dans le fichier</th><th>Compris comme</th><th class="n">Lignes</th></tr></thead><tbody>${variants.map(([k, n]) => `<tr><td><code>${esc(JSON.stringify(k))}</code></td><td><span class="pill ${statut(k)}">${STATUT_LABEL[statut(k)]}</span></td><td class="n">${n}</td></tr>`).join('')}</tbody></table></div></div>
    <div class="card"><h3>Comptes en doublon</h3><p class="hint">Même compte, plusieurs lignes. La répartition les donne au même agent.</p>
      <div class="tbl-wrap" style="max-height:320px"><table><thead><tr><th>N° compte</th><th class="n">Lignes</th><th>Agents</th></tr></thead><tbody>${comptes.sort((a, b) => b[1].length - a[1].length).map(([c, l]) => `<tr><td class="num">${esc(c)}</td><td class="n">${l.length}</td><td>${esc([...new Set(l.map(r => r.agent || '—'))].join(', '))}</td></tr>`).join('') || '<tr><td colspan="3" class="muted">Aucun</td></tr>'}</tbody></table></div></div>
  </div>`;
}
const colLetter = i => { let s = ''; i++; while (i) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - m - 1) / 26; } return s; };

function bindView(sh) {
  document.querySelectorAll('[data-map]').forEach(sel => sel.onchange = () => { sh.map[sel.dataset.map] = +sel.value; compute(sh); render(); });
  document.querySelectorAll('[data-abs]').forEach(cb => cb.onchange = () => { cb.checked ? S.absents.delete(cb.dataset.abs) : S.absents.add(cb.dataset.abs); render(); });
  const cap = $('#cap'); if (cap) cap.onchange = () => { S.parAgent = Math.max(0, +cap.value || 0); render(); };
  const dl = $('#dl'); if (dl) dl.onclick = () => exportAttribution(sh);
  document.querySelectorAll('[data-motif]').forEach(b => b.onclick = () => { S.motif = b.dataset.motif; S.tab = 'lignes'; S.statutFilter = ''; S.agentFilter = ''; render(); });
  const more = $('#more'); if (more) more.ontoggle = () => { if (more.open !== S.more) { S.more = more.open; render(); } };
  document.querySelectorAll('[data-dagent]').forEach(tr => tr.onclick = () => { S.dashAgent = S.dashAgent === tr.dataset.dagent ? null : tr.dataset.dagent; render(); });
  document.querySelectorAll('[data-dclose]').forEach(b => b.onclick = () => { S.dashAgent = null; render(); });
  document.querySelectorAll('[data-dsort]').forEach(t => t.onclick = () => { const k = t.dataset.dsort, cur = S.dashSort || { k: 'score', dir: -1 }; S.dashSort = { k, dir: cur.k === k ? -cur.dir : (k === 'name' ? 1 : -1) }; render(); });
  document.querySelectorAll('[data-ddr]').forEach(el => el.onclick = () => { S.dashDr = el.dataset.ddr; render(); });
  bindPrep();
  const tt = $('#tout'); if (tt) tt.onchange = () => { S.toutRedistribuer = tt.checked; render(); };
  const fa = $('#fAgent'); if (fa) fa.onchange = () => { S.agentFilter = fa.value; render(); };
  const fs = $('#fStatut'); if (fs) fs.onchange = () => { S.statutFilter = fs.value; render(); };
  const cm = $('#clrMotif'); if (cm) cm.onclick = () => { S.motif = null; render(); };
  const q = $('#fSearch'); if (q) { q.oninput = () => { S.search = q.value; clearTimeout(q._t); q._t = setTimeout(() => { render(); const n = $('#fSearch'); n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250); }; }
}

/* ---------- 6. Export : le fichier reçu lui-même, avec la colonne agent remplie ---------- */
function exportAttribution(sh) {
  const { res, pool, rows } = planAttribution(sh);
  const who = new Map();
  for (const [a, l] of res) l.forEach(r => who.set(r.i, sh.agentName[a]));
  const ws = sh.ws, enc = XLSX.utils.encode_cell;
  const c = sh.c0 + (sh.map.agent >= 0 ? sh.map.agent : sh.headers.length); // pas de colonne agent : on en ajoute une à droite
  const saved = new Map(), put = (r, v) => { const k = enc({ r, c }); if (!saved.has(k)) saved.set(k, ws[k]); if (v) ws[k] = { t: 's', v }; else delete ws[k]; };
  const ref = ws['!ref'];
  if (sh.map.agent < 0) { put(sh.headerIdx, 'agents'); const rg = XLSX.utils.decode_range(ref); rg.e.c = Math.max(rg.e.c, c); ws['!ref'] = XLSX.utils.encode_range(rg); }
  rows.forEach(r => put(sh.sheetRow[r.i], who.get(r.i) || '')); // seules les lignes distribuées sont touchées
  const recap = [['Agent', 'Lignes distribuées', 'Ancienneté moyenne (j)', 'Plus ancienne (j)']];
  for (const [a, l] of res) { const ag = l.map(r => r.age).filter(x => x != null); recap.push([sh.agentName[a], l.length, Math.round(mean(ag)), ag.length ? Math.max(...ag) : null]); }
  recap.push(['Non attribuées (agent vide)', pool.length, null, null]);
  const wb = { ...S.wb, SheetNames: [...S.wb.SheetNames], Sheets: { ...S.wb.Sheets } };
  let recapName = 'Répartition'; while (wb.Sheets[recapName]) recapName += '+';
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(recap), recapName);
  try {
    XLSX.writeFile(wb, `${S.fileName.replace(/\.[^.]+$/, '')} - réparti ${isoDay(new Date())}.xlsx`, { cellDates: true });
  } finally { // on remet le fichier en mémoire dans son état d'origine
    for (const [k, v] of saved) { if (v) ws[k] = v; else delete ws[k]; }
    ws['!ref'] = ref;
  }
}

/* ---------- 7. Préparer le fichier : colonnes utiles + répartition + affichages Excel ---------- */
const STD = [['dr', 'DR'], ['unite', 'UNITE'], ['circuit', 'CIRCUIT'], ['date', 'ANCIENNETE'], ['compte', 'COMPTE'], ['affaire', "NUMERO D'AFFAIRE"],
  ['agent', 'AGENT'], ['traite', 'TRAITE'], ['comment', 'COMMENTAIRE']];
const STD2 = ['COMPETENCE', 'AGENT', 'TRAITE', 'COMMENTAIRE'];
const STD_W = { dr: 24, unite: 26, circuit: 34, date: 13, compte: 13, affaire: 18, agent: 14, traite: 13, comment: 48 };
const defaultCols = () => STD.map(([key, header]) => ({ key, header }));
const LISTE_TRAITE = ['oui', 'non', 'déjà traité'];
const xmlEsc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function initPrep() {
  const ok = S.sheets.filter(s => s.ok);
  const label = ok.map(s => s.t2Label).find(l => l !== 'équipe 2') || '';
  const noms = []; // la liste d'agents part vide : la manager tape les prénoms du jour
  // Notre équipe : la valeur la plus fréquente d'une colonne « Compétence » du fichier, s'il y en a une
  const comp = new Map();
  for (const s of ok) { const j = s.headers.findIndex(x => norm(x) === 'competence'); if (j >= 0) s.raw.forEach(r => { const v = String(r[j] ?? '').trim(); if (v) comp.set(v, (comp.get(v) || 0) + 1); }); }
  const ownTeam = [...comp].sort((a, b) => b[1] - a[1])[0]?.[0] || '';
  // Les noms d'équipe sont gardés dans le lien de la page (#equipe=…&autre=…) : mets la page en favori, ils reviennent tout seuls.
  const h = new URLSearchParams(location.hash.slice(1));
  const t2Name = h.has('autre') ? h.get('autre') : label;
  S.prep = {
    agents: noms, same: true, objAll: 20, method: 'rr', tout: false,
    t2: !!t2Name, t2Name, teamName: h.get('equipe') || ownTeam,
    cols: defaultCols(), views: { all: true, agents: true, pot: true, t2: true, team: true }, mode: 'obj', checkbox: true, protect: true, all: ok.length > 1, sort: true,
  };
}
function saveTeamsInLink() {
  const h = new URLSearchParams(); h.set('equipe', S.prep.teamName); h.set('autre', S.prep.t2Name);
  history.replaceState(null, '', '#' + h.toString());
}

const traiteStd = (raw, st, deja) => st === 'oui' ? (deja ? 'déjà traité' : 'oui') : st === 'non' ? 'non' : st === 'vide' ? '' : String(raw).trim();
const excelDay = d => d ? Math.round((Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - Date.UTC(1899, 11, 30)) / DAY) : null;
const prepTeam = () => S.prep.agents.filter(a => a.name.trim()).map(a => ({ name: a.name.trim(), obj: Math.max(0, S.prep.same ? S.prep.objAll : a.obj) || 0 }));

/* Parts équitables : chacun reçoit autant que possible, sans dépasser son objectif.
   Ex. 100 lignes, 5 agents à 24 → 20 chacun (et non 24-24-24-24-4). Le reste éventuel (+1) va aux premiers de la liste. */
function fairQuotas(objs, n) {
  const q = objs.map(() => 0); let left = n, active = objs.map((o, i) => i).filter(i => objs[i] > 0);
  while (left > 0 && active.length) {
    const share = Math.floor(left / active.length);
    if (share === 0) { for (const i of active) { if (!left) break; q[i]++; left--; } break; }
    for (const i of active) { const add = Math.min(share, objs[i] - q[i]); q[i] += add; left -= add; }
    active = active.filter(i => q[i] < objs[i]);
  }
  return q;
}

/* Répartition du fichier préparé.
   1-1-1-1 : une ligne chacun à tour de rôle (les plus anciennes d'abord), un compte en doublon reste au même agent.
   Classique : des blocs consécutifs (les N premières pour le 1er agent, les N suivantes pour le 2e…). */
function dealPrep(sh) {
  const P = S.prep, team = prepTeam().filter(a => P.mode === 'egal' || a.obj > 0);
  const rows = P.sort ? [...sh.rows].sort((a, b) => (a.date ?? Infinity) - (b.date ?? Infinity)) : [...sh.rows];
  const dejaNommees = sh.rows.some(r => r.agentKey), tout = P.tout || !dejaNommees;
  const todo = tout ? rows : rows.filter(r => !r.agentKey);
  const demande = P.mode === 'egal' ? Infinity : team.reduce((t, a) => t + a.obj, 0);
  const quotas = fairQuotas(team.map(a => P.mode === 'egal' ? Infinity : a.obj), todo.length);
  team.forEach((a, i) => { a.quota = quotas[i]; });
  const who = new Map(), n = new Map(team.map(a => [a, 0])), full = a => n.get(a) >= a.quota;
  const give = (r, a) => { who.set(r.i, a.name); n.set(a, n.get(a) + 1); };
  if (P.method === 'bloc') {
    let k = 0;
    for (const r of todo) { while (k < team.length && full(team[k])) k++; if (k >= team.length) break; give(r, team[k]); }
  } else {
    const byCompte = new Map(); let k = 0;
    for (const r of todo) {
      let a = r.compte && byCompte.get(r.compte);
      if (!a || full(a)) {
        a = null;
        for (let t = 0; t < team.length; t++) { const c = team[(k + t) % team.length]; if (!full(c)) { a = c; k = (k + t + 1) % team.length; break; } }
      }
      if (!a) break;
      give(r, a); if (r.compte && !byCompte.has(r.compte)) byCompte.set(r.compte, a);
    }
  }
  return { rows, who, tout, dejaNommees, todo: todo.length, given: who.size, demande, team };
}

// Colonne du fichier source qui alimente une colonne du fichier prêt
const srcIndex = (sh, c) => c.key ? (c.key === 'date' && sh.map.date < 0 ? sh.map.dateStock : sh.map[c.key]) : sh.headers.findIndex(h => norm(h) === norm(c.src));

function cellFor(c, r, sh, d) {
  switch (c.key) {
    case 'dr': return r.dr; case 'unite': return r.unite; case 'circuit': return sh.map.circuit >= 0 ? r.circuit : '';
    case 'date': return r.date; case 'compte': return r.compte; case 'affaire': return r.affaire;
    case 'agent': return d.who.get(r.i) || (d.tout ? '' : r.agent);
    case 'traite': return traiteStd(r.traiteRaw, r.statut, /^deja/.test(norm(r.traiteRaw)));
    case 'comment': return r.comment;
  }
  const j = srcIndex(sh, c), v = j >= 0 ? r.raw[j] : null;
  return v instanceof Date ? toDate(v) : v;
}

function prepRows(sh) {
  const P = S.prep, d = dealPrep(sh);
  return d.rows.map(r => {
    const line = P.cols.map(c => cellFor(c, r, sh, d));
    if (P.t2) { // colonnes de l'autre équipe : vides au départ, sauf si on réinjecte un fichier déjà préparé par l'outil
      const m = sh.map, mine = m.t2Agent >= 0 && /^agent \S/.test(norm(sh.headers[m.t2Agent]));
      if (mine) {
        const fv = m.t2Flag >= 0 ? r.raw[m.t2Flag] : null, yes = fv === true || /^(oui|vrai|true)/.test(norm(fv)), no = fv === false || /^(non|faux|false)/.test(norm(fv));
        const gT = m.t2Traite >= 0 ? r.raw[m.t2Traite] : null;
        line.push(P.checkbox ? yes : yes ? 'oui' : no ? 'non' : null, r.t2Agent || null, gT == null ? null : traiteStd(gT, statut(gT), /^deja/.test(norm(gT))), r.t2Comment || null);
      } else line.push(P.checkbox ? false : null, null, null, null);
    }
    return { line: line.map(v => v === '' ? null : v), jaune: r.marque };
  });
}

function viewPreparer(sh) {
  const P = S.prep, ok = S.sheets.filter(s => s.ok), N = (P.t2Name || 'équipe 2').toUpperCase();
  const d = dealPrep(sh), team = prepTeam();
  const seg = (name, opts, val) => `<span class="seg">${opts.map(([v, l]) => `<label class="${v === val ? 'on' : ''}"><input type="radio" name="${name}" value="${v}" ${v === val ? 'checked' : ''}>${l}</label>`).join('')}</span>`;
  const tg = (attrs, label, on) => `<label class="tg ${on ? 'on' : ''}"><input type="checkbox" ${attrs} ${on ? 'checked' : ''}><i></i>${label}</label>`;
  const row = (attrs, label, sub, on, extra = '') => `<label class="set-row ${on ? 'on' : ''}"><input type="checkbox" ${attrs} ${on ? 'checked' : ''}><span class="set-t">${label}${sub ? `<small>${sub}</small>` : ''}</span>${extra}<i></i></label>`;
  const vchip = (key, icon, label, on) => `<label class="vchip ${on ? 'on' : ''}"><input type="checkbox" data-pview="${key}" ${on ? 'checked' : ''}><span class="vi">${icon}</span>${label}${on ? '<b>✓</b>' : ''}</label>`;
  const given = name => count([...d.who.values()], v => v === name.trim());
  const agentsRows = P.agents.map((a, i) => `<tr>
    <td><span class="av" style="background:${CX_PAL[i % CX_PAL.length][0]}">${esc(a.name.trim().slice(0, 2).toUpperCase())}</span><input type="search" data-pname="${i}" value="${esc(a.name)}" placeholder="Prénom" style="width:150px"></td>
    ${P.mode === 'obj' ? `<td class="n">${P.same ? `<b>${fmt(P.objAll)}</b>` : `<input type="number" min="0" data-pobj="${i}" value="${a.obj}" style="width:70px">`}</td>` : ''}
    <td class="n"><b>${fmt(given(a.name))}</b></td>
    <td class="n"><button type="button" class="linkbtn" data-pdel="${i}" title="Retirer">✕</button></td></tr>`).join('');
  const nonAttr = d.dejaNommees && !d.tout ? count(sh.rows, r => !r.agentKey) - d.given : d.todo - d.given;
  return `
  <div class="kpis prep-kpis">
    ${kpi('Lignes', fmt(sh.rows.length), esc(sh.name), true)}
    ${kpi('Attribuées', fmt(d.given), team.length ? `${team.length} agent${team.length > 1 ? 's' : ''}` : 'aucun agent')}
    ${kpi('Non attribuées', fmt(nonAttr), '')}
    <div class="kpi dl"><button id="pGo" class="btn" type="button">⬇ Télécharger le fichier prêt</button></div>
  </div>
  ${d.dejaNommees && !d.tout && d.todo === 0 ? `<p class="callout">Toutes les lignes ont déjà un agent. Active « Tout redistribuer » pour les répartir à nouveau.</p>` : ''}
  ${P.mode === 'obj' && d.team.length && d.todo > 0 && d.demande > d.todo ? `<p class="callout">${fmt(d.demande)} lignes demandées pour ${fmt(d.todo)} : réparti à parts égales (${[...new Set(d.team.map(a => a.quota))].sort((a, b) => b - a).join(' ou ')} chacun).</p>` : ''}
  <div class="grid">
    <div class="card w7"><div class="ph"><span class="num">1</span>Agents<span class="sum">${seg('pMode', [['obj', 'Objectif par agent'], ['egal', 'Tout, à parts égales']], P.mode)}</span></div>
      ${P.agents.length ? `<div class="tbl-wrap" style="max-height:340px"><table><thead><tr><th>Agent</th>${P.mode === 'obj' ? '<th class="n">Objectif</th>' : ''}<th class="n">Lignes</th><th></th></tr></thead><tbody>${agentsRows}</tbody></table></div>` : '<p class="muted">Aucun agent.</p>'}
      <div class="tools" style="margin-top:12px">
        <input id="pAdd" type="search" placeholder="＋ Prénoms…" style="width:200px"><button id="pAddBtn" class="ghost" type="button">Ajouter</button>
        ${P.mode === 'obj' ? `${tg('id="pSame"', 'Même objectif', P.same)}${P.same ? `<input id="pObjAll" type="number" min="0" value="${P.objAll}" style="width:70px">` : ''}` : ''}
        <span style="flex:1"></span>${seg('pMeth', [['rr', '1-1-1-1'], ['bloc', 'À la suite']], P.method)}
      </div>
      ${d.dejaNommees ? `<div class="tools">${tg('id="pTout"', 'Tout redistribuer', P.tout)}</div>` : ''}
    </div>
    <div class="card w5"><div class="ph"><span class="num">2</span>Équipes</div>
      <label class="team-name"><span>Notre équipe</span><input id="pTeam" type="search" value="${esc(P.teamName)}" placeholder="Nom de ton équipe"></label>
      <label class="team-name"><span>Autre équipe</span><input id="pT2n" type="search" value="${esc(P.t2Name)}" placeholder="aucune"></label>
    </div>
    <div class="card w12"><div class="ph"><span class="num">3</span>Colonnes</div>${prepColsHtml(sh, team, N)}</div>
    <div class="card w12"><div class="ph"><span class="num">4</span>Aperçu du fichier prêt<span class="sum">${fmt(Math.min(15, sh.rows.length))} premières lignes sur ${fmt(sh.rows.length)}</span></div>${prepPreviewHtml(sh, team, N, 15)}</div>
  </div>`;
}

// Nombre d'affichages Excel qui seront créés (même logique que prepWorkbook)
function prepViewCount(sh, d) {
  const P = S.prep, iAgent = P.cols.findIndex(c => c.key === 'agent');
  const noms = new Set(); if (iAgent >= 0) sh.rows.forEach(r => { const a = d.who.get(r.i) || (d.tout ? '' : r.agent); if (a) noms.add(a); });
  const sans = iAgent >= 0 && sh.rows.some(r => !(d.who.get(r.i) || (d.tout ? '' : r.agent)));
  return (P.views.agents ? noms.size : 0) + (P.views.all && P.cols.length ? 1 : 0) + (P.views.pot && sans ? 1 : 0) + (P.t2 && P.views.t2 ? 1 : 0) + (P.t2 && P.views.team ? 1 : 0);
}

// Aperçu façon Excel : les premières lignes du fichier qui sera téléchargé
function prepPreviewHtml(sh, team, N, n) {
  const P = S.prep, rows = prepRows(sh).slice(0, n), iAgent = P.cols.findIndex(c => c.key === 'agent');
  const heads = [...P.cols.map(c => c.header), ...(P.t2 ? STD2.map(h => `${h} ${N}`) : [])];
  const color = name => { const k = team.findIndex(a => norm(a.name) === norm(name)); return k >= 0 ? LINK_COLORS[P.agents.findIndex(a => norm(a.name) === norm(name)) % LINK_COLORS.length] : ''; };
  const show = v => v instanceof Date ? frDay(v) : v === false ? '☐' : v === true ? '☑' : v == null ? '' : String(v);
  return `<div class="xl"><table>
    <tr class="let"><th></th>${heads.map((h, j) => `<th>${colLetter(j)}</th>`).join('')}</tr>
    <tr class="hd"><td class="rn">1</td>${heads.map((h, j) => `<th class="${P.t2 && j >= P.cols.length ? 'g' : ''}">${esc(h)}</th>`).join('')}</tr>
    ${rows.map((x, k) => `<tr${x.jaune ? ' class="jaune"' : ''}><td class="rn">${k + 2}</td>${x.line.map((v, j) => (j === iAgent || (P.t2 && j === P.cols.length + 1)) && v ? `<td>${nameChip(String(v), j !== iAgent)}</td>` : `<td title="${esc(show(v))}">${esc(show(v))}</td>`).join('')}</tr>`).join('')}
  </table></div>`;
}

/* Écran « Colonnes » : en haut la 1re ligne de ton fichier, en bas ce qu'elle devient.
   ✕ pour barrer une colonne ; pour relier, clic en haut puis clic en bas (ou l'inverse).
   Une colonne et sa place portent la même couleur. */
// Couleurs des prénoms : [texte, fond]. L'équipe et l'autre équipe ont des palettes différentes.
const CX_PAL = [['#4f52b2', '#e8ebfa'], ['#b4501a', '#fdebdc'], ['#0e7a55', '#dcf4ea'], ['#a42d97', '#f8e3f5'], ['#1f63b5', '#e0edfb'], ['#b3261e', '#fbe2e0'],
  ['#7a5a00', '#fbf0cc'], ['#036c6f', '#d8f3f3'], ['#5b3fd1', '#ebe6ff'], ['#b0175f', '#fde2ee']];
const T2_PAL = [['#0b6e4f', '#d3f0e2'], ['#8a4b08', '#f6e6cf'], ['#2b5797', '#dae6f5'], ['#8f2d56', '#f5dde8'], ['#4c5b12', '#ebf1d2'], ['#5c3c92', '#e9e1f5'], ['#6b4226', '#efe2d8'], ['#00627a', '#d4eef5']];
const hashName = n => [...norm(n)].reduce((h, c) => (h * 31 + c.charCodeAt(0)) >>> 0, 7);
function nameColor(name, other = false) {
  if (!name) return null;
  const pal = other ? T2_PAL : CX_PAL;
  if (!other && S.prep) { const k = S.prep.agents.findIndex(a => norm(a.name) === norm(name)); if (k >= 0) return pal[k % pal.length]; }
  // sinon : l'ordre des agents du fichier (une couleur différente pour chacun), puis les prénoms inconnus à la suite
  const sh = typeof cur === 'function' ? cur() : null, keys = sh && sh.ok ? (other ? Object.keys(sh.t2Name || {}).sort() : sh.agents) : [];
  const off = other ? 0 : (S.prep ? S.prep.agents.length : 0), j = keys.indexOf(norm(name));
  if (j >= 0) return pal[(off + j) % pal.length];
  return pal[hashName(name) % pal.length];
}
const nameChip = (name, other = false) => { const c = nameColor(name, other); return c ? `<span class="nchip" style="color:${c[0]};background:${c[1]}">${esc(name)}</span>` : ''; };

const LINK_COLORS = ['#5b5fc7', '#e37d39', '#1baf7a', '#c239b3', '#2a78d6', '#d13438', '#b38600', '#038387', '#7a5cff', '#e3008c', '#4f6bed', '#8a6e00'];
function prepColsHtml(sh, team, N) {
  const P = S.prep, pend = S.prepPend, crossed = sh.crossed || (sh.crossed = new Set());
  const srcOf = P.cols.map(c => srcIndex(sh, c)), finOf = j => srcOf.indexOf(j), color = f => LINK_COLORS[f % LINK_COLORS.length];
  const showV = v => v instanceof Date ? frDay(toDate(v)) : v == null ? '' : String(v);
  const r0 = sh.rows[0], d0 = { who: new Map(), tout: false };
  const filled = v => v != null && String(v).trim() !== '';
  const sample = j => { const r = sh.raw.find(x => filled(x[j])); return r ? r[j] : null; };
  const bar = f => `<span class="lk-bar" style="background:${color(f)}"></span>`;
  const top = sh.headers.map((h, j) => { if (crossed.has(j)) return ''; const f = finOf(j), on = pend && pend.t === 's' && pend.i === j;
    return `<div class="lk ${f < 0 ? 'unused' : ''} ${on ? 'pend' : ''}" data-pss="${j}" draggable="true" title="${esc(h)} — glisse-la dans le fichier prêt pour l'ajouter">
      <button type="button" class="lk-x" data-psx="${j}" title="Supprimer">✕</button>
      ${f < 0 ? `<button type="button" class="lk-add" data-psk="${j}" title="Ajouter au fichier prêt">＋</button>` : ''}
      <div class="lk-l">${colLetter(sh.c0 + j)}</div><div class="lk-h">${esc(h || '(sans titre)')}</div><div class="lk-v">${esc(showV(sample(j))) || '<span class="muted">vide</span>'}</div>${f >= 0 ? bar(f) : ''}</div>`; }).join('');
  // aperçu de ton fichier tel quel (colonnes non supprimées), même couleur que la carte quand la colonne va dans le fichier prêt
  const vis = sh.headers.map((h, j) => j).filter(j => !crossed.has(j));
  const srcPreview = `<div class="xl src-xl"><table>
    <tr class="let"><th></th>${vis.map(j => `<th>${colLetter(sh.c0 + j)}</th>`).join('')}</tr>
    <tr class="hd"><td class="rn">${sh.headerRow}</td>${vis.map(j => { const f = finOf(j); return `<th class="${f < 0 ? 'off' : ''}" ${f >= 0 ? `style="box-shadow:inset 0 -3px 0 ${color(f)}"` : ''} title="${esc(sh.headers[j])}"><button type="button" class="src-x" data-psx="${j}" title="Supprimer cette colonne">✕</button>${esc(sh.headers[j] || '(sans titre)')}</th>`; }).join('')}</tr>
    ${sh.raw.slice(0, 8).map((r, k) => `<tr><td class="rn">${sh.sheetRow[k] + 1}</td>${vis.map(j => `<td class="${finOf(j) < 0 ? 'off' : ''}" title="${esc(showV(r[j]))}">${esc(showV(r[j]))}</td>`).join('')}</tr>`).join('')}
  </table></div>`;
  const bottom = P.cols.map((c, f) => { const j = srcOf[f], on = pend && pend.t === 'f' && pend.i === f;
    const ex = j >= 0 ? sh.rows.find(r => filled(r.raw[j])) || r0 : r0;
    const v = c.key === 'agent' && team.length ? 'rempli par la répartition' : showV(ex ? cellFor(c, ex, sh, d0) : '');
    return `<div class="lk fin ${j < 0 ? 'empty' : ''} ${on ? 'pend' : ''}" data-pfc="${f}" title="${esc(c.header)}">
      <button type="button" class="lk-x" data-pfx="${f}" title="Retirer du fichier prêt">✕</button>
      <div class="lk-l">${colLetter(f)}</div><div class="lk-h">${esc(c.header)}</div><div class="lk-v">${j < 0 ? 'à relier' : esc(v) || '<span class="muted">vide</span>'}</div>
      <div class="lk-tools"><button type="button" data-pfm="${f}" data-dir="-1" ${f ? '' : 'disabled'} title="Vers la gauche">◀</button><button type="button" data-pfr="${f}" title="Renommer">✎</button><button type="button" data-pfm="${f}" data-dir="1" ${f < P.cols.length - 1 ? '' : 'disabled'} title="Vers la droite">▶</button></div>${j >= 0 ? bar(f) : ''}</div>`; }).join('')
    + (P.t2 ? STD2.map((h, k) => `<div class="lk fin g"><div class="lk-l">${colLetter(P.cols.length + k)}</div><div class="lk-h">${esc(`${h} ${N}`)}</div><div class="lk-v">vide (autre équipe)</div></div>`).join('') : '');
  const missing = P.cols.filter((c, f) => srcOf[f] < 0).map(c => c.header);
  const tip = pend ? `<p class="tip">${pend.t === 's' ? `« ${esc(sh.headers[pend.i])} » choisie : clique maintenant sa place dans le fichier prêt (en bas).` : `« ${esc(P.cols[pend.i].header)} » choisie : clique maintenant la colonne de ton fichier qui va dedans (en haut).`} <button type="button" class="linkbtn" data-pcancel="1">Annuler</button></p>`
    : missing.length ? `<p class="tip warn">À relier : ${missing.map(esc).join(', ')}. Clique la colonne en bas, puis la bonne colonne en haut — ou ✕ pour la retirer.</p>`
    : '';
  return `${tip}
    <h4 class="lk-t">Ton fichier${crossed.size ? `<button type="button" class="linkbtn" data-prestore="1" style="margin-left:auto;text-transform:none;letter-spacing:0;font-size:13px">↺ Revenir à l'original (${crossed.size} colonne${crossed.size > 1 ? 's' : ''} supprimée${crossed.size > 1 ? 's' : ''})</button>` : ''}</h4><div class="lk-row">${top}</div>${srcPreview}
    <h4 class="lk-t" style="margin-top:18px">Fichier prêt</h4><div class="lk-row" data-pdrop="1">${bottom}</div>
    <div class="tools"><button id="pColReset" class="linkbtn" type="button">Tout remettre comme au départ</button></div>`;
}

function bindPrep() {
  const P = S.prep; if (!P) return;
  const on = (id, ev, f) => { const el = $(id); if (el) el[ev] = () => f(el); };
  const addNames = () => { const el = $('#pAdd'); el.value.split(/[,;\n]+/).map(s => s.trim()).filter(Boolean).forEach(nm => { if (!P.agents.some(a => norm(a.name) === norm(nm))) P.agents.push({ name: nm, obj: P.objAll }); }); render(); };
  on('#pAddBtn', 'onclick', addNames);
  const pa = $('#pAdd'); if (pa) pa.addEventListener('keydown', e => { if (e.key === 'Enter') addNames(); });
  on('#pSame', 'onchange', el => { P.same = el.checked; render(); });
  on('#pObjAll', 'onchange', el => { P.objAll = Math.max(0, +el.value || 0); render(); });
  document.querySelectorAll('[data-pname]').forEach(el => el.onchange = () => { P.agents[+el.dataset.pname].name = el.value; render(); });
  document.querySelectorAll('[data-pobj]').forEach(el => el.onchange = () => { P.agents[+el.dataset.pobj].obj = Math.max(0, +el.value || 0); render(); });
  document.querySelectorAll('[data-pdel]').forEach(el => el.onclick = () => { P.agents.splice(+el.dataset.pdel, 1); render(); });
  document.querySelectorAll('[name="pMeth"]').forEach(el => el.onchange = () => { P.method = el.value; render(); });
  document.querySelectorAll('[data-pview]').forEach(el => el.onchange = () => { P.views[el.dataset.pview] = el.checked; render(); });
  on('#pTout', 'onchange', el => { P.tout = el.checked; render(); });
  on('#pT2', 'onchange', el => { P.t2 = el.checked; render(); });
  on('#pT2n', 'onchange', el => { P.t2Name = el.value.trim(); P.t2 = !!P.t2Name; saveTeamsInLink(); render(); });
  on('#pTeam', 'onchange', el => { P.teamName = el.value.trim(); saveTeamsInLink(); render(); });
  on('#pProt', 'onchange', el => { P.protect = el.checked; render(); });
  on('#pCheck', 'onchange', el => { P.checkbox = el.checked; render(); });
  document.querySelectorAll('[name="pMode"]').forEach(el => el.onchange = () => { P.mode = el.value; render(); });
  on('#pAll', 'onchange', el => { P.all = el.checked; render(); });
  on('#pSort', 'onchange', el => { P.sort = el.checked; render(); });
  on('#pGo', 'onclick', () => prepDownload());
  const sh = cur(), crossed = sh.crossed || (sh.crossed = new Set());
  const setSrc = (f, j) => { // relie la colonne finale f à la colonne j de ton fichier (-1 = rien)
    const c = P.cols[f]; if (!c) return;
    if (j >= 0) { crossed.delete(j); P.cols.forEach((o, k) => { if (k !== f && srcIndex(sh, o) === j) unlink(k); }); }
    if (c.key) { sh.map[c.key] = j; if (c.key === 'date' && j < 0) sh.map.dateStock = -1; compute(sh); } else c.src = j >= 0 ? sh.headers[j] : '';
  };
  const unlink = f => setSrc(f, -1);
  const done = () => { S.prepPend = null; render(); };
  document.querySelectorAll('[data-psx]').forEach(el => el.onclick = e => { e.stopPropagation(); const j = +el.dataset.psx;
    if (crossed.has(j)) crossed.delete(j); else { crossed.add(j); P.cols.forEach((c, f) => { if (srcIndex(sh, c) === j) unlink(f); }); } done(); });
  document.querySelectorAll('[data-psk]').forEach(el => el.onclick = e => { e.stopPropagation(); const j = +el.dataset.psk; P.cols.push({ src: sh.headers[j], header: (sh.headers[j] || 'COLONNE').toUpperCase() }); done(); });
  document.querySelectorAll('[data-pss]').forEach(el => el.onclick = () => { const j = +el.dataset.pss; if (crossed.has(j)) return;
    const p = S.prepPend; if (p && p.t === 'f') { setSrc(p.i, j); return done(); }
    S.prepPend = p && p.t === 's' && p.i === j ? null : { t: 's', i: j }; render(); });
  document.querySelectorAll('[data-pfc]').forEach(el => el.onclick = () => { const f = +el.dataset.pfc;
    const p = S.prepPend; if (p && p.t === 's') { setSrc(f, p.i); return done(); }
    S.prepPend = p && p.t === 'f' && p.i === f ? null : { t: 'f', i: f }; render(); });
  document.querySelectorAll('[data-pfx]').forEach(el => el.onclick = e => { e.stopPropagation(); P.cols.splice(+el.dataset.pfx, 1); done(); });
  document.querySelectorAll('[data-pfm]').forEach(el => el.onclick = e => { e.stopPropagation(); const i = +el.dataset.pfm, k = i + +el.dataset.dir;
    if (k < 0 || k >= P.cols.length) return; [P.cols[i], P.cols[k]] = [P.cols[k], P.cols[i]]; done(); });
  document.querySelectorAll('[data-pfr]').forEach(el => el.onclick = e => { e.stopPropagation(); const c = P.cols[+el.dataset.pfr];
    const v = prompt('Nouveau nom de la colonne :', c.header); if (v && v.trim()) c.header = v.trim(); done(); });
  on('#pColReset', 'onclick', () => { P.cols = defaultCols(); sh.crossed = new Set(); S.sheets.forEach(x => { if (x.ok) { x.map = detect(x.headers, x.raw); compute(x); } }); done(); });
  document.querySelectorAll('[data-pcancel]').forEach(el => el.onclick = done);
  document.querySelectorAll('[data-prestore]').forEach(el => el.onclick = () => { crossed.clear(); done(); });
  // Glisser une colonne de ton fichier dans le fichier prêt = l'ajouter (à l'endroit où on la lâche)
  const addAt = (j, at) => { const col = { src: sh.headers[j], header: (sh.headers[j] || 'COLONNE').toUpperCase() }; if (at == null || at >= P.cols.length) P.cols.push(col); else P.cols.splice(at, 0, col); done(); };
  document.querySelectorAll('[data-pss]').forEach(el => { el.ondragstart = e => { e.dataTransfer.setData('text/plain', el.dataset.pss); e.dataTransfer.effectAllowed = 'copy'; el.classList.add('dragging'); }; el.ondragend = () => el.classList.remove('dragging'); });
  const zone = $('[data-pdrop]');
  if (zone) {
    const clear = () => zone.querySelectorAll('.drop-before').forEach(x => x.classList.remove('drop-before'));
    zone.ondragover = e => { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; clear(); const c = e.target.closest('[data-pfc]'); if (c) c.classList.add('drop-before'); zone.classList.toggle('drop-end', !c); };
    zone.ondragleave = e => { if (!zone.contains(e.relatedTarget)) { clear(); zone.classList.remove('drop-end'); } };
    zone.ondrop = e => { e.preventDefault(); clear(); zone.classList.remove('drop-end'); const j = +e.dataTransfer.getData('text/plain'); if (isNaN(j)) return;
      const c = e.target.closest('[data-pfc]'); addAt(j, c ? +c.dataset.pfc : null); };
  }
}

function prepWorkbook() {
  const P = S.prep, N = (P.t2Name || 'équipe 2').toUpperCase();
  const sheets = P.all ? S.sheets.filter(s => s.ok) : [cur()];
  const headers = [...P.cols.map(c => c.header), ...(P.t2 ? STD2.map(h => `${h} ${N}`) : [])];
  const widths = [...P.cols.map(c => STD_W[c.key] || Math.min(40, Math.max(12, c.header.length + 2))), ...(P.t2 ? [16, 14, 13, 44] : [])];
  const at = k => P.cols.findIndex(c => c.key === k);
  const iAgent = at('agent'), iTraite = at('traite'), iComment = at('comment'), iT2 = P.t2 ? P.cols.length : -1;
  const editable = new Set([iAgent, iTraite, iComment, ...(P.t2 ? [iT2, iT2 + 1, iT2 + 2, iT2 + 3] : [])].filter(i => i >= 0));
  const wb = XLSX.utils.book_new(), cfg = [];
  for (const sh of sheets) {
    const rows = prepRows(sh), dateCols = new Set();
    rows.forEach(x => x.line.forEach((v, j) => { if (v instanceof Date) { x.line[j] = excelDay(v); dateCols.add(j); } }));
    const aoa = [headers, ...rows.map(x => x.jaune ? x.line.map(v => v ?? '') : x.line)];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    for (const j of dateCols) for (let i = 1; i <= rows.length; i++) { const c = ws[XLSX.utils.encode_cell({ r: i, c: j })]; if (c && c.t === 'n') c.z = 'dd/mm/yyyy'; }
    ws['!cols'] = widths.map(w => ({ wch: w }));
    const last = rows.length + 1, ref = `A1:${colLetter(headers.length - 1)}${last}`;
    ws['!autofilter'] = { ref };
    let name = sh.name.slice(0, 31).replace(/[\\/?*[\]:]/g, '_'); while (wb.Sheets[name]) name = name.slice(0, 29) + '+';
    XLSX.utils.book_append_sheet(wb, ws, name);
    // affichages de feuille
    const views = [], used = new Set(), add = v => { let nm = v.name.slice(0, 31); while (used.has(norm(nm))) nm = nm.slice(0, 29) + '+'; used.add(norm(nm)); views.push({ ...v, name: nm }); };
    const noms = iAgent >= 0 ? [...new Set(rows.map(x => x.line[iAgent]).filter(Boolean))] : [];
    if (P.views.agents) noms.forEach(nm => add({ name: nm, col: iAgent, vals: [String(nm)] }));
    if (P.views.all && P.cols.length) { // « Tout voir » : la colonne qui a le moins de valeurs différentes, toutes cochées
      const distinct = P.cols.map((c, j) => [...new Set(rows.map(x => x.line[j]).filter(v => v != null && v !== '').map(String))]);
      const j = distinct.reduce((best, v, k) => !editable.has(k) && v.length && (best < 0 || v.length < distinct[best].length) ? k : best, -1); // jamais une colonne que l'équipe remplit
      if (j >= 0 && distinct[j].length <= 200) add({ name: 'Tout voir', col: j, vals: distinct[j], blank: true });
    }
    if (P.views.pot && iAgent >= 0 && rows.some(x => !x.line[iAgent])) add({ name: 'Non attribuées', col: iAgent, vals: [], blank: true });
    if (P.t2 && P.views.t2) add({ name: P.t2Name || 'Autre équipe', col: iT2, vals: P.checkbox ? ['VRAI', 'TRUE'] : ['oui'] });
    if (P.t2 && P.views.team) add({ name: P.teamName || 'Mon équipe', col: iT2, vals: P.checkbox ? ['FAUX', 'FALSE'] : ['non'], blank: true });
    cfg.push({
      lists: [...(iTraite >= 0 ? [[`${colLetter(iTraite)}2:${colLetter(iTraite)}${last}`, LISTE_TRAITE]] : []),
        ...(P.t2 && !P.checkbox ? [[`${colLetter(iT2)}2:${colLetter(iT2)}${last}`, ['oui', 'non']]] : []), ...(P.t2 ? [[`${colLetter(iT2 + 2)}2:${colLetter(iT2 + 2)}${last}`, LISTE_TRAITE]] : [])],
      yellow: new Set(rows.map((x, i) => x.jaune ? i + 2 : 0).filter(Boolean)), t2From: P.t2 ? iT2 : Infinity,
      views, ref, protect: P.protect, editable, checkCol: P.t2 && P.checkbox ? iT2 : -1,
      colors: [
        ...(iAgent >= 0 ? [{ ref: `${colLetter(iAgent)}2:${colLetter(iAgent)}${last}`, names: [...new Set([...prepTeam().map(a => a.name), ...rows.map(x => x.line[iAgent]).filter(Boolean)].map(String))], other: false }] : []),
        ...(P.t2 ? [{ ref: `${colLetter(iT2 + 1)}2:${colLetter(iT2 + 1)}${last}`, names: [...new Set(rows.map(x => x.line[iT2 + 1]).filter(Boolean).map(String))], other: true }] : []),
      ],
    });
  }
  return polishXlsx(new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' })), cfg);
}

function prepDownload() {
  const u8 = prepWorkbook();
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([u8], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  a.download = `${S.fileName.replace(/\.[^.]+$/, '')} - prêt ${isoDay(new Date())}.xlsx`;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}

/* La version gratuite de la bibliothèque n'écrit ni listes déroulantes, ni couleurs, ni affichages :
   on les ajoute directement dans le fichier .xlsx (un zip de fichiers XML). */
function polishXlsx(u8, cfg) {
  const zip = XLSX.CFB.read(u8, { type: 'array' });
  const entry = path => zip.FileIndex[zip.FullPaths.findIndex(p => p.endsWith('/' + path))];
  const dec = e => new TextDecoder().decode(e.content), enc = (e, s) => { e.content = new TextEncoder().encode(s); e.size = e.content.length; };
  const guid = () => { const b = crypto.getRandomValues(new Uint8Array(16)); b[6] = b[6] & 15 | 64; b[8] = b[8] & 63 | 128; const h = [...b].map(x => x.toString(16).padStart(2, '0')).join('').toUpperCase(); return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`; };

  // styles : en-tête coloré, lignes jaunes, cellules modifiables (déverrouillées)
  const se = entry('xl/styles.xml'); let st = dec(se);
  const addTo = (tag, item) => { let id = 0; st = st.replace(new RegExp(`<${tag} count="(\\d+)">([\\s\\S]*?)</${tag}>`), (m0, n, body) => { id = +n; return `<${tag} count="${id + 1}">${body}${item}</${tag}>`; }); return id; };
  const fontB = addTo('fonts', '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>');
  const fill = rgb => addTo('fills', `<fill><patternFill patternType="solid"><fgColor rgb="FF${rgb}"/><bgColor indexed="64"/></patternFill></fill>`);
  const fHead = fill('5B5FC7'), fHead2 = fill('1BAF7A'), fYellow = fill('FFFF00');
  const xfs = [...st.match(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/)[1].matchAll(/<xf [^>]*?\/>|<xf [^>]*?>[\s\S]*?<\/xf>/g)].map(x => x[0]);
  const xHead = addTo('cellXfs', `<xf numFmtId="0" fontId="${fontB}" fillId="${fHead}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>`);
  const xHead2 = addTo('cellXfs', `<xf numFmtId="0" fontId="${fontB}" fillId="${fHead2}" borderId="0" xfId="0" applyFont="1" applyFill="1"/>`);
  const variants = new Map();
  const variant = (s, yellow, unlock) => {
    const key = `${s}|${yellow}|${unlock}`; if (!yellow && !unlock) return s;
    if (!variants.has(key)) {
      let x = (xfs[s] || xfs[0]).replace(/<xf ([^>]*?)\/>/, '<xf $1></xf>');
      if (yellow) x = x.replace(/fillId="\d+"/, `fillId="${fYellow}"`).replace(/ applyFill="\d"/, '').replace(/<xf /, '<xf applyFill="1" ');
      if (unlock) x = x.replace(/ applyProtection="\d"/, '').replace(/<xf /, '<xf applyProtection="1" ').replace(/<protection[^>]*\/>/, '').replace('</xf>', '<protection locked="0"/></xf>');
      variants.set(key, addTo('cellXfs', x));
    }
    return variants.get(key);
  };
  const xUnlock = cfg.some(c => c.protect) ? variant(0, false, true) : 0;
  const hasCheck = cfg.some(c => c.checkCol >= 0), anyProt = cfg.some(c => c.protect);
  const xCheck = hasCheck ? addTo('cellXfs', `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"${anyProt ? ' applyProtection="1"' : ''}>${anyProt ? '<protection locked="0"/>' : ''}<extLst><ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><xfpb:xfComplement i="0"/></ext></extLst></xf>`) : 0;

  const dxfList = [], dxfIds = new Map();
  const dxfFor = ([fg, bg]) => { const k = fg + bg; if (!dxfIds.has(k)) { dxfIds.set(k, dxfList.length); dxfList.push(`<dxf><font><b/><color rgb="FF${fg.slice(1).toUpperCase()}"/></font><fill><patternFill patternType="solid"><bgColor rgb="FF${bg.slice(1).toUpperCase()}"/></patternFill></fill></dxf>`); } return dxfIds.get(k); };
  const ct = entry('[Content_Types].xml'); let ctx = dec(ct);
  cfg.forEach((c, i) => {
    const e = entry(`xl/worksheets/sheet${i + 1}.xml`); let x = dec(e);
    x = x.replace(/<c r="([A-Z]+)(\d+)"([^>]*?)(\/?)>/g, (m0, col, row, attrs, close) => {
      const r = +row, ci = colIndex(col); let s = +((attrs.match(/ s="(\d+)"/) || [])[1] || 0);
      if (r === 1) s = ci >= c.t2From ? xHead2 : xHead;
      else if (ci === c.checkCol) s = xCheck; // case à cocher Excel
      else { const y = c.yellow.has(r), u = c.protect && c.editable.has(ci); if (!y && !u) return m0; s = variant(s, y, u); }
      return `<c r="${col}${row}"${attrs.replace(/ s="\d+"/, '')} s="${s}"${close}>`;
    });
    x = x.replace(/<sheetView([^>]*?)\/>/, '<sheetView$1><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView>');
    if (c.protect) {
      x = x.replace(/<col min="(\d+)" max="(\d+)"([^>]*?)\/>/g, (m0, a, b, rest) => c.editable.has(+a - 1) ? `<col min="${a}" max="${b}"${rest} style="${xUnlock}"/>` : m0);
      x = x.replace('</sheetData>', `</sheetData><sheetProtection sheet="1" objects="1" scenarios="1" formatColumns="0" formatRows="0" autoFilter="0" sort="0"/>`);
    }
    let prio = 1;
    const cf = (c.colors || []).filter(g => g.names.length).map(g => `<conditionalFormatting sqref="${g.ref}">` + g.names.map(nm => {
      const id = dxfFor(nameColor(nm, g.other)); return `<cfRule type="cellIs" dxfId="${id}" priority="${prio++}" operator="equal"><formula>"${xmlEsc(nm.replace(/"/g, '""'))}"</formula></cfRule>`; }).join('') + '</conditionalFormatting>').join('');
    const dv = !c.lists.length ? '' : `<dataValidations count="${c.lists.length}">` + c.lists.map(([ref, list]) =>
      `<dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Valeur non prévue" error="Choisis une valeur dans la liste." sqref="${ref}"><formula1>"${list.join(',')}"</formula1></dataValidation>`).join('') + '</dataValidations>';
    x = x.replace(/(<hyperlinks|<printOptions|<pageMargins|<pageSetup|<headerFooter|<rowBreaks|<colBreaks|<ignoredErrors|<drawing|<legacyDrawing|<tableParts|<extLst|<\/worksheet>)/, cf + dv + '$1');
    enc(e, x);

    if (c.views.length) { // un « affichage de feuille » par vue : chacun filtre sans gêner les autres
      const nsv = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><namedSheetViews xmlns="http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
        c.views.map(v => `<namedSheetView name="${xmlEsc(v.name)}" id="${guid()}"><nsvFilter filterId="${guid()}" ref="${c.ref}"><columnFilter colId="${v.col}"><filter colId="${v.col}"><x:filters${v.blank ? ' blank="1"' : ''}>${v.vals.map(val => `<x:filter val="${xmlEsc(val)}"/>`).join('')}</x:filters></filter></columnFilter></nsvFilter></namedSheetView>`).join('') +
        '</namedSheetViews>';
      XLSX.CFB.utils.cfb_add(zip, `xl/namedSheetViews/namedSheetView${i + 1}.xml`, new TextEncoder().encode(nsv));
      XLSX.CFB.utils.cfb_add(zip, `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, new TextEncoder().encode(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdNsv1" Type="http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView" Target="../namedSheetViews/namedSheetView${i + 1}.xml"/></Relationships>`));
      ctx = ctx.replace('</Types>', `<Override PartName="/xl/namedSheetViews/namedSheetView${i + 1}.xml" ContentType="application/vnd.ms-excel.namedsheetviews+xml"/></Types>`);
    }
  });
  if (hasCheck) { // les cases à cocher d'Excel 365 sont décrites dans une partie « featurePropertyBag »
    XLSX.CFB.utils.cfb_add(zip, 'xl/featurePropertyBag/featurePropertyBag.xml', new TextEncoder().encode('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag>'
      + '<bag type="XFComplement"><bagId k="XFControls">1</bagId></bag><bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag></FeaturePropertyBags>'));
    ctx = ctx.replace('</Types>', '<Override PartName="/xl/featurePropertyBag/featurePropertyBag.xml" ContentType="application/vnd.ms-excel.featurepropertybag+xml"/></Types>');
    const wr = entry('xl/_rels/workbook.xml.rels');
    enc(wr, dec(wr).replace('</Relationships>', '<Relationship Id="rIdFpb1" Type="http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag" Target="featurePropertyBag/featurePropertyBag.xml"/></Relationships>'));
  }
  if (dxfList.length) st = st.replace(/<dxfs count="0"\s*\/>|<dxfs count="0"><\/dxfs>/, `<dxfs count="${dxfList.length}">${dxfList.join('')}</dxfs>`);
  enc(ct, ctx);
  enc(se, st);
  return XLSX.CFB.write(zip, { fileType: 'zip', type: 'array', compression: true });
}
const colIndex = L => [...L].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
