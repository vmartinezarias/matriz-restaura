/* Matriz RESTAURA — visualizador interactivo con observaciones
   La matriz se lee directamente del Excel del repositorio (SheetJS).
   Las observaciones se guardan en el navegador y, si hay token, en un JSON del repositorio. */
(() => {
'use strict';

/* ============================================================
   Utilidades
   ============================================================ */
const CFG = Object.assign({
  excel: 'data/Matriz_RESTAURA.xlsx', owner: '', repo: '', branch: 'main',
  obsPath: 'observaciones/observaciones.json',
}, window.RESTAURA_CONFIG || {});

const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const nowIso = () => new Date().toISOString();
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch { /* sin almacenamiento */ } },
  del(k) { try { localStorage.removeItem(k); } catch { /* nada */ } },
};
const fmtDate = iso => {
  if (!iso) return '';
  const d = new Date(iso);
  return d.toLocaleString('es-CO', { day: 'numeric', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' });
};
const plural = (n, s, p) => `${n} ${n === 1 ? s : p}`;
let toastT;
function toast(msg, ms = 2800) {
  const t = $('#toast'); t.textContent = msg; t.classList.add('on');
  clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('on'), ms);
}
const ICON = {
  chev: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m9 6 6 6-6 6"/></svg>',
  note: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M21 12a8 8 0 0 1-11.6 7.1L4 20l1-4.6A8 8 0 1 1 21 12z"/><path d="M12 9v6M9 12h6"/></svg>',
};

/* ============================================================
   Modelo de la matriz
   ============================================================ */
// [clave interna, etiqueta, encabezado normalizado esperado, sección]
const FIELDS = [
  ['fase', 'Fase', 'fase', 'h'],
  ['componente', 'Componente evaluado', 'componente evaluado', 'h'],
  ['unidad', 'Unidad de monitoreo', 'unidad de monitoreo', 'h'],
  ['entidad', 'Entidad SIG / tabla (geometría)', 'entidad sig / tabla (geometria)', 'sig'],
  ['llaves', 'Llaves / relaciones', 'llaves / relaciones', 'sig'],
  ['ind_cod', 'Código del indicador', 'indicador (cod.)', 'ind'],
  ['ind_nombre', 'Indicador', 'indicador', 'ind'],
  ['ind_desc', 'Descripción del indicador', 'descripcion del indicador', 'ind'],
  ['ind_formula', 'Cálculo / fórmula del indicador', 'calculo / formula del indicador', 'ind'],
  ['ind_unidad', 'Unidad de medida', 'unidad de medida del indicador', 'ind'],
  ['ind_umbrales', 'Rangos / umbrales (semáforo)', 'rangos / umbrales (semaforo)', 'ind'],
  ['ind_agreg', 'Regla de agregación', 'regla de agregacion', 'ind'],
  ['var_cod', 'Código de sistema', 'variable (cod. de sistema)', 'var'],
  ['var_nombre', 'Nombre de la variable', 'nombre de la variable', 'var'],
  ['var_desc', 'Descripción de la variable', 'descripcion de la variable', 'var'],
  ['var_tipo', 'Tipo de dato', 'tipo de dato', 'var'],
  ['var_dominio', 'Dominio / rangos de valores', 'dominio / rangos de valores', 'var'],
  ['var_unidad', 'Unidad', 'unidad', 'var'],
  ['var_formula', 'Cálculo / fórmula de la variable', 'calculo / formula de la variable', 'var'],
  ['metodo', 'Método de captura / instrumento', 'metodo de captura / instrumento', 'cap'],
  ['momentos', 'Momento(s) de medición', 'momento(s) de medicion', 'cap'],
  ['frecuencia', 'Frecuencia', 'frecuencia', 'cap'],
  ['nivel', 'Fuente / nivel', 'fuente / nivel (n1, n2, n3, td)', 'cap'],
  ['responsable', 'Responsable', 'responsable', 'cap'],
  ['validacion', 'Validación', 'validacion', 'val'],
  ['obligatorio', 'Obligatorio', 'obligatorio', 'val'],
].map(([k, label, match, sec]) => ({ k, label, match, sec }));
const FIELD = Object.fromEntries(FIELDS.map(f => [f.k, f]));
const SECTIONS = [
  ['ind', 'Indicador'], ['var', 'Variable'], ['cap', 'Captura y responsables'],
  ['val', 'Validación'], ['sig', 'Ubicación en el modelo SIG'],
];
const OWN_SECTIONS = { indicador: ['ind'], variable: ['var', 'cap', 'val'], unidad: ['sig'] };
const FORMULA_KEYS = new Set(['ind_formula', 'var_formula']);
const TITLE_KEYS = new Set(['fase', 'componente', 'unidad', 'ind_cod', 'ind_nombre', 'var_cod', 'var_nombre']);
const HIER = ['fase', 'componente', 'unidad', 'ind_cod', 'var_cod'];

const LEVELS = {
  fase: { k: 'fase', name: 'Fase', plural: 'Fases' },
  componente: { k: 'componente', name: 'Componente', plural: 'Componentes' },
  unidad: { k: 'unidad', name: 'Unidad de monitoreo', plural: 'Unidades de monitoreo' },
  indicador: { k: 'ind_cod', name: 'Indicador', plural: 'Indicadores' },
  variable: { k: 'var_cod', name: 'Variable', plural: 'Variables' },
};
const VIEWS = [
  { id: 'fase', name: 'Fase → componente → unidad → indicador', levels: ['fase', 'componente', 'unidad', 'indicador', 'variable'] },
  { id: 'indicador', name: 'Indicador → fase → unidad', levels: ['indicador', 'fase', 'unidad', 'variable'] },
  { id: 'unidad', name: 'Unidad → fase → indicador', levels: ['unidad', 'fase', 'indicador', 'variable'] },
  { id: 'componente', name: 'Componente → indicador', levels: ['componente', 'indicador', 'variable'] },
];
const MOMENTOS = ['T-1', 'T0', 'M-A', 'M-S', 'M-P', 'E6', 'E12', 'E18', 'E24', 'A3', 'A4', 'A5', 'A10', 'A15', 'A20'];
const NIVELES = [['N1', '100 % de unidades'], ['N2', 'submuestra con control'], ['N3', 'parcelas permanentes'], ['TD', 'teledetección']];
const TIPOS = [
  ['comentario', 'Comentario'], ['ajuste', 'Propuesta de ajuste'],
  ['pregunta', 'Pregunta'], ['error', 'Error o inconsistencia'],
];
const TIPO = Object.fromEntries(TIPOS);

const S = {
  rows: [], meta: { title: '', note: '', source: '' }, entities: [], entityNotes: [], indInfo: {},
  phases: [], view: VIEWS[0], tree: null, index: new Map(), sel: null,
  expanded: new Set(store.get('restaura.exp', [])),
  filters: { q: '', niveles: new Set(), momento: '', oblig: '', conObs: false },
  visible: new Set(), obsDeep: store.get('restaura.deep', true),
  campo: '', editing: null,
};

const consKey = c => HIER.filter(k => k in c).map(k => `${k}=${c[k]}`).join('|');
const rowCons = r => Object.fromEntries(HIER.map(k => [k, r[k]]));
const matches = (cons, r) => Object.keys(cons).every(k => (r[k] ?? '') === cons[k]);
const within = (target, cons) => Object.keys(cons).every(k => target[k] === cons[k]);

function parsePhase(s) {
  const m = String(s).match(/^(F\d+)\s*[·\-–:]\s*(.*?)\s*(\((.*)\))?$/);
  if (!m) return { code: '', name: s, note: '' };
  return { code: m[1], name: m[2], note: m[4] || '' };
}
function phaseIdx(fase) {
  const m = String(fase).match(/^F(\d+)/);
  return m ? Math.min(+m[1], 4) : -1;
}
function phaseVars(idx) {
  return idx < 0 ? '--phc: var(--phx); --phi: #fff;' : `--phc: var(--ph${idx}); --phi: var(--on-ph${idx});`;
}

/* ---------- Lectura del Excel ---------- */
function parseWorkbook(wb, source) {
  const mainName = wb.SheetNames.find(n => /matriz/i.test(n)) || wb.SheetNames[0];
  const ws = wb.Sheets[mainName];
  const start = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']).s.r : 0;
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, blankrows: true });
  const hi = aoa.findIndex(r => norm(r[0]) === 'fase');
  if (hi < 0) throw new Error(`No encontré la fila de encabezados (la que empieza con «Fase») en la hoja «${mainName}».`);

  const headers = aoa[hi].map(norm);
  const used = new Set(); const col = {};
  FIELDS.forEach((f, i) => {
    let j = headers.findIndex((h, x) => !used.has(x) && h === f.match);
    if (j < 0) j = headers.findIndex((h, x) => !used.has(x) && h.startsWith(f.match.slice(0, 12)));
    if (j < 0) j = i;
    used.add(j); col[f.k] = j;
  });

  const pre = aoa.slice(0, hi).map(r => r.filter(c => String(c).trim()).join(' ').trim()).filter(Boolean);
  S.meta = { title: pre[0] || '', note: pre.find((p, i) => i > 0 && p.length > 80) || '', source, sheet: mainName };

  S.rows = [];
  aoa.slice(hi + 1).forEach((r, i) => {
    if (!r.some(c => String(c).trim())) return;
    const o = { _i: S.rows.length, _excelRow: start + hi + 2 + i };
    FIELDS.forEach(f => { o[f.k] = String(r[col[f.k]] ?? '').trim(); });
    if (!o.fase && !o.var_cod) return;
    o._text = norm(FIELDS.map(f => o[f.k]).join(' '));
    S.rows.push(o);
  });
  if (!S.rows.length) throw new Error('La hoja de la matriz no tiene filas de variables.');

  // Hoja de entidades SIG
  S.entities = []; S.entityNotes = [];
  const entName = wb.SheetNames.find(n => /entidad|modelo/i.test(n));
  if (entName) {
    const ea = XLSX.utils.sheet_to_json(wb.Sheets[entName], { header: 1, defval: '', raw: false });
    const eh = ea.findIndex(r => norm(r[0]) === 'codigo');
    if (eh >= 0) {
      ea.slice(eh + 1).forEach(r => {
        const cells = r.map(c => String(c).trim());
        if (!cells.some(Boolean)) return;
        if (cells[0] && cells[1]) S.entities.push({ cod: cells[0], entidad: cells[1], geom: cells[2], llaves: cells[3], fuente: cells[4], n: cells[5] });
        else S.entityNotes.push(cells.filter(Boolean).join(' '));
      });
    }
  }
  // Hoja resumen de indicadores (tipo, alcance, revisiones)
  S.indInfo = {};
  for (const name of wb.SheetNames) {
    const a = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '', raw: false });
    const h = a.findIndex(r => norm(r[0]) === 'codigo' && r.some(c => norm(c) === 'revisiones'));
    if (h < 0) continue;
    const hd = a[h].map(norm);
    const ix = k => hd.findIndex(x => x.startsWith(k));
    const iT = ix('tipo'), iA = ix('alcance'), iR = ix('revisiones');
    a.slice(h + 1).forEach(r => {
      const cod = String(r[0]).trim(); if (!cod) return;
      const clean = v => { const s = String(v ?? '').trim(); return s.startsWith('#') ? '' : s; };
      S.indInfo[cod] = { tipo: clean(r[iT]), alcance: clean(r[iA]), revisiones: clean(r[iR]) };
    });
    break;
  }
  S.phases = [...new Set(S.rows.map(r => r.fase))].filter(Boolean);
}

/* ---------- Árbol ---------- */
function nodeLabel(n) {
  const r = n.rows[0];
  if (n.level === 'fase') { const p = parsePhase(n.value); return p.code ? `${p.code} ${p.name}` : n.value; }
  if (n.level === 'indicador') return r.ind_nombre || n.value;
  if (n.level === 'variable') return r.var_nombre || n.value;
  return n.value;
}
function nodeCode(n) {
  if (n.level === 'indicador') return n.value;
  if (n.level === 'variable') return n.value;
  return '';
}
function buildTree(view) {
  const root = { key: '', level: null, value: '', rows: S.rows, cons: {}, parent: null, children: [], depth: 0 };
  S.index = new Map([['', root]]);
  const rec = (node, depth) => {
    if (depth >= view.levels.length) return;
    const lv = LEVELS[view.levels[depth]];
    const groups = new Map();
    for (const r of node.rows) {
      const v = r[lv.k] || '(sin valor)';
      if (!groups.has(v)) groups.set(v, []);
      groups.get(v).push(r);
    }
    for (const [v, rs] of groups) {
      const isLeaf = depth === view.levels.length - 1;
      let cons = { ...node.cons, [lv.k]: v };
      if (isLeaf && rs.length === 1) cons = rowCons(rs[0]);
      const child = { level: view.levels[depth], value: v, rows: rs, cons, parent: node, children: [], depth: depth + 1 };
      child.key = consKey(cons);
      child.label = nodeLabel(child);
      child.code = nodeCode(child);
      const fases = new Set(rs.map(r => r.fase));
      child.phase = fases.size === 1 ? phaseIdx(rs[0].fase) : -1;
      node.children.push(child);
      S.index.set(child.key, child);
      rec(child, depth + 1);
    }
  };
  rec(root, 0);
  return root;
}
const ancestors = n => { const a = []; for (let p = n.parent; p; p = p.parent) a.unshift(p); return a; };
const visCount = n => n.rows.reduce((s, r) => s + (S.visible.has(r._i) ? 1 : 0), 0);

function nodeForTarget(t) {
  const direct = S.index.get(consKey(t));
  if (direct) return direct;
  let node = S.tree;
  for (;;) {
    const next = node.children.find(c => Object.keys(c.cons).every(k => !(k in t) ? false : t[k] === c.cons[k]));
    if (!next) {
      // Coincidencia parcial: el hijo cuyas filas incluyen el objetivo
      const cand = node.children.filter(c => c.rows.some(r => matches(t, r)));
      if (cand.length === 1 && Object.keys(t).length >= Object.keys(cand[0].cons).length) { node = cand[0]; continue; }
      return node;
    }
    node = next;
  }
}

/* ============================================================
   Filtros
   ============================================================ */
function computeVisible() {
  const F = S.filters;
  const toks = norm(F.q).split(' ').filter(Boolean);
  const momRe = F.momento ? new RegExp(`(^|[^A-Za-z0-9-])${F.momento.replace(/-/g, '\\-')}(?![0-9])`) : null;
  const act = F.conObs ? Obs.active() : null;
  S.visible = new Set(S.rows.filter(r => {
    if (toks.length && !toks.every(t => r._text.includes(t))) return false;
    if (F.niveles.size && ![...F.niveles].some(n => r.nivel.toUpperCase().includes(n))) return false;
    if (momRe && !momRe.test(`${r.momentos} ${r.frecuencia}`)) return false;
    const si = /^s[ií]\b/i.test(r.obligatorio);
    if (F.oblig === 'si' && !si) return false;
    if (F.oblig === 'no' && si) return false;
    if (act && !act.some(o => matches(o.target, r))) return false;
    return true;
  }).map(r => r._i));
}
const filtersActive = () => {
  const F = S.filters;
  return !!(F.q || F.niveles.size || F.momento || F.oblig || F.conObs);
};

function renderFilters() {
  const F = S.filters;
  const n = S.visible.size;
  $('#filters').innerHTML = `
    <div class="fgroup" role="group" aria-label="Nivel de evidencia"><span>Nivel</span>
      ${NIVELES.map(([c, d]) => `<button class="chip" data-niv="${c}" aria-pressed="${F.niveles.has(c)}" title="${esc(d)}">${c}</button>`).join('')}
    </div>
    <label class="fgroup"><span>Momento</span>
      <select id="fMom"><option value="">Todos</option>${MOMENTOS.map(m => `<option ${F.momento === m ? 'selected' : ''}>${m}</option>`).join('')}</select>
    </label>
    <label class="fgroup"><span>Obligatoria</span>
      <select id="fObl">
        <option value="">Todas</option>
        <option value="si" ${F.oblig === 'si' ? 'selected' : ''}>Sí</option>
        <option value="no" ${F.oblig === 'no' ? 'selected' : ''}>No o condicional</option>
      </select>
    </label>
    <label class="check"><input type="checkbox" id="fObs" ${F.conObs ? 'checked' : ''}> Con observaciones</label>
    ${filtersActive() ? '<button class="linkish" id="fClear">Quitar filtros</button>' : ''}
    <span class="fcount">${n === S.rows.length ? plural(n, 'variable', 'variables') : `${n} de ${S.rows.length} variables`}</span>`;
}

/* ============================================================
   Render: ribbon, árbol, detalle
   ============================================================ */
function renderRibbon() {
  const cur = S.sel && S.sel !== S.tree ? new Set(S.sel.rows.map(r => r.fase)) : null;
  $('#ribbon').innerHTML = S.phases.map(f => {
    const p = parsePhase(f); const i = phaseIdx(f);
    const rows = S.rows.filter(r => r.fase === f);
    const nv = rows.filter(r => S.visible.has(r._i)).length;
    const on = cur && cur.size === 1 && cur.has(f);
    return `<button class="ph${on ? ' on' : ''}${nv ? '' : ' dim'}" data-phase="${esc(f)}" style="${phaseVars(i)} --n:${Math.max(rows.length, 6)}" title="${esc(f)}">
      <b>${esc(p.code || f)}</b><span>${esc(p.name)}</span><small>${p.note ? esc(p.note) + ', ' : ''}${plural(nv, 'variable', 'variables')}</small></button>`;
  }).join('');
}

function renderTree() {
  const out = [];
  const auto = !!S.filters.q || S.filters.conObs;
  const openPath = new Set(S.sel ? ancestors(S.sel).map(a => a.key) : []);
  const walk = (node, depth) => {
    for (const c of node.children) {
      const n = visCount(c); if (!n) continue;
      const kids = c.children.length > 0;
      const open = kids && (auto || S.expanded.has(c.key) || openPath.has(c.key));
      const ob = Obs.countOpen(c);
      const k = esc(c.key);
      out.push(`<div class="ti lv-${c.level}${S.sel === c ? ' sel' : ''}" role="treeitem" aria-level="${depth}" ${kids ? `aria-expanded="${open}"` : ''} aria-selected="${S.sel === c}" style="--d:${depth - 1}; ${c.phase >= 0 ? `--phc: var(--ph${c.phase});` : ''}">
        ${kids ? `<button class="tw" data-tog="${k}" aria-label="${open ? 'Contraer' : 'Expandir'}">${ICON.chev}</button>` : '<span class="tw-sp"></span>'}
        <button class="tl" data-go="${k}">${c.level === 'indicador' ? `<span class="code">${esc(c.code)}</span>` : ''}<span class="tx">${esc(c.label)}</span></button>
        ${ob ? `<span class="ob" title="${plural(ob, 'observación abierta', 'observaciones abiertas')}">${ob}</span>` : ''}
        <span class="tc">${kids ? n : ''}</span>
      </div>`);
      if (open) walk(c, depth + 1);
    }
  };
  walk(S.tree, 1);
  $('#tree').innerHTML = out.length ? out.join('') : '<p class="tree-empty">Ninguna variable coincide con la búsqueda o los filtros.</p>';
  const sel = $('.ti.sel', $('#tree'));
  if (sel) sel.scrollIntoView({ block: 'nearest' });
}

function constFields(node) {
  const m = {};
  for (const f of FIELDS) {
    const v0 = node.rows[0][f.k];
    if (node.rows.every(r => r[f.k] === v0)) m[f.k] = v0;
  }
  return m;
}

function semaforo(html) {
  return html
    .replace(/\b(Verde|verde)\b/g, '<i class="sem g" aria-hidden="true"></i>$1')
    .replace(/\b(Amarillo|amarillo)\b/g, '<i class="sem y" aria-hidden="true"></i>$1')
    .replace(/\b(Rojo|rojo)\b/g, '<i class="sem r" aria-hidden="true"></i>$1');
}
function fmtInline(k, v) {
  let h = esc(v);
  if (k === 'ind_umbrales' || k === 'var_dominio') h = semaforo(h);
  return h;
}
function fmtVal(k, v) {
  if (!v || v === '—' || v === '-') return '<span class="muted">—</span>';
  if (FORMULA_KEYS.has(k)) return `<code class="formula">${esc(v)}</code>`;
  if (k === 'var_cod') return `<code>${esc(v)}</code>`;
  const parts = v.split(/\s+·\s+/);
  if (parts.length >= 3) {
    return `<ul class="cat">${parts.map(p => {
      const m = p.match(/^([A-Z]{1,4}\d{0,3}[a-z]?)(\s+)(.+)$/);
      return `<li>${m ? `<b class="cc">${esc(m[1])}</b>${fmtInline(k, m[3])}` : fmtInline(k, p)}</li>`;
    }).join('')}</ul>`;
  }
  return fmtInline(k, v);
}

function fieldRows(keys, vals) {
  return `<dl class="fields">${keys.map(k => `
    <div class="f"><dt>${esc(FIELD[k].label)}</dt><dd>${fmtVal(k, vals[k])}</dd>
      <button class="cm" data-campo="${k}" title="Anotar sobre «${esc(FIELD[k].label)}»" aria-label="Anotar sobre ${esc(FIELD[k].label)}">${ICON.note}</button></div>`).join('')}</dl>`;
}

function sectionHtml(sec, title, keys, vals) {
  let head = `<h2>${esc(title)}</h2>`;
  if (sec === 'ind' && vals.ind_cod) {
    head = `<span class="code">${esc(vals.ind_cod)}</span><h2>${esc(vals.ind_nombre || 'Indicador')}</h2>`;
  } else if (sec === 'var' && vals.var_nombre) {
    head = `<h2>${esc(vals.var_nombre)}</h2>`;
  }
  return `<section class="sec s-${sec}"><header>${head}</header>${fieldRows(keys, vals)}</section>`;
}

function sectionsFor(vals, onlyKeys) {
  const html = [];
  for (const [sec, title] of SECTIONS) {
    const keys = FIELDS.filter(f => f.sec === sec && onlyKeys.has(f.k) && !(TITLE_KEYS.has(f.k) && f.k !== 'var_cod')).map(f => f.k);
    const hasTitle = (sec === 'ind' && onlyKeys.has('ind_cod')) || (sec === 'var' && onlyKeys.has('var_nombre'));
    if (!keys.length && !hasTitle) continue;
    let extra = '';
    if (sec === 'ind' && vals.ind_cod && S.indInfo[vals.ind_cod]) {
      const ii = S.indInfo[vals.ind_cod];
      const add = [['Tipo', ii.tipo], ['Alcance', ii.alcance], ['Revisiones', ii.revisiones]].filter(x => x[1]);
      if (add.length) extra = `<dl class="fields">${add.map(([l, v]) => `<div class="f"><dt>${l}</dt><dd>${esc(v)}</dd><span></span></div>`).join('')}</dl>`;
    }
    const sh = sectionHtml(sec, title, keys, vals);
    html.push(extra ? sh.replace(/<\/section>$/, `${extra}</section>`) : sh);
  }
  return html.join('');
}

function renderDetail() {
  const n = S.sel || S.tree;
  const el = $('#detail');
  $('#mainCol').scrollTop = 0;
  if (n === S.tree) { el.innerHTML = homeHtml(); el.scrollTop = 0; return; }

  const anc = ancestors(n);
  const crumbs = anc.map(a => `<button data-go="${esc(a.key)}">${a === S.tree ? 'Inicio' : esc(a.code ? `${a.code} ${a.label}` : a.label)}</button><span class="sep">/</span>`).join('');
  const lv = LEVELS[n.level];
  const vals = constFields(n);
  const parentVals = n.parent && n.parent !== S.tree ? constFields(n.parent) : {};
  const own = OWN_SECTIONS[n.level] || [];
  const newKeys = new Set(Object.keys(vals).filter(k => !(k in parentVals) || own.includes(FIELD[k].sec)));
  const inhKeys = new Set(Object.keys(vals).filter(k => !newKeys.has(k) && !['fase', 'componente', 'unidad'].includes(k)));
  const nVis = visCount(n);
  const p = n.level === 'fase' ? parsePhase(n.value) : null;

  let sub = '';
  if (n.level === 'variable') sub = `<code>${esc(n.value)}</code>`;
  if (n.level === 'indicador') sub = `<span class="code">${esc(n.value)}</span>`;
  if (p && p.note) sub += `<span>${esc(p.note)}</span>`;
  if (n.children.length || n.rows.length > 1) sub += `<span>${plural(n.rows.length, 'variable', 'variables')}${nVis !== n.rows.length ? ` (${nVis} con los filtros)` : ''}</span>`;
  if (n.rows.length === 1) sub += `<span class="muted">Fila ${n.rows[0]._excelRow} del Excel</span>`;

  let body = sectionsFor(vals, newKeys);
  if (inhKeys.size) {
    const inner = sectionsFor(vals, inhKeys);
    const names = SECTIONS.filter(([sec]) => FIELDS.some(f => f.sec === sec && inhKeys.has(f.k))).map(x => x[1].toLowerCase());
    if (inner) body += `<details class="inh"><summary>Contexto compartido con el nivel superior: ${esc(names.join(', '))}</summary>${inner}</details>`;
  }
  // Hoja con varias filas que no se diferencian por el árbol
  if (!n.children.length && n.rows.length > 1) {
    body += n.rows.map(r => {
      const keys = new Set(FIELDS.map(f => f.k).filter(k => !(k in vals)));
      return `<div class="rowblock"><p class="muted">Fila ${r._excelRow} del Excel</p>${sectionsFor(r, keys)}</div>`;
    }).join('');
  }

  el.innerHTML = `
    <nav class="crumbs" aria-label="Ruta">${crumbs}</nav>
    <div class="lvl" style="${n.phase >= 0 ? `--phc: var(--ph${n.phase})` : ''}"><i></i>${esc(lv.name)}</div>
    <h1 class="title">${esc(p ? `${p.code} ${p.name}` : n.label)}</h1>
    <div class="subline">${sub}</div>
    ${body}
    ${kidsHtml(n)}`;
  el.scrollTop = 0;
}

function kidsHtml(n) {
  const kids = n.children.filter(c => visCount(c));
  if (!kids.length) return '';
  const lv = LEVELS[kids[0].level];
  return `<section class="kids"><h2>${esc(lv.plural)} <span>${kids.length}</span></h2>
    <ul class="kid-list">${kids.map(c => {
      const ob = Obs.countOpen(c);
      const nv = visCount(c);
      let small = '';
      if (c.level === 'variable') small = `<small><code>${esc(c.value)}</code> ${esc(c.rows[0].var_desc)}</small>`;
      else if (c.level === 'indicador') small = `<small>${esc(c.rows[0].ind_desc)}</small>`;
      else if (c.level === 'unidad') small = `<small>${esc(c.rows[0].entidad)}</small>`;
      return `<li><button class="kid" data-go="${esc(c.key)}" style="${c.phase >= 0 ? `--phc: var(--ph${c.phase})` : ''}">
        <span class="bar"></span>
        <span class="kl"><strong>${c.level === 'indicador' ? `<span class="code">${esc(c.code)}</span> ` : ''}${esc(c.label)}</strong>${small}</span>
        <span class="km">${c.children.length ? plural(nv, 'variable', 'variables') : ''}${ob ? `<span class="ob">${ob}</span>` : ''}<span class="chev">${ICON.chev}</span></span>
      </button></li>`;
    }).join('')}</ul></section>`;
}

function homeHtml() {
  const count = k => new Set(S.rows.map(r => r[k])).size;
  const inds = [];
  const seen = new Set();
  for (const r of S.rows) if (!seen.has(r.ind_cod)) { seen.add(r.ind_cod); inds.push(r); }
  return `
    <div class="lvl">${esc(S.meta.sheet || 'Matriz')}</div>
    <h1 class="title">${esc(S.meta.title ? S.meta.title.split('—')[0].trim() : 'Matriz de monitoreo')}</h1>
    ${S.meta.title.includes('—') ? `<p class="subline">${esc(S.meta.title.split('—').slice(1).join('—').trim())}</p>` : ''}
    <div class="home-stats">
      <div><b>${S.phases.length}</b><span>fases</span></div>
      <div><b>${count('componente')}</b><span>componentes</span></div>
      <div><b>${count('unidad')}</b><span>unidades de monitoreo</span></div>
      <div><b>${count('ind_cod')}</b><span>indicadores y grupos</span></div>
      <div><b>${S.rows.length}</b><span>variables</span></div>
    </div>
    ${S.meta.note ? `<p class="lead">${esc(S.meta.note)}</p>` : ''}
    ${kidsHtml(S.tree)}
    <h2 class="block">Indicadores</h2>
    <div class="ind-grid">${inds.map(r => `<button class="ind-card" data-ind="${esc(r.ind_cod)}">
      <span class="code">${esc(r.ind_cod)}</span><span>${esc(r.ind_nombre)}</span>
      <small>${plural(S.rows.filter(x => x.ind_cod === r.ind_cod).length, 'variable', 'variables')}</small></button>`).join('')}</div>
    ${S.entities.length ? `
      <h2 class="block">Entidades SIG (capas y tablas)</h2>
      <div class="tbl-wrap"><table class="ent"><thead><tr><th>Código</th><th>Entidad</th><th>Geometría / tabla</th><th>Llaves</th><th>Fuente</th></tr></thead>
      <tbody>${S.entities.map(e => `<tr><td><code>${esc(e.cod)}</code></td><td>${esc(e.entidad)}</td><td>${esc(e.geom)}</td><td>${esc(e.llaves)}</td><td>${esc(e.fuente)}</td></tr>`).join('')}</tbody></table></div>
      ${S.entityNotes.length ? `<div class="notes">${S.entityNotes.map(t => `<p>${esc(t)}</p>`).join('')}</div>` : ''}` : ''}`;
}

/* ============================================================
   Observaciones (almacenamiento local + GitHub)
   ============================================================ */
const Obs = {
  KEY: 'restaura.obs.v1',
  list: store.get('restaura.obs.v1', []),
  save() { store.set(this.KEY, this.list); },
  active() { return this.list.filter(o => !o.eliminado); },
  pending() { return this.list.filter(o => o._pend).length; },
  countOpen(node) { return this.active().filter(o => o.estado !== 'resuelta' && within(o.target, node.cons)).length; },
  forNode(node, deep) {
    const a = this.active();
    if (node === S.tree) return a;
    return deep ? a.filter(o => within(o.target, node.cons)) : a.filter(o => consKey(o.target) === node.key);
  },
  add(node, data) {
    const v = constFields(node);
    const o = {
      id: uid(), target: { ...node.cons }, nivel: node.level, ruta: pathText(node),
      ctx: { fase: v.fase || '', componente: v.componente || '', unidad: v.unidad || '', ind_cod: v.ind_cod || '', ind_nombre: v.ind_nombre || '', var_cod: v.var_cod || '', var_nombre: v.var_nombre || '' },
      campo: data.campo || '', tipo: data.tipo, texto: data.texto, autor: data.autor,
      estado: 'abierta', creado: nowIso(), actualizado: nowIso(), eliminado: false, _pend: true,
    };
    this.list.push(o); this.changed();
    return o;
  },
  update(id, patch) {
    const o = this.list.find(x => x.id === id); if (!o) return;
    Object.assign(o, patch, { actualizado: nowIso(), _pend: true });
    this.changed();
  },
  changed() { this.save(); refreshObsViews(); Sync.schedule(); },
  merge(remote, markPending = false) {
    const map = new Map(this.list.map(o => [o.id, o]));
    let added = 0;
    for (const r of remote || []) {
      if (!r || !r.id || !r.target) continue;
      const l = map.get(r.id);
      if (!l) { const n = { ...r, _pend: markPending }; this.list.push(n); map.set(r.id, n); added++; }
      else if ((r.actualizado || '') > (l.actualizado || '')) { Object.assign(l, r, { _pend: markPending }); added++; }
      else if (r.actualizado === l.actualizado && !markPending) l._pend = false;
    }
    this.save();
    return added;
  },
  clean() { return this.list.map(({ _pend, ...o }) => o); },
};
function pathText(node) {
  return [...ancestors(node).filter(a => a !== S.tree), node].map(a => a.code ? `${a.code} ${a.label}` : a.label).join(' / ');
}
function autor() { return store.get('restaura.autor', ''); }

function renderObs() {
  const n = S.sel || S.tree;
  const isRoot = n === S.tree;
  const list = Obs.forNode(n, S.obsDeep).sort((a, b) => (b.creado || '').localeCompare(a.creado || ''));
  const vals = isRoot ? {} : constFields(n);
  const campos = FIELDS.filter(f => f.k in vals && vals[f.k] && vals[f.k] !== '—');
  const who = autor();

  const form = isRoot ? `<p class="hint" style="margin-bottom:12px">Elige una fase, unidad, indicador o variable en el árbol para anotar sobre ese elemento. Aquí ves todas las observaciones.</p>` : `
    <form class="obs-form" id="obsForm">
      ${who ? '' : '<label>Tu nombre<input name="autor" required placeholder="Para firmar tus observaciones" autocomplete="name"></label>'}
      <div class="two">
        <label>Tipo<select name="tipo">${TIPOS.map(([k, l]) => `<option value="${k}">${l}</option>`).join('')}</select></label>
        <label>Sobre<select name="campo"><option value="">Todo el elemento</option>${campos.map(f => `<option value="${f.k}" ${S.campo === f.k ? 'selected' : ''}>${esc(FIELD[f.k].label)}</option>`).join('')}</select></label>
      </div>
      <label>Observación<textarea name="texto" required placeholder="Escribe la observación…"></textarea></label>
      <div class="actions">
        <span class="hint">${who ? `Firmas como ${esc(who)}` : ''}</span>
        <button class="btn primary" type="submit">Guardar observación</button>
      </div>
    </form>`;

  $('#obsPane').innerHTML = `
    <div class="obs-head"><h2>Observaciones</h2><p>${isRoot ? 'De toda la matriz' : esc(n.code ? `${n.code} ${n.label}` : n.label)}</p></div>
    ${form}
    ${isRoot || !n.children.length ? '' : `<div class="obs-tools"><label><input type="checkbox" id="obsDeep" ${S.obsDeep ? 'checked' : ''}> Incluir las de los elementos internos</label></div>`}
    <div class="obs-list">${list.length ? list.map(o => obsHtml(o, n)).join('') : `<p class="empty">${isRoot ? 'Todavía no hay observaciones.' : 'Sin observaciones en este elemento.'}</p>`}</div>`;
}

function obsHtml(o, node) {
  const direct = consKey(o.target) === node.key;
  const editing = S.editing === o.id;
  return `<article class="obs t-${esc(o.tipo)}${o.estado === 'resuelta' ? ' resuelta' : ''}">
    <div class="meta"><span class="tipo">${esc(TIPO[o.tipo] || o.tipo)}</span><span class="who">${esc(o.autor || 'Sin nombre')}</span><time datetime="${esc(o.creado)}">${fmtDate(o.creado)}</time>${o.estado === 'resuelta' ? '<span>Resuelta</span>' : ''}</div>
    ${direct ? '' : `<button class="linkish tgt" data-target="${esc(o.id)}">${esc(o.ruta || consKey(o.target))}</button>`}
    ${o.campo ? `<div class="campo">Sobre: ${esc(FIELD[o.campo]?.label || o.campo)}</div>` : ''}
    ${editing
      ? `<textarea data-edit-text="${esc(o.id)}">${esc(o.texto)}</textarea>
         <div class="acts"><button data-edit-save="${esc(o.id)}">Guardar cambios</button><button data-edit-cancel>Cancelar</button></div>`
      : `<p class="txt">${esc(o.texto)}</p>
         <div class="acts">
           <button data-res="${esc(o.id)}">${o.estado === 'resuelta' ? 'Reabrir' : 'Marcar resuelta'}</button>
           <button data-edit="${esc(o.id)}">Editar</button>
           <button class="del" data-del="${esc(o.id)}">Eliminar</button>
           ${o._pend && GH.canWrite() ? '<span class="pend">Sin subir</span>' : ''}
         </div>`}
  </article>`;
}

function refreshObsViews() {
  $('#obsTotal').textContent = Obs.active().filter(o => o.estado !== 'resuelta').length;
  if (S.filters.conObs) { computeVisible(); renderFilters(); renderRibbon(); }
  Sync.paint();
  if (!S.tree) return;
  renderTree(); renderObs();
  if (S.sel) renderDetailKidsOnly();
  if ($('#dlgAll').open) renderAll();
  Sync.paint();
}
function renderDetailKidsOnly() {
  const old = $('#detail .kids');
  const n = S.sel || S.tree;
  if (old) { const tmp = document.createElement('div'); tmp.innerHTML = kidsHtml(n); old.replaceWith(tmp.firstElementChild || document.createTextNode('')); }
}

/* ---------- GitHub ---------- */
const GH = {
  cfg() {
    const saved = store.get('restaura.gh', {});
    const c = { owner: CFG.owner, repo: CFG.repo, branch: CFG.branch || 'main', obsPath: CFG.obsPath };
    if (!c.owner && /\.github\.io$/i.test(location.hostname)) {
      c.owner = location.hostname.split('.')[0];
      const seg = location.pathname.split('/').filter(Boolean)[0];
      c.repo = c.repo || (seg && !seg.includes('.') ? seg : `${c.owner}.github.io`);
    }
    for (const k of ['owner', 'repo', 'branch', 'obsPath']) if (saved[k]) c[k] = saved[k];
    c.token = store.get('restaura.token', '');
    return c;
  },
  canRead() { const c = this.cfg(); return !!(c.owner && c.repo); },
  canWrite() { return this.canRead() && !!this.cfg().token; },
  url(c) { return `https://api.github.com/repos/${encodeURIComponent(c.owner)}/${encodeURIComponent(c.repo)}/contents/${c.obsPath.split('/').map(encodeURIComponent).join('/')}`; },
  headers(c) {
    const h = { Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
    if (c.token) h.Authorization = `Bearer ${c.token}`;
    return h;
  },
  async get() {
    const c = this.cfg();
    const res = await fetch(`${this.url(c)}?ref=${encodeURIComponent(c.branch)}&t=${Date.now()}`, { headers: this.headers(c), cache: 'no-store' });
    if (res.status === 404) return { sha: null, data: { observaciones: [] } };
    if (!res.ok) throw new Error(await ghError(res));
    const j = await res.json();
    let text;
    if (j.content) text = b64dec(j.content);
    else if (j.download_url) text = await (await fetch(j.download_url, { cache: 'no-store' })).text();
    else text = '{}';
    let data;
    try { data = JSON.parse(text || '{}'); } catch { throw new Error('El archivo de observaciones del repositorio no es un JSON válido.'); }
    if (Array.isArray(data)) data = { observaciones: data };
    return { sha: j.sha, data };
  },
  async put(list, sha, n) {
    const c = this.cfg();
    const body = {
      message: `Observaciones: ${plural(n, 'cambio', 'cambios')}${autor() ? ` (${autor()})` : ''}`,
      content: b64enc(JSON.stringify({ esquema: 1, actualizado: nowIso(), observaciones: list }, null, 2) + '\n'),
      branch: c.branch,
    };
    if (sha) body.sha = sha;
    const res = await fetch(this.url(c), { method: 'PUT', headers: { ...this.headers(c), 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    if (res.status === 409 || res.status === 422) { const e = new Error('conflicto'); e.conflict = true; throw e; }
    if (!res.ok) throw new Error(await ghError(res));
    return res.json();
  },
};
async function ghError(res) {
  let msg = '';
  try { msg = (await res.json()).message || ''; } catch { /* nada */ }
  if (res.status === 401) return 'GitHub rechazó el token (401). Revisa que esté vigente.';
  if (res.status === 403) return /rate limit/i.test(msg) ? 'Se alcanzó el límite de consultas a GitHub; vuelve a intentar en unos minutos o agrega un token.' : 'El token no tiene permiso de escritura sobre este repositorio (403).';
  if (res.status === 404) return 'No encontré el repositorio o la rama. Revisa la configuración.';
  return `GitHub respondió ${res.status}${msg ? `: ${msg}` : ''}`;
}
function b64enc(s) {
  const bytes = new TextEncoder().encode(s); let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}
function b64dec(b) {
  const bin = atob(b.replace(/\s/g, ''));
  return new TextDecoder().decode(Uint8Array.from(bin, ch => ch.charCodeAt(0)));
}

const Sync = {
  state: 'local', msg: '', last: store.get('restaura.lastSync', null), timer: null, running: false, again: false,
  schedule() { if (!GH.canWrite()) { this.paint(); return; } clearTimeout(this.timer); this.timer = setTimeout(() => this.run(), 1500); this.paint(); },
  async run(manual = false) {
    if (!GH.canRead()) { this.paint(); if (manual) openSettings(); return; }
    if (this.running) { this.again = true; return; }
    this.running = true; this.state = 'busy'; this.paint();
    try {
      for (let attempt = 0; attempt < 4; attempt++) {
        const { sha, data } = await GH.get();
        Obs.merge(data.observaciones);
        const pend = Obs.list.filter(o => o._pend);
        if (!pend.length || !GH.canWrite()) break;
        const snapshot = new Map(pend.map(o => [o.id, o.actualizado]));
        try {
          await GH.put(Obs.clean(), sha, pend.length);
          Obs.list.forEach(o => { if (snapshot.get(o.id) === o.actualizado) o._pend = false; });
          Obs.save();
          break;
        } catch (e) {
          if (!e.conflict || attempt === 3) throw e;
        }
      }
      this.last = nowIso(); store.set('restaura.lastSync', this.last);
      this.state = 'ok'; this.msg = '';
    } catch (e) {
      this.state = 'error'; this.msg = e.message || String(e);
      if (manual) toast(this.msg, 5000);
    } finally {
      this.running = false;
      refreshObsViews();
      if (this.again) { this.again = false; this.run(); }
    }
  },
  paint() {
    const b = $('#btnSync'); const t = $('#syncTxt');
    const pend = Obs.pending();
    let st = this.state, txt;
    if (!GH.canRead()) { st = 'local'; txt = 'Solo en este navegador'; }
    else if (st === 'busy') txt = 'Sincronizando…';
    else if (st === 'error') txt = 'Error al sincronizar';
    else if (!GH.canWrite()) { st = 'readonly'; txt = pend ? `${plural(pend, 'nota local', 'notas locales')}, sin token` : 'GitHub: solo lectura'; }
    else if (pend) { st = 'pending'; txt = `${plural(pend, 'cambio', 'cambios')} sin subir`; }
    else txt = this.last ? `Sincronizado ${new Date(this.last).toLocaleTimeString('es-CO', { hour: '2-digit', minute: '2-digit' })}` : 'Sincronizado';
    b.dataset.state = st; t.textContent = txt;
    b.title = this.msg || (GH.canRead() ? 'Actualizar desde GitHub' : 'Configurar el guardado en GitHub');
  },
};

/* ============================================================
   Todas las observaciones + exportación
   ============================================================ */
function exportRows() {
  return Obs.active().sort((a, b) => (a.creado || '').localeCompare(b.creado || '')).map(o => {
    const c = o.ctx || {}; const t = o.target || {};
    return {
      'ID': o.id, 'Creada': o.creado, 'Actualizada': o.actualizado, 'Autor': o.autor, 'Tipo': TIPO[o.tipo] || o.tipo,
      'Estado': o.estado === 'resuelta' ? 'Resuelta' : 'Abierta', 'Nivel': LEVELS[o.nivel]?.name || o.nivel || '',
      'Fase': t.fase || c.fase || '', 'Componente': t.componente || c.componente || '', 'Unidad de monitoreo': t.unidad || c.unidad || '',
      'Indicador (cód.)': t.ind_cod || c.ind_cod || '', 'Indicador': c.ind_nombre || '', 'Variable (cód.)': t.var_cod || c.var_cod || '',
      'Variable': c.var_nombre || '', 'Campo': o.campo ? (FIELD[o.campo]?.label || o.campo) : 'Todo el elemento',
      'Ruta': o.ruta || '', 'Observación': o.texto,
    };
  });
}
const stamp = () => new Date().toISOString().slice(0, 10).replace(/-/g, '');
function download(name, blob) {
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
function expXlsx() {
  const rows = exportRows(); if (!rows.length) return toast('No hay observaciones para descargar.');
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = Object.keys(rows[0]).map(k => ({ wch: k === 'Observación' ? 70 : k === 'Ruta' ? 50 : k === 'ID' ? 12 : 22 }));
  ws['!autofilter'] = { ref: ws['!ref'] };
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, ws, 'Observaciones');
  XLSX.writeFile(wb, `observaciones_RESTAURA_${stamp()}.xlsx`);
}
function expCsv() {
  const rows = exportRows(); if (!rows.length) return toast('No hay observaciones para descargar.');
  const keys = Object.keys(rows[0]);
  const q = v => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csv = [keys.map(q).join(';'), ...rows.map(r => keys.map(k => q(r[k])).join(';'))].join('\r\n');
  download(`observaciones_RESTAURA_${stamp()}.csv`, new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' }));
}
function expJson() {
  download(`observaciones_RESTAURA_${stamp()}.json`, new Blob([JSON.stringify({ esquema: 1, actualizado: nowIso(), observaciones: Obs.clean() }, null, 2)], { type: 'application/json' }));
}

function renderAll() {
  const q = norm($('#allQ').value); const est = $('#allEstado').value; const tip = $('#allTipo').value;
  const list = Obs.active().filter(o => (!est || (o.estado || 'abierta') === est) && (!tip || o.tipo === tip)
    && (!q || norm([o.texto, o.autor, o.ruta, o.campo && FIELD[o.campo]?.label].join(' ')).includes(q)))
    .sort((a, b) => (b.creado || '').localeCompare(a.creado || ''));
  $('#allTable').innerHTML = list.length ? `<thead><tr><th>Fecha</th><th>Autor</th><th>Tipo</th><th>Elemento</th><th>Sobre</th><th>Observación</th><th>Estado</th></tr></thead>
    <tbody>${list.map(o => `<tr class="${o.estado === 'resuelta' ? 'res' : ''}">
      <td>${fmtDate(o.creado)}</td><td>${esc(o.autor)}</td><td>${esc(TIPO[o.tipo] || o.tipo)}</td>
      <td><button data-target="${esc(o.id)}">${esc(o.ruta || consKey(o.target))}</button></td>
      <td>${esc(o.campo ? FIELD[o.campo]?.label || o.campo : 'Todo el elemento')}</td>
      <td class="t">${esc(o.texto)}</td><td>${o.estado === 'resuelta' ? 'Resuelta' : 'Abierta'}</td></tr>`).join('')}</tbody>`
    : '<tbody><tr><td class="empty">No hay observaciones con esos criterios.</td></tr></tbody>';
  const c = GH.cfg();
  const lnk = $('#lnkGhFile');
  if (GH.canRead()) { lnk.hidden = false; lnk.href = `https://github.com/${c.owner}/${c.repo}/blob/${c.branch}/${c.obsPath}`; } else lnk.hidden = true;
}

/* ============================================================
   Navegación
   ============================================================ */
function select(node, { push = true } = {}) {
  S.sel = node && node !== S.tree ? node : null;
  S.campo = ''; S.editing = null;
  if (S.sel) ancestors(S.sel).forEach(a => a.key && S.expanded.add(a.key));
  renderRibbon(); renderTree(); renderDetail(); renderObs();
  if (push) {
    const h = `#v=${S.view.id}${S.sel ? `&n=${encodeURIComponent(S.sel.key)}` : ''}`;
    if (location.hash !== h) history.pushState(null, '', h);
  }
  document.body.classList.remove('tree-open'); $('#scrim').hidden = true;
  $('#treeToggle').setAttribute('aria-expanded', 'false');
}
function setView(id, targetKey) {
  const v = VIEWS.find(x => x.id === id) || VIEWS[0];
  const prev = S.sel ? { ...S.sel.cons } : null;
  S.view = v; $('#viewSel').value = v.id;
  S.tree = buildTree(v);
  let node = S.tree;
  if (targetKey !== undefined) node = S.index.get(targetKey) || S.tree;
  else if (prev) node = nodeForTarget(prev);
  select(node);
}
function fromHash() {
  const p = new URLSearchParams(location.hash.slice(1));
  const v = p.get('v') || S.view.id; const n = p.get('n') || '';
  if (v !== S.view.id || !S.tree) { S.view = VIEWS.find(x => x.id === v) || VIEWS[0]; $('#viewSel').value = S.view.id; S.tree = buildTree(S.view); }
  select(S.index.get(n) || S.tree, { push: false });
}
function rerenderAll() {
  computeVisible(); renderFilters(); renderRibbon(); renderTree(); renderDetail(); renderObs();
}

/* ============================================================
   Carga del Excel
   ============================================================ */
async function loadExcelFromRepo() {
  const res = await fetch(`${CFG.excel}?t=${Date.now()}`, { cache: 'no-store' });
  if (!res.ok) throw new Error(`No pude descargar ${CFG.excel} (${res.status}).`);
  const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
  parseWorkbook(wb, `${CFG.excel} (repositorio)`);
}
async function loadExcelFile(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
  parseWorkbook(wb, `${file.name} (archivo local, no se sube al repositorio)`);
  start();
  toast(`Matriz cargada desde ${file.name}`);
}
function showLoadError(err) {
  $('#detail').innerHTML = `<div class="drop" id="drop">
    <h2>No se pudo cargar la matriz</h2>
    <p class="err">${esc(err.message || err)}</p>
    <p>Si abriste el archivo <code>index.html</code> directamente desde tu computador, el navegador bloquea la lectura del Excel. Publica la carpeta en GitHub Pages, o arrastra aquí el Excel para verlo ahora.</p>
    <button class="btn primary" id="btnPick">Elegir el Excel</button></div>`;
  const d = $('#drop');
  d.addEventListener('dragover', e => { e.preventDefault(); d.classList.add('over'); });
  d.addEventListener('dragleave', () => d.classList.remove('over'));
  d.addEventListener('drop', e => { e.preventDefault(); d.classList.remove('over'); const f = e.dataTransfer.files[0]; if (f) loadExcelFile(f).catch(showLoadError); });
  $('#btnPick').onclick = () => $('#fileExcel').click();
}
function start() {
  const doc = S.meta.title ? S.meta.title.split('—')[0].trim() : '';
  $('#docTitle').textContent = doc || 'Monitoreo de la restauración';
  computeVisible(); renderFilters();
  S.tree = null;
  fromHash();
  refreshObsViews();
}

/* ============================================================
   Configuración
   ============================================================ */
function openSettings() {
  const f = $('#frmSettings'); const c = GH.cfg();
  f.autor.value = autor(); f.owner.value = c.owner || ''; f.repo.value = c.repo || '';
  f.branch.value = c.branch || 'main'; f.obsPath.value = c.obsPath || ''; f.token.value = c.token || '';
  $('#setStatus').textContent = GH.canWrite() ? 'Hay un token guardado en este navegador.' : '';
  $('#excelSource').textContent = `Excel en uso: ${S.meta.source || '—'}`;
  $('#dlgSettings').showModal();
}
function saveSettings() {
  const f = $('#frmSettings');
  store.set('restaura.autor', f.autor.value.trim());
  store.set('restaura.gh', { owner: f.owner.value.trim(), repo: f.repo.value.trim(), branch: f.branch.value.trim() || 'main', obsPath: f.obsPath.value.trim() || CFG.obsPath });
  const tk = f.token.value.trim();
  if (tk) store.set('restaura.token', tk); else store.del('restaura.token');
  Sync.paint(); renderObs();
}

/* ============================================================
   Eventos
   ============================================================ */
function bind() {
  document.addEventListener('click', e => {
    const t = e.target.closest('button, a'); if (!t) return;
    if (t.dataset.go !== undefined) { const n = S.index.get(t.dataset.go); if (n) select(n); return; }
    if (t.dataset.tog !== undefined) {
      const k = t.dataset.tog;
      if (S.expanded.has(k)) S.expanded.delete(k); else S.expanded.add(k);
      store.set('restaura.exp', [...S.expanded]); renderTree(); return;
    }
    if (t.dataset.phase) {
      const key = `fase=${t.dataset.phase}`;
      if (S.view.id !== 'fase') setView('fase', key); else select(S.index.get(key) || S.tree);
      return;
    }
    if (t.dataset.ind) { setView('indicador', `ind_cod=${t.dataset.ind}`); return; }
    if (t.dataset.niv) {
      const n = t.dataset.niv; const set = S.filters.niveles;
      set.has(n) ? set.delete(n) : set.add(n); rerenderAll(); return;
    }
    if (t.id === 'fClear') { S.filters = { q: '', niveles: new Set(), momento: '', oblig: '', conObs: false }; $('#q').value = ''; rerenderAll(); return; }
    if (t.dataset.campo) {
      S.campo = t.dataset.campo; renderObs();
      const ta = $('#obsForm textarea');
      if (ta) { ta.focus({ preventScroll: false }); ta.scrollIntoView({ block: 'center', behavior: 'smooth' }); }
      return;
    }
    if (t.dataset.target) {
      const o = Obs.list.find(x => x.id === t.dataset.target);
      if (o) { $('#dlgAll').close(); select(nodeForTarget(o.target)); }
      return;
    }
    if (t.dataset.res) { const o = Obs.list.find(x => x.id === t.dataset.res); Obs.update(o.id, { estado: o.estado === 'resuelta' ? 'abierta' : 'resuelta' }); return; }
    if (t.dataset.edit) { S.editing = t.dataset.edit; renderObs(); const ta = $(`[data-edit-text="${CSS.escape(S.editing)}"]`); ta && ta.focus(); return; }
    if (t.dataset.editSave) {
      const ta = $(`[data-edit-text="${CSS.escape(t.dataset.editSave)}"]`);
      const txt = ta.value.trim(); if (!txt) return;
      S.editing = null; Obs.update(t.dataset.editSave, { texto: txt }); toast('Observación actualizada'); return;
    }
    if (t.hasAttribute('data-edit-cancel')) { S.editing = null; renderObs(); return; }
    if (t.dataset.del) {
      if (!confirm('¿Eliminar esta observación?')) return;
      Obs.update(t.dataset.del, { eliminado: true }); toast('Observación eliminada'); return;
    }
  });

  document.addEventListener('submit', e => {
    if (e.target.id !== 'obsForm') return;
    e.preventDefault();
    const f = e.target;
    if (f.autor) { const a = f.autor.value.trim(); if (!a) return; store.set('restaura.autor', a); }
    const texto = f.texto.value.trim(); if (!texto) return;
    const node = S.sel; if (!node) return;
    Obs.add(node, { tipo: f.tipo.value, campo: f.campo.value, texto, autor: autor() });
    S.campo = '';
    toast(GH.canWrite() ? 'Observación guardada; subiendo a GitHub…' : 'Observación guardada en este navegador');
  });

  document.addEventListener('change', e => {
    const t = e.target;
    if (t.id === 'fMom') { S.filters.momento = t.value; rerenderAll(); }
    if (t.id === 'fObl') { S.filters.oblig = t.value; rerenderAll(); }
    if (t.id === 'fObs') { S.filters.conObs = t.checked; rerenderAll(); }
    if (t.id === 'obsDeep') { S.obsDeep = t.checked; store.set('restaura.deep', S.obsDeep); renderObs(); }
    if (t.id === 'viewSel') setView(t.value);
    if (t.id === 'allEstado' || t.id === 'allTipo') renderAll();
  });

  let qT;
  $('#q').addEventListener('input', e => { clearTimeout(qT); qT = setTimeout(() => { S.filters.q = e.target.value; rerenderAll(); }, 160); });
  $('#allQ').addEventListener('input', renderAll);
  $('#btnHome').onclick = () => select(S.tree);
  $('#btnSettings').onclick = openSettings;
  $('#btnSync').onclick = () => Sync.run(true);
  $('#btnAllObs').onclick = () => { renderAll(); $('#dlgAll').showModal(); };
  $('#btnCloseAll').onclick = () => $('#dlgAll').close();
  $('#btnSaveSettings').onclick = () => { saveSettings(); $('#dlgSettings').close(); toast('Configuración guardada'); Sync.run(); };
  $('#btnForgetToken').onclick = () => { store.del('restaura.token'); $('#frmSettings').token.value = ''; Sync.paint(); $('#setStatus').textContent = 'Token eliminado de este navegador.'; };
  $('#btnTestGh').onclick = async () => {
    saveSettings(); const s = $('#setStatus');
    if (!GH.canRead()) { s.textContent = 'Indica el usuario y el repositorio.'; return; }
    s.textContent = 'Probando…';
    try {
      const { data } = await GH.get();
      const n = (data.observaciones || []).length;
      s.textContent = `Lectura correcta: el repositorio tiene ${plural(n, 'observación', 'observaciones')}. ${GH.canWrite() ? 'La escritura se probará con la próxima observación.' : 'Sin token, solo lectura.'}`;
    } catch (err) { s.textContent = err.message; }
  };
  $('#btnLoadExcel').onclick = () => $('#fileExcel').click();
  $('#btnReloadExcel').onclick = async () => { try { await loadExcelFromRepo(); start(); toast('Matriz del repositorio cargada'); $('#excelSource').textContent = `Excel en uso: ${S.meta.source}`; } catch (err) { toast(err.message, 5000); } };
  $('#fileExcel').onchange = e => { const f = e.target.files[0]; e.target.value = ''; if (f) loadExcelFile(f).then(() => { $('#excelSource').textContent = `Excel en uso: ${S.meta.source}`; }).catch(err => toast(err.message, 6000)); };
  $('#btnExpXlsx').onclick = expXlsx; $('#btnExpCsv').onclick = expCsv; $('#btnExpJson').onclick = expJson;
  $('#btnImport').onclick = () => $('#fileJson').click();
  $('#fileJson').onchange = async e => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    try {
      let d = JSON.parse(await f.text()); if (Array.isArray(d)) d = { observaciones: d };
      const n = Obs.merge(d.observaciones, true);
      Obs.changed(); toast(`${plural(n, 'observación importada', 'observaciones importadas')}`);
    } catch { toast('El archivo no es un JSON de observaciones válido.', 5000); }
  };
  $('#treeToggle').onclick = () => {
    const open = !document.body.classList.contains('tree-open');
    document.body.classList.toggle('tree-open', open); $('#scrim').hidden = !open;
    $('#treeToggle').setAttribute('aria-expanded', String(open));
  };
  $('#scrim').onclick = () => $('#treeToggle').click();
  window.addEventListener('popstate', fromHash);
  window.addEventListener('hashchange', fromHash);
  document.addEventListener('keydown', e => {
    if (e.key === '/' && !/input|textarea|select/i.test(document.activeElement.tagName)) { e.preventDefault(); $('#q').focus(); }
  });
}

/* ============================================================
   Inicio
   ============================================================ */
async function init() {
  $('#viewSel').innerHTML = VIEWS.map(v => `<option value="${v.id}">${esc(v.name)}</option>`).join('');
  $('#allTipo').innerHTML = '<option value="">Todos los tipos</option>' + TIPOS.map(([k, l]) => `<option value="${k}">${l}</option>`).join('');
  bind();
  Sync.paint();
  if (typeof XLSX === 'undefined') { showLoadError(new Error('No cargó la librería para leer Excel (revisa la conexión a internet).')); return; }
  try { await loadExcelFromRepo(); start(); }
  catch (err) { showLoadError(err); }
  if (GH.canRead()) Sync.run();
}
init();
})();
