/* Spese di Famiglia — logica dell'app */
'use strict';

const APP_VERSION = '1.0.0';
const STORAGE_KEY = 'spese-famiglia-dati';

// ============================================================
// Utilità
// ============================================================
const $ = (sel, root = document) => root.querySelector(sel);
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const compact = s => String(s ?? '').toUpperCase().normalize('NFD').replace(/[^A-Z0-9]/g, '');
const sum = arr => arr.reduce((a, b) => a + b, 0);
const debounce = (fn, ms) => { let t; return (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); }; };

// Importi in centesimi → "1.234,56 €" (il punto delle migliaia anche sotto 10.000)
const groupThousands = n => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
const euro = cents => `${groupThousands(Math.floor(Math.abs(cents || 0) / 100))},${pad(Math.abs(cents || 0) % 100)} €`;
const fmt = cents => (cents < 0 ? '−' : '') + euro(cents);
const fmtSigned = cents => (cents > 0 ? '+' : cents < 0 ? '−' : '') + euro(cents);
const fmtInt = cents => (cents < 0 ? '−' : '') + groupThousands(Math.round(Math.abs(cents) / 100));
// Nei riquadri riassuntivi, sopra i 1.000 € i centesimi si omettono
const fmtTile = (cents, signed) => (Math.abs(cents) < 100000 ? (signed ? fmtSigned(cents) : fmt(cents))
  : (cents < 0 ? '−' : signed && cents > 0 ? '+' : '') + fmtInt(Math.abs(cents)) + ' €');
const pct = (a, b) => (b ? Math.round((a / b) * 100) : 0);
function fmtShort(cents) {
  const v = cents / 100;
  if (Math.abs(v) >= 1000) return (v / 1000).toLocaleString('it-IT', { maximumFractionDigits: Math.abs(v) >= 10000 ? 0 : 1 }) + 'k';
  return String(Math.round(v));
}

// Numeri scritti all'italiana ("1.234,56") o all'inglese ("1234.56")
function parseNumber(v) {
  if (typeof v === 'number') return v;
  let s = String(v ?? '').replace(/[€\s ]/g, '').replace(/^\((.*)\)$/, '-$1').replace(/^−/, '-');
  if (!s) return NaN;
  const comma = s.lastIndexOf(','), dot = s.lastIndexOf('.');
  if (comma > -1 && dot > -1) s = comma > dot ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (comma > -1) s = s.split(',').length > 2 ? s.replace(/,/g, '') : s.replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  const n = Number(s);
  return Number.isFinite(n) ? n : NaN;
}
const toCents = n => Math.round(n * 100);

const MONTHS = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];
const MONTHS_SHORT = ['Gen', 'Feb', 'Mar', 'Apr', 'Mag', 'Giu', 'Lug', 'Ago', 'Set', 'Ott', 'Nov', 'Dic'];
const pad = n => String(n).padStart(2, '0');
const isoLocal = d => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const today = () => isoLocal(new Date());
const yesterday = () => { const d = new Date(); d.setDate(d.getDate() - 1); return isoLocal(d); };
const thisMonth = () => today().slice(0, 7);
function addMonths(ym, n) {
  let y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + n;
  y += Math.floor(m / 12);
  m = ((m % 12) + 12) % 12;
  return `${y}-${pad(m + 1)}`;
}
const monthName = ym => `${MONTHS[+ym.slice(5, 7) - 1]} ${ym.slice(0, 4)}`;
const monthShort = ym => `${MONTHS_SHORT[+ym.slice(5, 7) - 1]} ${ym.slice(2, 4)}`;
const fmtDate = iso => (iso ? iso.split('-').reverse().join('/') : '');
function dayLabel(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  return new Date(y, m - 1, d).toLocaleDateString('it-IT', { weekday: 'long', day: 'numeric', month: 'long' });
}
function parseDate(v) {
  if (typeof v === 'number') return v > 20000 && v < 80000 ? XlsxLite.serialToISO(v) : null;
  const s = String(v ?? '').trim();
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2,4})/);
  if (m) {
    const y = +m[3] < 100 ? 2000 + +m[3] : +m[3], mo = +m[2], d = +m[1];
    if (mo >= 1 && mo <= 12 && d >= 1 && d <= 31) return `${y}-${pad(mo)}-${pad(d)}`;
  }
  if (/^\d{5}(\.\d+)?$/.test(s)) return parseDate(Number(s));
  return null;
}
const titleCase = s => String(s).toLowerCase().replace(/(^|[\s\-/(&.'])(\p{L})/gu, (m, a, b) => a + b.toUpperCase());
const isIOS = /iphone|ipad|ipod/i.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
const isMobile = isIOS || /android|mobile/i.test(navigator.userAgent);
const isStandalone = () => window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;

const ICON = {
  left: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 6l-6 6 6 6"/></svg>',
  right: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>',
  close: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 6l12 12M18 6L6 18"/></svg>',
  warn: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3.5l9.5 17h-19z"/><path d="M12 10v4.5M12 17.5v.5"/></svg>',
  info: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 11v6M12 7.5v.5"/></svg>',
  upload: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 15V4M7 9l5-5 5 5"/><path d="M4 15v4a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-4"/></svg>',
  chev: '<svg viewBox="0 0 24 24" width="18" height="18" aria-hidden="true" style="fill:none;stroke:currentColor;stroke-width:2"><path d="M9 6l6 6-6 6"/></svg>',
};

// ============================================================
// Categorie predefinite (come nel foglio "Spese Mensili")
// ============================================================
const DEFAULT_CATEGORIES = [
  ['casa', 'Casa', [['condominio', 'Spese condominiali'], ['internet', 'Internet'], ['luce', 'Luce'], ['tari', 'Tari'], ['manutenzione', 'Manutenzione'], ['spesa', 'Spesa'], ['mutuo', 'Mutuo Attilio'], ['arredamento', 'Arredamento e Varie'], ['amazon', 'Amazon']]],
  ['trasporti', 'Trasporti', [['auto500', 'Spese Auto 500L'], ['carburante', 'Carburante'], ['ass500', 'Assicurazione Auto 500L'], ['vespa', 'Spese Scooter Vespa/moto'], ['assvespa', 'Assicurazione Scooter Vespa'], ['assmoto', 'Assicurazione Moto'], ['telepass', 'Telepass'], ['multe', 'Multe']]],
  ['cibo', 'Cibo fuori', [['hobby', 'Hobby'], ['ristorante', 'Pizzeria / Ristorante / Bar']]],
  ['salute', 'Salute / Persona', [['farmacia', 'Farmacia / Lenti a contatto'], ['mediche', 'Spese Mediche'], ['parrucchiere', 'Parrucchiere/estetista'], ['abbigliamento', 'Abbigliamento']]],
  ['istruzione', 'Istruzione', [['scuole', 'Scuole bambini'], ['formazione', 'Formazione professionale']]],
  ['altre', 'Altre spese', [['sportbimbi', 'Sport bambini'], ['sportadulti', 'Sport adulti'], ['vacanze', 'Vacanze'], ['regali', 'Regali'], ['beneficenza', 'Beneficenza']]],
];
const DEFAULT_INCOME = [['stip_giacomo', 'Stipendio Giacomo'], ['stip_sara', 'Stipendio Sara'], ['assegno', 'Assegno unico'], ['altre', 'Altre entrate']];

// Regole di classificazione predefinite: [parola chiave, destinazione, dove cercarla]
// dove: m = nel nome dell'esercente, f = in tutta la descrizione, t = prima parola dell'esercente, w = parola intera
const R = (target, mode, ...kws) => kws.map(k => [k, target, mode]);
const BUILTIN_RULES = [
  ...R('ign', 'f', 'PRELIEVO', 'PRELEVAMENTO', 'GIROCONTO', 'GIROFONDI'),
  ...R('exp:casa:spesa', 'm', 'ESSELUNGA', 'TIGROS', 'CARREFOUR', 'CRFEXP', 'CRFMARKET', 'LIDL', 'CONAD', 'IPERCOOP', 'PENNYMARKET', 'EUROSPIN', 'BENNET', 'IPERAL', 'PAMCITY', 'PAMLOCAL', 'PAMPANORAMA', 'PANORAMA', 'MDDISCOUNT', 'NATURASI', 'SUPERMERCAT', 'PANIFICIO', 'MACELLERIA', 'FRUTTAEVERDURA', 'DESPAR', 'INTERSPAR', 'EUROSPAR', 'TODIS', 'FAMILA', 'GALLERIAALIMENTARE'),
  ...R('exp:casa:spesa', 't', 'ALDI', 'UNES', 'COOP', 'CRAI', 'SIGMA', 'PRIX', 'SPAR', 'PAM'),
  ...R('exp:casa:luce', 'm', 'ENELENERGIA', 'EDISON', 'SORGENIA', 'PLENITUDE', 'OCTOPUS', 'ENERGIA'),
  ...R('exp:casa:luce', 't', 'A2A', 'ENEL', 'IREN', 'HERA', 'ACEA'),
  ...R('exp:casa:internet', 'm', 'FASTWEB', 'VODAFONE', 'ILIAD', 'WINDTRE', 'TELECOMITALIA', 'VERYMOBILE', 'HOMOBILE', 'KENAMOBILE', 'POSTEMOBILE', 'TISCALI', 'EOLO'),
  ...R('exp:casa:internet', 't', 'TIM'),
  ...R('exp:casa:tari', 'w', 'TARI'),
  ...R('exp:casa:tari', 'f', 'TASSARIFIUTI'),
  ...R('exp:casa:condominio', 'f', 'CONDOMIN'),
  ...R('exp:casa:mutuo', 'f', 'MUTUO'),
  ...R('exp:casa:amazon', 'f', 'AMAZON', 'AMZN'),
  ...R('exp:casa:manutenzione', 'm', 'BRICOCENTER', 'LEROYMERLIN', 'BRICOMAN', 'BRICOIO', 'BRICOFER', 'FERRAMENTA', 'IDRAULIC', 'ELETTRICIST'),
  ...R('exp:casa:manutenzione', 't', 'OBI'),
  ...R('exp:casa:arredamento', 'm', 'IKEA', 'JYSK', 'MAISONSDUMONDE', 'FLYINGTIGER', 'MONDOCONVENIENZA', 'MEDIAWORLD', 'UNIEURO', 'EURONICS'),
  ...R('exp:casa:arredamento', 't', 'ACTION'),
  ...R('exp:trasporti:carburante', 't', 'IP', 'ENI', 'Q8', 'ESSO', 'API', 'AGIP', 'TAMOIL', 'SHELL', 'ENILIVE', 'TOTALERG', 'EG'),
  ...R('exp:trasporti:carburante', 'm', 'CARBURANT', 'DISTRIBUTORE', 'STAZIONESERVIZIO'),
  ...R('exp:trasporti:telepass', 'm', 'TELEPASS', 'UNIPOLMOVE', 'UNIPOLTECH', 'ASPIT', 'AUTOSTRAD', 'PEDAGG'),
  ...R('exp:trasporti:multe', 'f', 'VERBALE', 'CONTRAVVENZION', 'POLIZIALOCALE'),
  ...R('exp:trasporti:multe', 'w', 'MULTA'),
  ...R('exp:trasporti:auto500', 'm', 'PARCHEGG', 'EASYPARK', 'AUTOLAVAGGIO', 'GOMMISTA', 'OFFICINA', 'CARROZZERIA', 'REVISIONE'),
  ...R('exp:cibo:ristorante', 't', 'BAR'),
  ...R('exp:cibo:ristorante', 'm', 'RISTORANT', 'PIZZERI', 'TRATTORIA', 'OSTERIA', 'SUSHI', 'MCDONALD', 'BURGERKING', 'ROADHOUSE', 'OLDWILDWEST', 'STARBUCKS', 'CAFFE', 'CAFE', 'GELATERI', 'PASTICCERI', 'JUSTEAT', 'DELIVEROO', 'GLOVO', 'KEBAB', 'PIADIN', 'AUTOGRILL', 'CHEFEXPRESS', 'BISTROT', 'BIRRERIA', 'PANINOTECA'),
  ...R('exp:cibo:hobby', 'm', 'CINEMA', 'THESPACE', 'TEATRO', 'MUSEO', 'TICKETONE', 'VIVATICKET', 'FELTRINELLI', 'MONDADORI', 'LIBRERI', 'NETFLIX', 'SPOTIFY', 'DISNEY', 'PLAYSTATION', 'NINTENDO', 'AUDIBLE', 'PRIMEVIDEO', 'DAZN'),
  ...R('exp:salute:farmacia', 'm', 'FARMACI', 'OTTICA', 'OTTICO', 'SALMOIRAGHI', 'GRANDVISION', 'VISIONOTTICA', 'AFFLELOU'),
  ...R('exp:salute:farmacia', 'w', 'LENTI'),
  ...R('exp:salute:mediche', 'm', 'MEDICAL', 'MEDICO', 'MEDICI', 'POLIAMBULATOR', 'DENTIST', 'ODONTOIATR', 'OSPEDAL', 'SYNLAB', 'FISIOTERAP', 'LABORATORIOANALISI', 'CLINICA'),
  ...R('exp:salute:mediche', 't', 'CDI', 'ASST', 'ATS'),
  ...R('exp:salute:parrucchiere', 'm', 'PARRUCCH', 'HAIR', 'NAILS', 'BEAUTY', 'ESTETIC', 'BARBIER', 'BARBER', 'ACCONCIAT', 'COIFFEUR'),
  ...R('exp:salute:abbigliamento', 'm', 'HENNES', 'BOGGI', 'CALZEDONIA', 'INTIMISSIMI', 'TEZENIS', 'UNIQLO', 'PRIMARK', 'BENETTON', 'ZALANDO', 'PIAZZAITALIA', 'TERRANOVA', 'CARPISA', 'GEOX', 'FOOTLOCKER', 'PITTAROSSO', 'BERSHKA', 'PULLBEAR', 'STRADIVARIUS', 'PRENATAL', 'ORIGINALMARINES', 'ABBIGLIAMENTO', 'CALZATUR'),
  ...R('exp:salute:abbigliamento', 't', 'OVS', 'ZARA', 'COIN', 'CHICCO', 'HM'),
  ...R('exp:istruzione:scuole', 'f', 'SCUOLA', 'SCOLASTIC', 'ISTITUTOCOMPRENSIVO', 'INFANZIA', 'ASILO', 'PAGOINRETE'),
  ...R('exp:istruzione:scuole', 'w', 'RETTA'),
  ...R('exp:istruzione:formazione', 'm', 'UDEMY', 'COURSERA', 'LINKEDINLEARNING'),
  ...R('exp:altre:sportadulti', 'm', 'PALESTRA', 'VIRGINACTIVE', 'MCFIT', 'FITACTIVE', 'GETFIT', 'ANYTIMEFITNESS'),
  ...R('exp:altre:vacanze', 'm', 'BOOKING', 'AIRBNB', 'RYANAIR', 'EASYJET', 'ITAAIRWAYS', 'VUELING', 'WIZZ', 'EXPEDIA', 'EDREAMS', 'VOLAGRATIS', 'LASTMINUTE', 'HOTEL', 'TRENITALIA', 'FLIXBUS', 'GRIMALDI', 'MOBYLINES', 'TIRRENIA'),
  ...R('exp:altre:vacanze', 't', 'ITALO'),
  ...R('exp:altre:beneficenza', 'f', 'ADOZIONE', 'DONAZIONE', 'ONLUS', 'UNICEF', 'EMERGENCY', 'AMNESTY', 'MEDICISENZAFRONTIERE', 'SAVETHECHILDREN', 'ACTIONAID'),
  ...R('exp:altre:beneficenza', 'w', 'ETS'),
  ...R('inc:assegno', 'f', 'ASSEGNOUNICO'),
  ...R('inc:assegno', 'w', 'INPS'),
].sort((a, b) => (a[1] === 'ign' ? -1 : 0) - (b[1] === 'ign' ? -1 : 0) || b[0].length - a[0].length);

// Categoria proposta dalla banca (es. colonna "Moneymap" di Fineco) → nostra sottocategoria
const BANK_CATEGORY_MAP = [
  ['alimentari', 'exp:casa:spesa'], ['supermercat', 'exp:casa:spesa'], ['ristorant', 'exp:cibo:ristorante'],
  ['carburant', 'exp:trasporti:carburante'], ['pedagg', 'exp:trasporti:telepass'], ['trasport', 'exp:trasporti:auto500'], ['auto', 'exp:trasporti:auto500'],
  ['utenze', 'exp:casa:luce'], ['internet', 'exp:casa:internet'], ['telefon', 'exp:casa:internet'], ['mutu', 'exp:casa:mutuo'],
  ['istruzione', 'exp:istruzione:scuole'], ['scuol', 'exp:istruzione:scuole'], ['sport', 'exp:altre:sportbimbi'],
  ['salute', 'exp:salute:mediche'], ['medic', 'exp:salute:mediche'], ['farmac', 'exp:salute:farmacia'], ['bellezza', 'exp:salute:parrucchiere'],
  ['abbigliamento', 'exp:salute:abbigliamento'], ['giochi', 'exp:cibo:hobby'], ['cultura', 'exp:cibo:hobby'], ['spettacol', 'exp:cibo:hobby'], ['hobby', 'exp:cibo:hobby'],
  ['regali', 'exp:altre:regali'], ['donazion', 'exp:altre:beneficenza'], ['beneficenz', 'exp:altre:beneficenza'], ['viagg', 'exp:altre:vacanze'], ['vacanz', 'exp:altre:vacanze'],
  ['casa', 'exp:casa:arredamento'], ['assegn', 'inc:assegno'],
];

// ============================================================
// Dati (salvati nel browser del telefono)
// ============================================================
function defaultState() {
  return {
    version: 1,
    categories: DEFAULT_CATEGORIES.map(([id, name, subs], i) => ({ id, name, slot: i + 1, subs: subs.map(([sid, sname]) => ({ id: sid, name: sname })) })),
    incomeTypes: DEFAULT_INCOME.map(([id, name]) => ({ id, name })),
    entries: [],
    rules: [],
    ignoredKeys: [],
    settings: { mappings: {}, lastBackup: null, installDismissed: false },
  };
}
function migrate(s) {
  const d = defaultState();
  return { ...d, ...s, settings: { ...d.settings, ...(s.settings || {}) } };
}
function load() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) return migrate(JSON.parse(raw));
  } catch (err) { console.error(err); }
  return defaultState();
}
let state = load();
function save() {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(state)); }
  catch (err) { alert('Impossibile salvare i dati: la memoria del browser è piena o non disponibile.'); }
}
let persistAsked = false;
function requestPersist() {
  if (persistAsked) return;
  persistAsked = true;
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});
}

const catById = id => state.categories.find(c => c.id === id);
const subOf = (catId, subId) => catById(catId)?.subs.find(s => s.id === subId);
const incById = id => state.incomeTypes.find(t => t.id === id);
const catColor = cat => `var(--s${cat?.slot || 0})`;
const isUnclassified = e => (e.type === 'exp' ? !subOf(e.cat, e.sub) : !incById(e.inc));
function entryLabel(e) {
  if (e.type === 'inc') return incById(e.inc)?.name || 'Entrata da classificare';
  const c = catById(e.cat), s = subOf(e.cat, e.sub);
  return c && s ? `${c.name} › ${s.name}` : 'Da classificare';
}
const entryTarget = e => (e.type === 'exp' ? (subOf(e.cat, e.sub) ? `exp:${e.cat}:${e.sub}` : '') : (incById(e.inc) ? `inc:${e.inc}` : ''));
const entriesIn = ym => state.entries.filter(e => e.date.startsWith(ym));

// Destinazioni: "exp:categoria:sottocategoria", "inc:voce", "ign" (ignora), "" (da classificare)
function targetValid(t) {
  if (t === 'ign') return true;
  const p = String(t || '').split(':');
  if (p[0] === 'exp') return !!subOf(p[1], p[2]);
  if (p[0] === 'inc') return !!incById(p[1]);
  return false;
}
function targetLabel(t) {
  if (t === 'ign') return 'Ignora (non importare)';
  const p = String(t || '').split(':');
  if (p[0] === 'exp' && subOf(p[1], p[2])) return `${catById(p[1]).name} › ${subOf(p[1], p[2]).name}`;
  if (p[0] === 'inc' && incById(p[1])) return `Entrata › ${incById(p[1]).name}`;
  return 'Da classificare';
}
// bankCents: importo con il segno della banca (negativo = uscita)
function applyTarget(e, target, bankCents) {
  const p = String(target || '').split(':');
  if (p[0] === 'inc') Object.assign(e, { type: 'inc', inc: p[1], cat: null, sub: null, amount: bankCents });
  else if (p[0] === 'exp') Object.assign(e, { type: 'exp', cat: p[1], sub: p[2], inc: null, amount: -bankCents });
  else Object.assign(e, { type: bankCents < 0 ? 'exp' : 'inc', cat: null, sub: null, inc: null, amount: Math.abs(bankCents) });
}
const bankCentsOf = e => (e.type === 'exp' ? -e.amount : e.amount);

function targetOptions(selected, { income = true, ignore = false, refund = false, emptyLabel = '— Da classificare —' } = {}) {
  const opt = (v, label) => `<option value="${esc(v)}"${v === selected ? ' selected' : ''}>${esc(label)}</option>`;
  let h = opt('', emptyLabel);
  if (income) h += `<optgroup label="Entrate">${state.incomeTypes.map(t => opt(`inc:${t.id}`, t.name)).join('')}</optgroup>`;
  for (const c of state.categories) {
    h += `<optgroup label="${esc(c.name + (refund ? ' (rimborso)' : ''))}">${c.subs.map(s => opt(`exp:${c.id}:${s.id}`, s.name)).join('')}</optgroup>`;
  }
  if (ignore) h += `<optgroup label="Altro">${opt('ign', 'Ignora (non importare)')}</optgroup>`;
  return h;
}

// ============================================================
// Lettura delle descrizioni bancarie e classificazione automatica
// ============================================================
function analyzeDescription(short, full) {
  const text = String(full || short || '').replace(/\s+/g, ' ').trim();
  const shortT = String(short || '').replace(/\s+/g, ' ').trim();
  let m, kind = 'other', name = '', extra = '';
  if ((m = text.match(/presso:\s*(.+?)\s+Carta\s*N/i))) { kind = 'pos'; name = m[1]; }
  else if ((m = text.match(/^(.+?)\s+Carta\s*N\.?\s*\*/i))) { kind = 'card'; name = m[1].replace(/\s+[A-Z]{2,3}$/, ''); }
  else if ((m = text.match(/Ben:\s*(.+?)\s+Ins:/i))) { kind = 'bonifico'; name = m[1]; extra = (text.match(/Cau:\s*(.+)$/i) || [])[1] || ''; }
  else if ((m = text.match(/^(.+?)\s+Addebito SDD/i))) { kind = 'sdd'; name = m[1]; extra = (text.match(/\sPer\s+(.+)$/i) || [])[1] || ''; }
  else if ((m = text.match(/Ord(?:inante)?:\s*(.+?)\s+(?:Ins:|Iban|Cau:|Da:|Data)/i))) { kind = 'accredito'; name = m[1]; extra = (text.match(/Cau:\s*(.+)$/i) || [])[1] || ''; }
  else {
    name = shortT && shortT.length < text.length ? shortT : text;
    if (shortT && text !== shortT) extra = text;
  }

  // Nome "pulito" dell'esercente: senza città, indirizzo e numeri
  let base = name;
  if (kind === 'card') {
    const toks = base.trim().split(/\s+/);
    if (toks.length > 1) base = toks.slice(0, -1).join(' ');
  }
  if (kind === 'pos' || kind === 'card') {
    const cut = base.search(/\s(VIA|VIALE|V\.LE|VICOLO|V\.LO|PIAZZA|P\.ZA|PZA|PIAZZALE|CORSO|C\.SO|LARGO|STRADA|LOC\.?|LOCALITA)\s/i);
    if (cut >= 2) base = base.slice(0, cut);
    const digit = base.slice(3).search(/\d/);
    if (digit > -1) base = base.slice(0, digit + 3);
  }
  base = base.replace(/[\s\-,.]+$/, '').trim() || name;
  let key = compact(base).slice(0, 24);
  if (key.length < 3) key = compact(name).slice(0, 24);

  const display = kind === 'pos' ? base : name;
  let title = titleCase(display);
  if (kind === 'bonifico') title = 'Bonifico a ' + title;
  if (kind === 'accredito') title = 'Accredito da ' + title;
  return {
    kind, key, title,
    extra: extra.length > 140 ? extra.slice(0, 137) + '…' : extra,
    merchantC: compact(base),
    fullC: compact(shortT + ' ' + text),
    upper: (shortT + ' ' + text).toUpperCase(),
    firstTok: (base.trim().split(/[\s\-.,/'*]+/)[0] || '').replace(/&/g, '').toUpperCase(),
  };
}

function matchRules(info, bankCents) {
  const fits = t => targetValid(t) && !(t.startsWith('inc:') && bankCents < 0);
  // 1) Regole imparate o create da te: vince la più specifica
  for (const r of [...state.rules].sort((a, b) => b.key.length - a.key.length)) {
    if (!fits(r.target)) continue;
    const k = r.key;
    if (info.key.startsWith(k) || (k.length >= 4 && info.merchantC.includes(k)) || (k.length >= 5 && info.fullC.includes(k))) return { target: r.target, via: 'regola' };
  }
  // 2) Regole predefinite
  for (const [kw, target, mode] of BUILTIN_RULES) {
    if (!fits(target)) continue;
    const hit = mode === 'm' ? info.merchantC.includes(kw)
      : mode === 'f' ? info.fullC.includes(kw)
      : mode === 't' ? info.firstTok === kw
      : new RegExp(`\\b${kw}\\b`).test(info.upper);
    if (hit) return { target, via: 'auto' };
  }
  return null;
}
function suggestFromBank(bankCat, bankCents) {
  const s = String(bankCat || '').toLowerCase();
  if (!s) return null;
  for (const [kw, target] of BANK_CATEGORY_MAP) {
    if (!s.includes(kw) || !targetValid(target)) continue;
    if ((target.startsWith('inc:') && bankCents < 0) || (target.startsWith('exp:') && bankCents > 0)) continue;
    return { target, via: 'banca' };
  }
  return null;
}
function learnRule(key, target) {
  if (!key || key.length < 3 || !target) return;
  const r = state.rules.find(x => x.key === key);
  if (r) r.target = target;
  else state.rules.push({ key, target, learned: true });
}

// ============================================================
// Stato dell'interfaccia
// ============================================================
const ui = {
  tab: 'aggiungi',
  month: thisMonth(),
  year: +thisMonth().slice(0, 4),
  listFilter: 'all',
  search: '',
  sumTab: 'mese',
  showZero: false,
  expanded: new Set(),
  openCats: new Set(),
  trend: { target: 'exp:all', mode: 'mesi', range: '12' },
};
const newForm = (kind = 'exp') => ({ kind, amount: '', cat: null, sub: null, inc: null, method: 'contanti', date: today(), note: '' });
let addForm = newForm();
let imp = null;          // sessione di importazione in corso
let importError = '';
let installPrompt = null;
let backupMode = 'merge';
const main = $('#main');

// ============================================================
// Componenti comuni
// ============================================================
let toastTimer;
function toast(msg, actionLabel, onAction) {
  const t = $('#toast');
  t.innerHTML = `<span>${esc(msg)}</span>${actionLabel ? `<button type="button">${esc(actionLabel)}</button>` : ''}`;
  t.hidden = false;
  if (actionLabel) t.querySelector('button').onclick = () => { t.hidden = true; onAction(); };
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { t.hidden = true; }, actionLabel ? 6000 : 3200);
}

const sheet = $('#sheet'), sheetBody = $('#sheetBody');
function openSheet() {
  sheet.hidden = false;
  document.body.style.overflow = 'hidden';
  history.pushState({ sheet: true }, '');
}
function closeSheet(fromHistory) {
  if (sheet.hidden) return;
  sheet.hidden = true;
  document.body.style.overflow = '';
  sheetBody.innerHTML = '';
  sheet.onclick = sheet.oninput = sheet.onchange = null;
  if (!fromHistory && history.state && history.state.sheet) history.back();
}
window.addEventListener('popstate', () => closeSheet(true));

const monthNavHTML = label => `<div class="month-nav">
  <button class="icon-btn" data-nav="-1" aria-label="Precedente">${ICON.left}</button>
  <strong>${esc(label)}</strong>
  <button class="icon-btn" data-nav="1" aria-label="Successivo">${ICON.right}</button></div>`;

const tile = (label, value, sub = '', cls = '') => `<div class="tile"><div class="t-label">${esc(label)}</div><div class="t-value num ${cls}">${value}</div>${sub ? `<div class="t-sub">${sub}</div>` : ''}</div>`;

function entryRowHTML(e, { showDate = false } = {}) {
  const uncl = isUnclassified(e);
  const color = e.type === 'inc' ? 'var(--good)' : uncl ? 'var(--s0)' : catColor(catById(e.cat));
  const title = e.note || e.desc || (e.type === 'inc' ? incById(e.inc)?.name : subOf(e.cat, e.sub)?.name) || 'Movimento';
  const parts = [uncl ? '<span class="badge warn">Da classificare</span>' : esc(entryLabel(e))];
  if (e.type === 'exp') parts.push(e.method === 'contanti' ? 'Contanti' : 'Conto');
  if (showDate) parts.push(fmtDate(e.date));
  const amt = e.type === 'inc' ? `<span class="pos">${fmtSigned(e.amount)}</span>`
    : e.amount < 0 ? `<span class="pos">+${fmt(-e.amount)}</span>` : `−${fmt(e.amount)}`;
  return `<button type="button" class="row" data-entry="${e.id}">
    <i class="dot" style="background:${color}"></i>
    <div class="r-main"><div class="r-title">${esc(title)}</div><div class="r-sub">${parts.join(' · ')}</div></div>
    <div class="r-amt num">${amt}</div></button>`;
}

function totals(list) {
  const t = { inc: 0, exp: 0, cash: 0, bank: 0, uncl: 0, unclCount: 0 };
  for (const e of list) {
    if (e.type === 'inc') t.inc += e.amount;
    else {
      t.exp += e.amount;
      if (e.method === 'contanti') t.cash += e.amount; else t.bank += e.amount;
    }
    if (isUnclassified(e)) { t.uncl += e.type === 'exp' ? e.amount : 0; t.unclCount++; }
  }
  t.bal = t.inc - t.exp;
  return t;
}

// ============================================================
// Modulo movimento (usato per aggiungere e per modificare)
// ============================================================
function recentSubs() {
  const counts = new Map();
  const manual = state.entries.filter(e => e.type === 'exp' && e.source === 'manual' && subOf(e.cat, e.sub))
    .sort((a, b) => b.created - a.created).slice(0, 60);
  manual.forEach((e, i) => {
    const k = `${e.cat}:${e.sub}`;
    counts.set(k, (counts.get(k) || 0) + 1 + (60 - i) / 1000);
  });
  return [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 6).map(([k]) => k);
}

function formHTML(f, mode) {
  const recent = mode === 'add' && f.kind === 'exp' ? recentSubs() : [];
  const cat = catById(f.cat);
  const on = c => (c ? ' on' : '');
  let h = `<div class="seg" role="group" aria-label="Tipo di movimento">
      <button type="button" data-kind="exp" class="${f.kind === 'exp' ? 'on' : ''}">Spesa</button>
      <button type="button" data-kind="inc" class="${f.kind === 'inc' ? 'on' : ''}">Entrata</button></div>
    <label class="amount-field"><input id="fAmount" inputmode="decimal" autocomplete="off" placeholder="0,00" value="${esc(f.amount)}" aria-label="Importo in euro"><span>€</span></label>`;
  if (f.kind === 'exp') {
    if (recent.length) {
      h += `<div class="label">Usate di recente</div><div class="chips">${recent.map(k => {
        const [c, s] = k.split(':');
        return `<button type="button" class="chip${on(f.cat === c && f.sub === s)}" data-pick="${esc(k)}"><i class="dot" style="background:${catColor(catById(c))}"></i>${esc(subOf(c, s).name)}</button>`;
      }).join('')}</div>`;
    }
    h += `<div class="label">Categoria</div><div class="chips">${state.categories.map(c =>
      `<button type="button" class="chip${on(f.cat === c.id)}" data-cat="${esc(c.id)}"><i class="dot" style="background:${catColor(c)}"></i>${esc(c.name)}</button>`).join('')}</div>`;
    if (cat) {
      h += `<div class="label">Sottocategoria di ${esc(cat.name)}</div><div class="chips">${cat.subs.map(s =>
        `<button type="button" class="chip${on(f.sub === s.id)}" data-sub="${esc(s.id)}">${esc(s.name)}</button>`).join('') || '<span class="muted small">Nessuna sottocategoria: aggiungine una in Altro → Categorie.</span>'}</div>`;
    }
    h += `<div class="label">Pagata con</div><div class="seg" role="group" aria-label="Metodo di pagamento">
      <button type="button" data-method="contanti" class="${f.method === 'contanti' ? 'on' : ''}">Contanti</button>
      <button type="button" data-method="banca" class="${f.method === 'banca' ? 'on' : ''}">Conto / carta</button></div>`;
  } else {
    h += `<div class="label">Tipo di entrata</div><div class="chips">${state.incomeTypes.map(t =>
      `<button type="button" class="chip${on(f.inc === t.id)}" data-inc="${esc(t.id)}">${esc(t.name)}</button>`).join('')}</div>`;
  }
  h += `<div class="label">Data</div><div class="date-row">
      <input type="date" id="fDate" value="${esc(f.date)}" aria-label="Data">
      <button type="button" class="btn small" data-date="today">Oggi</button>
      <button type="button" class="btn small" data-date="yesterday">Ieri</button></div>
    <label class="field"><span>Nota (facoltativa)</span><input id="fNote" value="${esc(f.note)}" maxlength="120"
      placeholder="${f.kind === 'exp' ? 'es. mercato, pizza con amici…' : 'es. tredicesima, rimborso 730…'}"></label>
    <div class="form-error" id="fError" hidden></div>`;
  return h;
}

function formClick(ev, f, rerender) {
  const b = ev.target.closest('button');
  if (!b) return false;
  const d = b.dataset;
  if (d.kind) { if (f.kind !== d.kind) { f.kind = d.kind; rerender(); } return true; }
  if (d.pick) { [f.cat, f.sub] = d.pick.split(':'); rerender(); return true; }
  if (d.cat) {
    if (f.cat !== d.cat) { f.cat = d.cat; const c = catById(d.cat); f.sub = c && c.subs.length === 1 ? c.subs[0].id : null; }
    rerender(); return true;
  }
  if (d.sub) { f.sub = d.sub; rerender(); return true; }
  if (d.inc) { f.inc = d.inc; rerender(); return true; }
  if (d.method) { f.method = d.method; rerender(); return true; }
  if (d.date) { f.date = d.date === 'today' ? today() : yesterday(); $('#fDate').value = f.date; return true; }
  return false;
}
function formInput(ev, f) {
  if (ev.target.id === 'fAmount') f.amount = ev.target.value;
  else if (ev.target.id === 'fNote') f.note = ev.target.value;
  else if (ev.target.id === 'fDate') f.date = ev.target.value;
}
function readForm(f, { requireCategory, allowNegative }) {
  const n = parseNumber(f.amount);
  if (!Number.isFinite(n) || n === 0 || (!allowNegative && n < 0)) return { error: 'Inserisci un importo valido (es. 12,50).' };
  if (requireCategory && f.kind === 'exp' && !(f.cat && subOf(f.cat, f.sub))) return { error: 'Scegli categoria e sottocategoria.' };
  if (requireCategory && f.kind === 'inc' && !incById(f.inc)) return { error: 'Scegli il tipo di entrata.' };
  if (!parseDate(f.date)) return { error: 'Inserisci una data valida.' };
  return { cents: toCents(n) };
}
function showFormError(root, msg) {
  const el = $('#fError', root);
  el.textContent = msg;
  el.hidden = false;
}

// ============================================================
// Vista: Aggiungi
// ============================================================
function installBannerHTML() {
  if (isStandalone() || !isMobile || state.settings.installDismissed) return '';
  return `<div class="notice info">${ICON.info}<div style="flex:1">Installa l'app sul telefono per aprirla dalla schermata Home, anche senza connessione.
    <div style="margin-top:6px"><button type="button" class="linkish" data-act="install">Come si fa</button></div></div>
    <button type="button" class="linkish" data-act="dismiss-install" aria-label="Chiudi">${ICON.close}</button></div>`;
}

function viewAdd() {
  const f = addForm;
  const recent = [...state.entries].filter(e => e.source === 'manual').sort((a, b) => b.created - a.created).slice(0, 5);
  main.innerHTML = `${installBannerHTML()}
    <section class="card"><div id="formBox">${formHTML(f, 'add')}</div>
      <button type="button" class="btn primary big block" id="saveBtn" style="margin-top:18px">${f.kind === 'exp' ? 'Salva spesa' : 'Salva entrata'}</button>
    </section>
    <h2 class="section-title">Ultime registrazioni a mano</h2>
    ${recent.length ? `<div class="rows">${recent.map(e => entryRowHTML(e, { showDate: true })).join('')}</div>`
      : '<p class="hint">Le spese in contanti e le entrate che registri compariranno qui.</p>'}`;
  const rerender = () => {
    $('#formBox').innerHTML = formHTML(f, 'add');
    $('#saveBtn').textContent = f.kind === 'exp' ? 'Salva spesa' : 'Salva entrata';
  };
  main.onclick = ev => {
    if (formClick(ev, f, rerender)) return;
    const act = ev.target.closest('[data-act]')?.dataset.act;
    if (ev.target.closest('#saveBtn')) saveAdd();
    else if (act === 'install') startInstall();
    else if (act === 'dismiss-install') { state.settings.installDismissed = true; save(); render(); }
    else { const row = ev.target.closest('[data-entry]'); if (row) openEdit(row.dataset.entry); }
  };
  main.oninput = main.onchange = ev => formInput(ev, f);
}

function saveAdd() {
  const f = addForm;
  const r = readForm(f, { requireCategory: true, allowNegative: false });
  if (r.error) { showFormError(main, r.error); return; }
  const e = {
    id: uid(), type: f.kind, date: parseDate(f.date), amount: r.cents,
    cat: f.kind === 'exp' ? f.cat : null, sub: f.kind === 'exp' ? f.sub : null, inc: f.kind === 'inc' ? f.inc : null,
    method: f.kind === 'exp' ? f.method : 'banca', source: 'manual', note: f.note.trim(), created: Date.now(),
  };
  state.entries.push(e);
  save();
  requestPersist();
  addForm = { ...newForm(f.kind), method: f.method, date: f.date };
  render();
  window.scrollTo(0, 0);
  toast(`${f.kind === 'exp' ? 'Spesa' : 'Entrata'} salvata: ${fmt(r.cents)} · ${entryLabel(e)}`, 'Annulla', () => {
    state.entries = state.entries.filter(x => x.id !== e.id);
    save();
    render();
  });
}

// ============================================================
// Modifica di un movimento (pannello dal basso)
// ============================================================
function openEdit(id) {
  const e = state.entries.find(x => x.id === id);
  if (!e) return;
  const f = { kind: e.type, amount: (e.amount / 100).toFixed(2).replace('.', ','), cat: e.cat, sub: e.sub, inc: e.inc, method: e.method || 'banca', date: e.date, note: e.note || '' };
  const before = entryTarget(e);
  const similar = e.mKey ? state.entries.filter(x => x.id !== e.id && x.mKey === e.mKey) : [];
  sheetBody.innerHTML = `<div class="sheet-head"><h2 id="sheetTitle">Modifica movimento</h2>
      <button type="button" class="icon-btn" data-close aria-label="Chiudi">${ICON.close}</button></div>
    ${e.source === 'import' ? `<div class="bank-desc"><b>Dalla banca:</b> ${esc(e.bankDesc || e.desc)}</div>` : ''}
    <div id="formBox">${formHTML(f, 'edit')}</div>
    ${e.mKey ? `<label class="check"><input type="checkbox" id="fLearn" checked><span>Se cambi categoria, ricordala per i prossimi import da «${esc(e.desc)}»${similar.length === 1 ? ' e applicala anche all\'altro movimento simile già importato' : similar.length ? ` e applicala anche agli altri ${similar.length} movimenti simili già importati` : ''}.</span></label>` : ''}
    <div class="btn-row" style="margin-top:18px">
      <button type="button" class="btn danger" data-act="delete">Elimina</button>
      <button type="button" class="btn primary" data-act="save">Salva</button></div>`;
  openSheet();
  const rerender = () => { $('#formBox', sheetBody).innerHTML = formHTML(f, 'edit'); };
  sheet.oninput = sheet.onchange = ev => formInput(ev, f);
  sheet.onclick = ev => {
    if (ev.target.closest('[data-close]')) { closeSheet(); return; }
    if (formClick(ev, f, rerender)) return;
    const act = ev.target.closest('[data-act]')?.dataset.act;
    if (act === 'delete') {
      if (!confirm('Eliminare questo movimento?')) return;
      state.entries = state.entries.filter(x => x.id !== e.id);
      if (e.bankKey && !state.ignoredKeys.includes(e.bankKey)) state.ignoredKeys.push(e.bankKey);
      save();
      closeSheet();
      render();
      toast('Movimento eliminato', 'Annulla', () => {
        state.entries.push(e);
        state.ignoredKeys = state.ignoredKeys.filter(k => k !== e.bankKey);
        save();
        render();
      });
    } else if (act === 'save') {
      const r = readForm(f, { requireCategory: false, allowNegative: true });
      if (r.error) { showFormError(sheetBody, r.error); return; }
      Object.assign(e, { type: f.kind, amount: r.cents, date: parseDate(f.date), note: f.note.trim() });
      if (f.kind === 'exp') Object.assign(e, { cat: subOf(f.cat, f.sub) ? f.cat : null, sub: subOf(f.cat, f.sub) ? f.sub : null, inc: null, method: f.method });
      else Object.assign(e, { inc: incById(f.inc) ? f.inc : null, cat: null, sub: null });
      const after = entryTarget(e);
      let changed = 0;
      if (e.mKey && after && after !== before && $('#fLearn', sheetBody)?.checked) {
        learnRule(e.mKey, after);
        for (const x of similar) {
          const bank = bankCentsOf(x);
          if (after.startsWith('inc:') && bank < 0) continue;
          if (entryTarget(x) !== after) { applyTarget(x, after, bank); changed++; }
        }
      }
      save();
      closeSheet();
      render();
      toast(changed === 1 ? 'Salvato. Aggiornato anche un movimento simile' : changed ? `Salvato. Aggiornati anche ${changed} movimenti simili` : 'Movimento aggiornato');
    }
  };
}

// ============================================================
// Vista: Movimenti
// ============================================================
function backupNoticeHTML() {
  if (state.entries.length < 15) return '';
  const last = state.settings.lastBackup;
  const days = last ? (Date.now() - new Date(last).getTime()) / 86400000 : Infinity;
  if (days < 30) return '';
  return `<div class="notice">${ICON.warn}<div>${last ? `L'ultimo backup è del ${fmtDate(last)}.` : 'Non hai ancora fatto un backup.'}
    I dati sono salvati solo su questo dispositivo: <a href="#altro">fai un backup</a>.</div></div>`;
}

function viewList() {
  const q = ui.search.trim().toLowerCase();
  const monthList = entriesIn(ui.month);
  const t = totals(monthList);
  let list = q
    ? state.entries.filter(e => [e.note, e.desc, e.bankDesc, entryLabel(e), fmt(e.amount)].join(' ').toLowerCase().includes(q))
    : monthList;
  const unclCount = list.filter(isUnclassified).length;
  const filters = [['all', 'Tutti'], ['exp', 'Spese'], ['inc', 'Entrate'], ['cash', 'Contanti'], ['uncl', `Da classificare${unclCount ? ` (${unclCount})` : ''}`]];
  const F = ui.listFilter;
  list = list.filter(e => F === 'all' || (F === 'exp' && e.type === 'exp') || (F === 'inc' && e.type === 'inc')
    || (F === 'cash' && e.type === 'exp' && e.method === 'contanti') || (F === 'uncl' && isUnclassified(e)));
  list.sort((a, b) => b.date.localeCompare(a.date) || (b.created || 0) - (a.created || 0));

  const groups = [];
  for (const e of list) {
    if (!groups.length || groups[groups.length - 1].date !== e.date) groups.push({ date: e.date, items: [] });
    groups[groups.length - 1].items.push(e);
  }
  main.innerHTML = `${backupNoticeHTML()}
    ${q ? `<p class="hint" style="margin:0 0 10px">Risultati della ricerca in tutti i mesi: ${list.length}</p>` : monthNavHTML(monthName(ui.month))}
    ${q ? '' : `<div class="tiles">${tile('Entrate', fmtTile(t.inc), '', 'pos')}${tile('Spese', fmtTile(t.exp))}${tile('Bilancio', fmtTile(t.bal, true), '', t.bal < 0 ? 'neg' : 'pos')}</div>`}
    <input type="search" id="search" placeholder="Cerca (es. Esselunga, pizza…)" value="${esc(ui.search)}" aria-label="Cerca nei movimenti" style="margin-bottom:10px">
    <div class="chips" style="margin-bottom:4px">${filters.map(([k, l]) => `<button type="button" class="chip${F === k ? ' on' : ''}" data-filter="${k}">${esc(l)}</button>`).join('')}</div>
    ${groups.length ? groups.map(g => {
      const dayExp = sum(g.items.filter(e => e.type === 'exp').map(e => e.amount));
      return `<div class="day-head"><span>${esc(dayLabel(g.date))}</span>${dayExp ? `<span class="num">−${fmt(dayExp)}</span>` : ''}</div>
        <div class="rows">${g.items.map(e => entryRowHTML(e, { showDate: !!q })).join('')}</div>`;
    }).join('') : `<div class="empty"><strong>Nessun movimento</strong>${q ? 'Prova con un\'altra parola.' : F !== 'all' ? 'Nessun movimento con questo filtro.' : 'Aggiungi una spesa o importa l\'estratto conto.'}</div>`}`;

  main.onclick = ev => {
    const nav = ev.target.closest('[data-nav]');
    if (nav) { ui.month = addMonths(ui.month, +nav.dataset.nav); render(); return; }
    const fl = ev.target.closest('[data-filter]');
    if (fl) { ui.listFilter = fl.dataset.filter; render(); return; }
    const row = ev.target.closest('[data-entry]');
    if (row) openEdit(row.dataset.entry);
  };
  main.oninput = debounce(ev => {
    if (ev.target.id !== 'search') return;
    ui.search = ev.target.value;
    render();
    const s = $('#search');
    s.focus();
    s.setSelectionRange(s.value.length, s.value.length);
  }, 250);
  main.onchange = null;
}

// ============================================================
// Vista: Importa dalla banca
// ============================================================
const HEADER_PATTERNS = {
  date: /^data( |_)?(operazione|contabile|contabilizzazione|registrazione|movimento)?$|^date$/i,
  valuta: /valuta/i,
  amount: /^(importo|amount|ammontare)/i,
  inc: /^(entrat|accredit|avere)/i,
  out: /^(uscit|addebit|dare)/i,
  desc: /^(descrizione|causale|description|operazione|dettagli|descrizione operazione)$/i,
  full: /(descrizione.?complet|descrizione.?estesa|causale.?complet|dettagli)/i,
  cat: /(moneymap|categori)/i,
  state: /^(stato|status)$/i,
};
function detectMapping(rows) {
  let best = -1, bestScore = 1;
  for (let i = 0; i < Math.min(rows.length, 40); i++) {
    const score = rows[i].filter(c => String(c).trim() && Object.values(HEADER_PATTERNS).some(re => re.test(String(c).trim()))).length;
    if (score > bestScore) { bestScore = score; best = i; }
  }
  if (best < 0) return null;
  const headers = rows[best].map(c => String(c ?? '').trim());
  const find = (re, skip = -2) => headers.findIndex((h, i) => h && i !== skip && re.test(h));
  const m = { header: best };
  const opDate = find(HEADER_PATTERNS.date), valuta = find(HEADER_PATTERNS.valuta);
  const fineco = headers.some(h => /moneymap|descrizione_completa/i.test(h));
  // Per Fineco la data valuta è quella reale del pagamento con carta
  m.date = fineco && valuta > -1 ? valuta : opDate > -1 ? opDate : valuta;
  m.full = find(HEADER_PATTERNS.full);
  m.desc = find(HEADER_PATTERNS.desc, m.full);
  if (m.desc < 0 && m.full > -1) { m.desc = m.full; m.full = -1; }
  m.inc = find(HEADER_PATTERNS.inc);
  m.out = find(HEADER_PATTERNS.out);
  m.amount = m.inc > -1 && m.out > -1 ? -1 : find(HEADER_PATTERNS.amount);
  m.cat = find(HEADER_PATTERNS.cat);
  m.state = find(HEADER_PATTERNS.state);
  return { mapping: m, headers };
}
const mappingValid = m => m.date > -1 && (m.amount > -1 || m.inc > -1 || m.out > -1) && (m.desc > -1 || m.full > -1);

async function loadImportFile(file) {
  importError = '';
  main.innerHTML = '<div class="empty"><strong>Lettura del file…</strong></div>';
  try {
    const rows = await XlsxLite.read(file);
    const det = detectMapping(rows);
    if (!det) throw new Error('Non trovo la riga con i nomi delle colonne (Data, Descrizione, Importo…). Controlla che sia l\'elenco movimenti esportato dalla banca.');
    const sig = det.headers.filter(Boolean).join('|');
    const saved = state.settings.mappings[sig];
    imp = { fileName: file.name, rows, headers: det.headers, sig, mapping: saved ? { ...det.mapping, ...saved, header: det.mapping.header } : det.mapping, filter: 'new', showMap: false };
    if (!mappingValid(imp.mapping)) imp.showMap = true;
    buildImportItems();
  } catch (err) {
    console.error(err);
    imp = null;
    importError = err.message || 'Impossibile leggere il file.';
  }
  render();
}

function buildImportItems() {
  const { rows, mapping: m } = imp;
  imp.items = [];
  if (!mappingValid(m)) return;
  const seen = new Map();
  const existing = new Set(state.entries.map(e => e.bankKey).filter(Boolean));
  const ignored = new Set(state.ignoredKeys);
  const cell = (r, i) => (i > -1 ? r[i] : '');
  for (let i = m.header + 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r.length) continue;
    const date = parseDate(cell(r, m.date));
    if (!date) continue;
    let amount;
    if (m.amount > -1) amount = parseNumber(cell(r, m.amount));
    else {
      const a = parseNumber(cell(r, m.inc)), b = parseNumber(cell(r, m.out));
      amount = (Number.isFinite(a) ? Math.abs(a) : 0) - (Number.isFinite(b) ? Math.abs(b) : 0);
    }
    if (!Number.isFinite(amount) || amount === 0) continue;
    const cents = toCents(amount);
    const short = String(cell(r, m.desc) ?? '');
    const full = m.full > -1 && m.full !== m.desc ? String(cell(r, m.full) ?? '') : '';
    const info = analyzeDescription(short, full);
    const base = `${date}|${cents}|${compact(full || short).slice(0, 48)}`;
    const n = (seen.get(base) || 0) + 1;
    seen.set(base, n);
    const bankKey = `${base}#${n}`;
    const pending = /non\s*contab|in\s*attesa|pending|da\s*contab|prenotat/i.test(String(cell(r, m.state) ?? ''));
    const bankCat = String(cell(r, m.cat) ?? '').trim();
    const sug = matchRules(info, cents) || suggestFromBank(bankCat, cents);
    imp.items.push({
      row: i, date, cents, short, full, info, bankCat, bankKey, pending,
      dup: existing.has(bankKey) ? 'imported' : ignored.has(bankKey) ? 'ignored' : '',
      target: sug ? sug.target : '', via: sug ? sug.via : '',
    });
  }
  imp.items.sort((a, b) => b.date.localeCompare(a.date) || a.row - b.row);
}

function mappingEditorHTML() {
  const m = imp.mapping;
  const opts = sel => `<option value="-1">—</option>` + imp.headers.map((h, i) => (h ? `<option value="${i}"${i === sel ? ' selected' : ''}>${esc(h)}</option>` : '')).join('');
  const field = (key, label) => `<label class="field"><span>${label}</span><select data-map="${key}">${opts(m[key])}</select></label>`;
  return `<div class="map-grid">
    ${field('date', 'Data del movimento')}
    ${field('amount', 'Importo (colonna unica con il segno)')}
    ${field('inc', 'Oppure: colonna Entrate')}
    ${field('out', 'Oppure: colonna Uscite')}
    ${field('desc', 'Descrizione')}
    ${field('full', 'Descrizione completa (facoltativa)')}
    ${field('cat', 'Categoria della banca (facoltativa)')}
    ${field('state', 'Stato (facoltativo)')}</div>`;
}

function importItemHTML(it, idx) {
  const off = it.dup || it.pending;
  const badge = it.dup ? `<span class="badge">${it.dup === 'imported' ? 'Già importato' : 'Già escluso in passato'}</span>`
    : it.pending ? '<span class="badge">Non ancora contabilizzato</span>'
    : it.target === 'ign' ? '<span class="badge">Non verrà importato</span>'
    : it.via === 'regola' ? '<span class="badge">Regola tua</span>'
    : it.via === 'auto' ? '<span class="badge">Automatica</span>'
    : it.via === 'banca' ? '<span class="badge">Suggerita dalla banca</span>'
    : !it.target ? '<span class="badge warn">Da classificare</span>' : '';
  const extra = [fmtDate(it.date), it.info.extra, it.bankCat].filter(Boolean).map(esc).join(' · ');
  return `<div class="imp-row${off ? ' off' : ''}">
    <div class="imp-top"><div class="r-main"><div class="imp-title">${esc(it.info.title)}</div><div class="imp-extra">${extra}</div></div>
      <div class="r-amt num${it.cents > 0 ? ' pos' : ''}">${it.cents > 0 ? '+' : '−'}${fmt(Math.abs(it.cents))}</div></div>
    ${off ? '' : `<select data-item="${idx}" class="${it.target ? '' : 'unset'}" aria-label="Categoria per ${esc(it.info.title)}">${targetOptions(it.target, { income: it.cents > 0, ignore: true, refund: it.cents > 0 })}</select>`}
    ${badge ? `<div style="margin-top:6px">${badge}</div>` : ''}</div>`;
}

function viewImport() {
  if (!imp) {
    main.innerHTML = `${importError ? `<div class="notice">${ICON.warn}<div>${esc(importError)}</div></div>` : ''}
      <section class="card">
        <label class="drop" for="fileIn">${ICON.upload}
          <strong>Scegli il file dei movimenti</strong>
          <span class="muted small">Excel (.xlsx) o CSV scaricato dall'home banking</span>
          <span class="btn primary" style="margin-top:6px">Scegli file</span></label>
        <input type="file" id="fileIn" accept=".xlsx,.xlsm,.xls,.csv,.txt,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv" hidden>
      </section>
      <section class="card"><h2>Come funziona</h2>
        <ol class="steps">
          <li>Dal sito della banca apri l'elenco movimenti del conto, scegli il periodo (es. il mese appena finito) e scarica il file Excel.</li>
          <li>Caricalo qui: l'app riconosce le colonne e propone una categoria per ogni spesa.</li>
          <li>Controlla le proposte e correggi quelle sbagliate: l'app impara e dalla volta successiva le classifica da sola.</li>
          <li>I movimenti già importati vengono riconosciuti e saltati, quindi puoi importare anche periodi che si sovrappongono.</li>
        </ol>
        <p class="hint">I prelievi di contanti vengono ignorati: le spese in contanti le registri a mano, così non vengono contate due volte.</p>
      </section>`;
    main.onclick = main.oninput = null;
    main.onchange = ev => { if (ev.target.id === 'fileIn' && ev.target.files[0]) loadImportFile(ev.target.files[0]); };
    return;
  }

  const items = imp.items;
  const fresh = items.filter(it => !it.dup && !it.pending);
  const todo = fresh.filter(it => !it.target).length;
  const old = items.filter(it => it.dup).length;
  const pend = items.filter(it => it.pending).length;
  const toImport = fresh.filter(it => it.target !== 'ign');
  const dates = items.map(it => it.date).sort();
  const m = imp.mapping, H = i => esc(imp.headers[i] || '—');
  const amountLabel = m.amount > -1 ? H(m.amount) : [m.inc, m.out].filter(i => i > -1).map(H).join(' / ');
  const list = items.map((it, i) => [it, i]).filter(([it]) =>
    imp.filter === 'all' || (imp.filter === 'new' && !it.dup && !it.pending) || (imp.filter === 'todo' && !it.dup && !it.pending && !it.target) || (imp.filter === 'old' && (it.dup || it.pending)));
  const chips = [['new', `Nuovi (${fresh.length})`], ['todo', `Da classificare (${todo})`]];
  if (old + pend) chips.push(['old', `Esclusi (${old + pend})`]);

  main.innerHTML = `<section class="card">
      <div style="display:flex;justify-content:space-between;gap:10px;align-items:flex-start">
        <div style="min-width:0"><b style="overflow-wrap:anywhere">${esc(imp.fileName)}</b>
          <div class="muted small">${items.length ? `${items.length} movimenti dal ${fmtDate(dates[0])} al ${fmtDate(dates[dates.length - 1])}` : 'Nessun movimento riconosciuto'}</div></div>
        <button type="button" class="btn small" data-act="cancel">Annulla</button></div>
      <p class="small muted" style="margin:10px 0 0">Colonne usate: data = <b>${H(m.date)}</b> · importo = <b>${amountLabel || '—'}</b> · descrizione = <b>${H(m.full > -1 ? m.full : m.desc)}</b>${m.cat > -1 ? ` · categoria banca = <b>${H(m.cat)}</b>` : ''}
        · <button type="button" class="linkish" data-act="map">${imp.showMap ? 'Chiudi' : 'Modifica'}</button></p>
      ${imp.showMap ? mappingEditorHTML() : ''}
    </section>
    ${old ? `<div class="notice info">${ICON.info}<div>${old} movimenti erano già stati importati (o esclusi) e verranno saltati.</div></div>` : ''}
    ${pend ? `<div class="notice info">${ICON.info}<div>${pend} movimenti non sono ancora contabilizzati: li importerai la prossima volta, quando saranno definitivi.</div></div>` : ''}
    ${items.length ? `<div class="chips" style="margin-bottom:10px">${chips.map(([k, l]) => `<button type="button" class="chip${imp.filter === k ? ' on' : ''}" data-ifilter="${k}">${esc(l)}</button>`).join('')}</div>
      ${list.length ? `<div class="rows">${list.map(([it, i]) => importItemHTML(it, i)).join('')}</div>` : '<div class="empty">Nessun movimento in questo elenco.</div>'}
      <div class="sticky-bar">
        <button type="button" class="btn primary block big" data-act="commit"${toImport.length ? '' : ' disabled'}>${toImport.length ? `Importa ${toImport.length} movimenti` : 'Niente di nuovo da importare'}</button>
        ${todo ? `<div class="small muted" style="text-align:center;margin-top:6px">${todo} da classificare: puoi farlo anche dopo, da Movimenti.</div>` : ''}
      </div>` : ''}`;

  main.onclick = ev => {
    const act = ev.target.closest('[data-act]')?.dataset.act;
    const chip = ev.target.closest('[data-ifilter]');
    if (chip) { imp.filter = chip.dataset.ifilter; render(); }
    else if (act === 'cancel') { imp = null; render(); }
    else if (act === 'map') { imp.showMap = !imp.showMap; render(); }
    else if (act === 'commit') commitImport();
  };
  main.oninput = null;
  main.onchange = ev => {
    const mapSel = ev.target.closest('[data-map]');
    if (mapSel) {
      imp.mapping[mapSel.dataset.map] = +mapSel.value;
      const { header, ...rest } = imp.mapping;
      state.settings.mappings[imp.sig] = rest;
      save();
      buildImportItems();
      render();
      return;
    }
    const sel = ev.target.closest('select[data-item]');
    if (!sel) return;
    const it = imp.items[+sel.dataset.item];
    it.target = sel.value;
    it.via = 'manual';
    let others = 0;
    if (it.target) {
      learnRule(it.info.key, it.target);
      save();
      for (const o of imp.items) {
        if (o === it || o.dup || o.pending || o.via === 'manual' || o.info.key !== it.info.key) continue;
        if (it.target.startsWith('inc:') && o.cents < 0) continue;
        o.target = it.target;
        o.via = 'regola';
        others++;
      }
    }
    render();
    if (others) toast(others === 1 ? 'Applicato anche a un altro movimento simile' : `Applicato anche ad altri ${others} movimenti simili`);
  };
}

function commitImport() {
  let added = 0, uncl = 0, skipped = 0;
  const months = [];
  for (const it of imp.items) {
    if (it.dup || it.pending) continue;
    if (it.target === 'ign') { state.ignoredKeys.push(it.bankKey); skipped++; continue; }
    const e = {
      id: uid(), date: it.date, method: 'banca', source: 'import', note: '', created: Date.now(),
      bankKey: it.bankKey, bankDesc: (it.full || it.short).replace(/\s+/g, ' ').trim(), desc: it.info.title, mKey: it.info.key,
    };
    applyTarget(e, it.target, it.cents);
    if (isUnclassified(e)) uncl++;
    state.entries.push(e);
    months.push(it.date.slice(0, 7));
    added++;
  }
  save();
  requestPersist();
  imp = null;
  if (months.length) ui.month = months.sort()[months.length - 1];
  ui.listFilter = uncl ? 'uncl' : 'all';
  ui.search = '';
  location.hash = '#movimenti';
  toast(`Importati ${added} movimenti${uncl ? `, ${uncl} da classificare` : ''}${skipped ? ` (${skipped} esclusi)` : ''}`);
}

// ============================================================
// Grafici (SVG disegnati a mano)
// ============================================================
let chartSpecs = new Map(), chartSeq = 0;
function chartSlot(spec) {
  const id = 'c' + ++chartSeq;
  chartSpecs.set(id, spec);
  return `<div class="chart" data-chart="${id}"></div>`;
}
function mountCharts() {
  $$('.chart[data-chart]').forEach(el => {
    const spec = chartSpecs.get(el.dataset.chart);
    if (spec) { el._spec = spec; drawChart(el); }
  });
  chartSpecs = new Map();
}
const drawChart = el => (el._spec.kind === 'lines' ? drawLines : drawColumns)(el, el._spec);
window.addEventListener('resize', debounce(() => $$('.chart').forEach(el => el._spec && drawChart(el)), 150));
document.addEventListener('pointerdown', ev => $$('.chart').forEach(c => { if (!c.contains(ev.target) && c._hideTip) c._hideTip(); }));

function niceTicks(max) {
  const raw = max / 4;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const step = [1, 2, 2.5, 5, 10].map(f => f * mag).find(s => s >= raw);
  const ticks = [];
  for (let v = 0; v < max + step * 0.999; v += step) ticks.push(Math.round(v));
  return ticks;
}
const roundTop = (x, y, w, h, r) => `M${x},${y + h}V${y + r}Q${x},${y} ${x + r},${y}H${x + w - r}Q${x + w},${y} ${x + w},${y + r}V${y + h}Z`;
const tipRow = (name, v, color, line) => `<div class="tip-r"><span><i class="${line ? 'key-line' : 'dot'}" style="background:${color}"></i>${esc(name)}</span><span class="num">${fmt(v)}</span></div>`;

function chartFrame(el, spec, maxValue) {
  const W = Math.max(260, Math.floor(el.clientWidth)), H = spec.height || 220;
  const M = { l: 40, r: 10, t: 14, b: 26 };
  const n = spec.labels.length, pw = W - M.l - M.r, ph = H - M.t - M.b;
  const ticks = niceTicks(Math.max(100, maxValue));
  const top = ticks[ticks.length - 1];
  const band = pw / n;
  const g = { W, H, M, n, pw, ph, band, Y: v => M.t + ph - (Math.max(0, v) / top) * ph, X: i => M.l + band * i + band / 2 };
  let s = `<svg width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(spec.aria || '')}">`;
  s += `<rect class="hl" x="0" y="${M.t}" width="${band}" height="${ph}" fill="var(--hover)" visibility="hidden"/>`;
  for (const t of ticks) {
    const y = Math.round(g.Y(t)) + 0.5;
    s += `<line x1="${M.l}" x2="${W - M.r}" y1="${y}" y2="${y}" stroke="${t === 0 ? 'var(--axis)' : 'var(--grid)'}" stroke-width="1"/>`;
    s += `<text x="${M.l - 6}" y="${y + 4}" text-anchor="end" font-size="11" fill="var(--muted)">${fmtShort(t)}</text>`;
  }
  const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(pw / 30))));
  spec.labels.forEach((lb, i) => {
    if (i % every === 0) s += `<text x="${g.X(i)}" y="${H - 8}" text-anchor="middle" font-size="11" fill="var(--muted)">${esc(lb)}</text>`;
  });
  return { g, s };
}

function bindTip(el, g, tipFn, onIndex) {
  const svg = $('svg', el), tip = $('.tip', el), hl = $('.hl', svg);
  const hide = () => { tip.hidden = true; hl.setAttribute('visibility', 'hidden'); if (onIndex) onIndex(-1); };
  const show = ev => {
    const px = ev.clientX - svg.getBoundingClientRect().left;
    const i = Math.floor((px - g.M.l) / g.band);
    if (i < 0 || i >= g.n) { hide(); return; }
    tip.innerHTML = tipFn(i);
    tip.hidden = false;
    if (!onIndex) { hl.setAttribute('x', g.M.l + g.band * i); hl.setAttribute('width', g.band); hl.setAttribute('visibility', 'visible'); }
    else onIndex(i);
    const tw = tip.offsetWidth, cx = g.X(i);
    let left = cx + 14;
    if (left + tw > g.W) left = cx - 14 - tw;
    tip.style.left = Math.max(0, left) + 'px';
  };
  el.onpointermove = show;
  el.onpointerdown = show;
  el.onpointerleave = ev => { if (ev.pointerType === 'mouse') hide(); };
  el._hideTip = hide;
}

// Colonne (anche impilate) con linea facoltativa e linea della media
function drawColumns(el, spec) {
  const sums = spec.labels.map((_, i) => sum(spec.stacks.map(st => Math.max(0, st.values[i] || 0))));
  const lineVals = spec.line ? spec.line.values.filter(v => v != null) : [];
  const { g, s: frame } = chartFrame(el, spec, Math.max(0, ...sums, ...lineVals, spec.avg || 0));
  let s = frame;
  const bw = Math.max(3, Math.min(24, g.band * 0.62));
  for (let i = 0; i < g.n; i++) {
    let acc = 0;
    const segs = spec.stacks.filter(st => (st.values[i] || 0) > 0);
    segs.forEach((st, k) => {
      const v = st.values[i];
      const yTop = g.Y(acc + v);
      const yBot = g.Y(acc) - (k > 0 ? 2 : 0);
      acc += v;
      const h = yBot - yTop;
      if (h < 0.6) return;
      const x0 = g.X(i) - bw / 2;
      s += k === segs.length - 1
        ? `<path d="${roundTop(x0, yTop, bw, h, Math.min(4, h))}" fill="${st.color}"/>`
        : `<rect x="${x0}" y="${yTop}" width="${bw}" height="${h}" fill="${st.color}"/>`;
    });
  }
  if (spec.avg) {
    const y = Math.round(g.Y(spec.avg)) + 0.5;
    s += `<line x1="${g.M.l}" x2="${g.W - g.M.r}" y1="${y}" y2="${y}" stroke="var(--ink-2)" stroke-width="1"/>`;
    s += `<text x="${g.W - g.M.r}" y="${y - 5}" text-anchor="end" font-size="11" fill="var(--ink-2)" stroke="var(--card)" stroke-width="3" paint-order="stroke">media ${fmtShort(spec.avg)} €</text>`;
  }
  if (spec.line) {
    let d = '', pen = false;
    spec.line.values.forEach((v, i) => {
      if (v == null) { pen = false; return; }
      d += `${pen ? 'L' : 'M'}${g.X(i)},${g.Y(v)}`;
      pen = true;
    });
    s += `<path d="${d}" fill="none" stroke="${spec.line.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    spec.line.values.forEach((v, i) => { if (v != null) s += `<circle cx="${g.X(i)}" cy="${g.Y(v)}" r="4" fill="${spec.line.color}" stroke="var(--card)" stroke-width="2"/>`; });
  }
  el.innerHTML = s + '</svg><div class="tip" hidden></div>';
  bindTip(el, g, i => {
    const rows = spec.stacks.filter(st => (st.values[i] || 0) !== 0);
    let h = `<div class="tip-h">${esc(spec.fullLabels ? spec.fullLabels[i] : spec.labels[i])}</div>`;
    h += rows.map(st => tipRow(st.name, st.values[i], st.color)).join('');
    if (spec.stacks.length > 1 && rows.length) h += `<div class="tip-r tot"><span>${esc(spec.totalLabel || 'Totale')}</span><span class="num">${fmt(sums[i])}</span></div>`;
    if (!rows.length) h += `<div class="muted">${esc(spec.emptyLabel || 'Nessuna spesa')}</div>`;
    if (spec.line && spec.line.values[i] != null) h += tipRow(spec.line.name, spec.line.values[i], spec.line.color, true);
    return h;
  });
}

// Linee (confronto tra anni)
function drawLines(el, spec) {
  const all = spec.series.flatMap(se => se.values.filter(v => v != null));
  const { g, s: frame } = chartFrame(el, spec, Math.max(0, ...all));
  let s = frame + `<line class="xh" x1="0" x2="0" y1="${g.M.t}" y2="${g.M.t + g.ph}" stroke="var(--axis)" stroke-width="1" visibility="hidden"/>`;
  for (const se of spec.series) {
    let d = '', pen = false, last = null;
    se.values.forEach((v, i) => {
      if (v == null) { pen = false; return; }
      d += `${pen ? 'L' : 'M'}${g.X(i)},${g.Y(v)}`;
      pen = true;
      last = i;
    });
    s += `<path d="${d}" fill="none" stroke="${se.color}" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>`;
    if (last != null) s += `<circle cx="${g.X(last)}" cy="${g.Y(se.values[last])}" r="4" fill="${se.color}" stroke="var(--card)" stroke-width="2"/>`;
  }
  s += '<g class="hd"></g>';
  el.innerHTML = s + '</svg><div class="tip" hidden></div>';
  const xh = $('.xh', el), hd = $('.hd', el);
  bindTip(el, g, i => {
    const rows = spec.series.filter(se => se.values[i] != null).slice().reverse();
    return `<div class="tip-h">${esc(spec.fullLabels ? spec.fullLabels[i] : spec.labels[i])}</div>` +
      (rows.length ? rows.map(se => tipRow(se.name, se.values[i], se.color, true)).join('') : '<div class="muted">Nessun dato</div>');
  }, i => {
    if (i < 0) { xh.setAttribute('visibility', 'hidden'); hd.innerHTML = ''; return; }
    const x = Math.round(g.X(i)) + 0.5;
    xh.setAttribute('x1', x);
    xh.setAttribute('x2', x);
    xh.setAttribute('visibility', 'visible');
    hd.innerHTML = spec.series.filter(se => se.values[i] != null)
      .map(se => `<circle cx="${g.X(i)}" cy="${g.Y(se.values[i])}" r="4" fill="${se.color}" stroke="var(--card)" stroke-width="2"/>`).join('');
  });
}

function hbarsHTML(items, total) {
  const max = Math.max(1, ...items.map(it => it.value));
  return `<div class="hbars">${items.map(it => `
    <div class="hbar"${it.trend ? ` data-trend="${esc(it.trend)}" role="button" tabindex="0" style="cursor:pointer"` : ''}>
      <div class="hbar-top"><span class="lbl"><i class="dot" style="background:${it.color}"></i>${esc(it.label)}</span>
        <span class="val">${fmt(it.value)}<span class="pct">${pct(it.value, total)}%</span></span></div>
      <div class="hbar-track"><div class="hbar-fill" style="width:${Math.max(0, (it.value / max) * 100)}%;background:${it.color}"></div></div>
    </div>`).join('')}</div>`;
}
const legendHTML = items => `<div class="legend">${items.map(it =>
  `<span><i class="${it.line ? 'key-line' : 'dot'}" style="background:${it.color}"></i>${esc(it.name)}</span>`).join('')}</div>`;

// ============================================================
// Vista: Riepilogo
// ============================================================
function viewSummary() {
  const tabs = [['mese', 'Mese'], ['anno', 'Anno'], ['andamento', 'Andamento']];
  let body = '';
  if (!state.entries.length) {
    body = '<div class="empty"><strong>Ancora nessun dato</strong>Registra qualche spesa o importa l\'estratto conto per vedere il riepilogo.</div>';
  } else if (ui.sumTab === 'mese') body = summaryMonthHTML();
  else if (ui.sumTab === 'anno') body = summaryYearHTML();
  else body = summaryTrendHTML();
  main.innerHTML = `<div class="seg" role="tablist" style="margin-bottom:12px">${tabs.map(([k, l]) =>
    `<button type="button" role="tab" aria-selected="${ui.sumTab === k}" class="${ui.sumTab === k ? 'on' : ''}" data-stab="${k}">${l}</button>`).join('')}</div>${body}`;

  main.onclick = ev => {
    const t = ev.target.closest('[data-stab]');
    if (t) { ui.sumTab = t.dataset.stab; render(); return; }
    const nav = ev.target.closest('[data-nav]');
    if (nav) {
      if (ui.sumTab === 'mese') ui.month = addMonths(ui.month, +nav.dataset.nav);
      else ui.year += +nav.dataset.nav;
      render();
      return;
    }
    const tr = ev.target.closest('[data-trend]');
    if (tr) { ui.trend.target = tr.dataset.trend; ui.sumTab = 'andamento'; render(); window.scrollTo(0, 0); return; }
    const ex = ev.target.closest('[data-expand]');
    if (ex) { const id = ex.dataset.expand; ui.expanded.has(id) ? ui.expanded.delete(id) : ui.expanded.add(id); render(); return; }
    const go = ev.target.closest('[data-goto-uncl]');
    if (go) { ui.listFilter = 'uncl'; ui.search = ''; ui.month = go.dataset.gotoUncl; location.hash = '#movimenti'; return; }
    const tm = ev.target.closest('[data-tmode]');
    if (tm) { ui.trend.mode = tm.dataset.tmode; render(); }
  };
  main.onchange = ev => {
    if (ev.target.id === 'showZero') { ui.showZero = ev.target.checked; render(); }
    else if (ev.target.id === 'trendTarget') { ui.trend.target = ev.target.value; render(); }
    else if (ev.target.id === 'trendRange') { ui.trend.range = ev.target.value; render(); }
  };
  main.oninput = null;
}

function summaryMonthHTML() {
  const ym = ui.month;
  const list = entriesIn(ym);
  const t = totals(list);
  const prevYm = addMonths(ym, -1);
  const prev = totals(entriesIn(prevYm));
  const bySub = new Map(), byCat = new Map(), byInc = new Map();
  let incUncl = 0;
  for (const e of list) {
    if (isUnclassified(e)) { if (e.type === 'inc') incUncl += e.amount; continue; }
    if (e.type === 'exp') {
      bySub.set(`${e.cat}:${e.sub}`, (bySub.get(`${e.cat}:${e.sub}`) || 0) + e.amount);
      byCat.set(e.cat, (byCat.get(e.cat) || 0) + e.amount);
    } else byInc.set(e.inc, (byInc.get(e.inc) || 0) + e.amount);
  }
  let delta = '';
  // Il confronto ha senso solo se anche il mese precedente è stato registrato per intero
  if (prev.exp > 0 && entriesIn(prevYm).length >= 10) {
    const d = Math.round(((t.exp - prev.exp) / prev.exp) * 100);
    delta = `${d > 0 ? '▲' : d < 0 ? '▼' : '='} ${Math.abs(d)}% su ${MONTHS[+prevYm.slice(5) - 1].toLowerCase()}`;
  }
  const bars = state.categories.filter(c => (byCat.get(c.id) || 0) > 0)
    .map(c => ({ label: c.name, value: byCat.get(c.id), color: catColor(c), trend: `exp:cat:${c.id}` }))
    .sort((a, b) => b.value - a.value);
  if (t.uncl > 0) bars.push({ label: 'Da classificare', value: t.uncl, color: 'var(--s0)' });

  let rows = '';
  for (const c of state.categories) {
    const tot = byCat.get(c.id) || 0;
    const subs = c.subs.filter(s => ui.showZero || bySub.get(`${c.id}:${s.id}`));
    if (!tot && !ui.showZero) continue;
    rows += `<tr class="cat-row click" data-trend="exp:cat:${esc(c.id)}"><td><span class="cat-cell"><i class="dot" style="background:${catColor(c)}"></i>${esc(c.name)}</span></td><td class="num">${fmt(tot)}</td></tr>`;
    rows += subs.map(s => {
      const v = bySub.get(`${c.id}:${s.id}`) || 0;
      return `<tr class="sub-row click${v ? '' : ' zero'}" data-trend="exp:sub:${esc(c.id)}:${esc(s.id)}"><td>${esc(s.name)}</td><td class="num">${v ? fmt(v) : '–'}</td></tr>`;
    }).join('');
  }
  if (t.uncl) rows += `<tr class="cat-row click" data-goto-uncl="${ym}"><td><span class="cat-cell"><i class="dot" style="background:var(--s0)"></i>Da classificare</span></td><td class="num">${fmt(t.uncl)}</td></tr>`;

  const incRows = state.incomeTypes.filter(it => ui.showZero || byInc.get(it.id)).map(it =>
    `<tr class="sub-row click${byInc.get(it.id) ? '' : ' zero'}" data-trend="inc:type:${esc(it.id)}"><td>${esc(it.name)}</td><td class="num">${byInc.get(it.id) ? fmt(byInc.get(it.id)) : '–'}</td></tr>`).join('')
    + (incUncl ? `<tr class="sub-row click" data-goto-uncl="${ym}"><td>Da classificare</td><td class="num">${fmt(incUncl)}</td></tr>` : '');

  return `${monthNavHTML(monthName(ym))}
    <div class="tiles">
      ${tile('Entrate', fmtTile(t.inc), '', 'pos')}
      ${tile('Spese', fmtTile(t.exp), delta)}
      ${tile('Bilancio', fmtTile(t.bal, true), t.inc ? `${t.bal >= 0 ? 'Risparmiato' : 'Sforato'} ${Math.abs(pct(t.bal, t.inc))}% delle entrate` : '', t.bal < 0 ? 'neg' : 'pos')}
    </div>
    ${t.unclCount ? `<div class="notice">${ICON.warn}<div>${t.unclCount} movimenti da classificare${t.uncl ? ` (${fmt(t.uncl)} di spese)` : ''}. <button type="button" class="linkish" data-goto-uncl="${ym}">Classificali ora</button></div></div>` : ''}
    ${list.length ? `
    <section class="card"><h2>Spese per categoria</h2>
      ${bars.length ? hbarsHTML(bars, t.exp) : '<p class="hint">Nessuna spesa in questo mese.</p>'}
      ${bars.length ? '<p class="hint small">Tocca una categoria per vederne l\'andamento.</p>' : ''}
    </section>
    <section class="card"><h2>Spese mensili</h2>
      <table class="sum-table"><thead><tr><th>Categoria / voce</th><th class="num">Totale</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="2" class="muted">Nessuna spesa</td></tr>'}
        <tr class="total-row"><td>Spese totali</td><td class="num">${fmt(t.exp)}</td></tr></tbody></table>
      <label class="check"><input type="checkbox" id="showZero"${ui.showZero ? ' checked' : ''}><span>Mostra anche le voci a zero</span></label>
    </section>
    <section class="card"><h2>Entrate</h2>
      <table class="sum-table"><tbody>${incRows || '<tr><td colspan="2" class="muted">Nessuna entrata registrata in questo mese</td></tr>'}
        <tr class="total-row"><td>Entrate totali</td><td class="num">${fmt(t.inc)}</td></tr>
        <tr><td><b>Bilancio del mese</b></td><td class="num"><b class="${t.bal < 0 ? 'neg' : 'pos'}">${fmtSigned(t.bal)}</b></td></tr></tbody></table>
    </section>
    <section class="card"><h2>Come hai pagato</h2>
      <div class="stat-line"><span>Contanti <b class="num">${fmt(t.cash)}</b></span><span>Conto / carta <b class="num">${fmt(t.bank)}</b></span></div>
    </section>` : '<div class="empty"><strong>Nessun movimento</strong>In questo mese non ci sono spese né entrate.</div>'}`;
}

function yearAgg(y) {
  const arr = () => Array(12).fill(0);
  const A = { cat: {}, sub: {}, inc: {}, uncl: arr(), incUncl: arr(), exp: arr(), incTot: arr(), active: Array(12).fill(false) };
  for (const e of state.entries) {
    if (!e.date.startsWith(y + '-')) continue;
    const m = +e.date.slice(5, 7) - 1;
    A.active[m] = true;
    if (e.type === 'exp') {
      A.exp[m] += e.amount;
      if (isUnclassified(e)) A.uncl[m] += e.amount;
      else {
        (A.cat[e.cat] ||= arr())[m] += e.amount;
        (A.sub[`${e.cat}:${e.sub}`] ||= arr())[m] += e.amount;
      }
    } else {
      A.incTot[m] += e.amount;
      if (isUnclassified(e)) A.incUncl[m] += e.amount;
      else (A.inc[e.inc] ||= arr())[m] += e.amount;
    }
  }
  return A;
}

function summaryYearHTML() {
  const y = ui.year;
  const A = yearAgg(y);
  // Il mese in corso non è ancora completo: non entra nelle medie (se ci sono altri mesi)
  const curM = String(y) === thisMonth().slice(0, 4) ? +thisMonth().slice(5) - 1 : -1;
  const complete = A.active.map((on, i) => on && i !== curM);
  const nAct = complete.some(Boolean) ? complete.filter(Boolean).length : A.active.filter(Boolean).length;
  const totExp = sum(A.exp), totInc = sum(A.incTot), bal = totInc - totExp;
  const labels = MONTHS_SHORT.map(m => m[0]);
  const stacks = state.categories.filter(c => A.cat[c.id] && sum(A.cat[c.id]))
    .map(c => ({ name: c.name, color: catColor(c), values: A.cat[c.id] }));
  if (sum(A.uncl)) stacks.push({ name: 'Da classificare', color: 'var(--s0)', values: A.uncl });
  const incLine = totInc ? { name: 'Entrate', color: 'var(--ink-2)', values: A.incTot.map(v => v || null) } : null;

  const cell = v => (v ? `<td>${fmtInt(v)}</td>` : '<td class="zero">–</td>');
  const avg = a => (nAct ? Math.round(sum(a.filter((_, i) => complete[i] || !complete.some(Boolean))) / nAct) : 0);
  const line = (label, a, cls = '', attrs = '') =>
    `<tr class="${cls}"${attrs}><td><div class="yl">${label}</div></td>${a.map(cell).join('')}<td class="col-total">${fmtInt(sum(a))}</td><td>${fmtInt(avg(a))}</td></tr>`;
  let body = '';
  for (const c of state.categories) {
    const a = A.cat[c.id] || Array(12).fill(0);
    const open = ui.expanded.has(c.id);
    body += line(`<span class="cat-cell"><i class="dot" style="background:${catColor(c)}"></i>${esc(c.name)} <span class="muted">${open ? '▾' : '▸'}</span></span>`, a, 'cat-row', ` data-expand="${esc(c.id)}" style="cursor:pointer"`);
    if (open) for (const s of c.subs) body += line(esc(s.name), A.sub[`${c.id}:${s.id}`] || Array(12).fill(0), 'sub-row', ` data-trend="exp:sub:${esc(c.id)}:${esc(s.id)}" style="cursor:pointer"`);
  }
  if (sum(A.uncl)) body += line('Da classificare', A.uncl, 'cat-row');
  body += line('Spese totali', A.exp, 'total-row');
  for (const it of state.incomeTypes) if (A.inc[it.id] && sum(A.inc[it.id])) body += line(esc(it.name), A.inc[it.id], 'sub-row', ` data-trend="inc:type:${esc(it.id)}" style="cursor:pointer"`);
  if (sum(A.incUncl)) body += line('Entrate da classificare', A.incUncl, 'sub-row');
  body += line('Entrate totali', A.incTot, 'total-row');
  const balA = A.incTot.map((v, i) => v - A.exp[i]);
  body += `<tr class="total-row"><td>Bilancio</td>${balA.map(v => (v ? `<td class="${v < 0 ? 'neg' : 'pos'}">${fmtInt(v)}</td>` : '<td class="zero">–</td>')).join('')}<td class="${bal < 0 ? 'neg' : 'pos'}">${fmtInt(bal)}</td><td>${fmtInt(avg(balA))}</td></tr>`;

  return `${monthNavHTML(String(y))}
    <div class="tiles two">
      ${tile('Entrate dell\'anno', fmtTile(totInc), '', 'pos')}
      ${tile('Spese dell\'anno', fmtTile(totExp))}
      ${tile('Bilancio dell\'anno', fmtTile(bal, true), totInc ? `${bal >= 0 ? 'Risparmiato' : 'Sforato'} ${Math.abs(pct(bal, totInc))}% delle entrate` : '', bal < 0 ? 'neg' : 'pos')}
      ${tile('Spesa media al mese', fmtTile(avg(A.exp)), nAct ? `su ${nAct} ${nAct === 1 ? 'mese completo' : 'mesi completi'}` : '')}
    </div>
    ${nAct ? `<section class="card"><h2>Spese e entrate mese per mese</h2>
      ${legendHTML([...stacks.map(s => ({ name: s.name, color: s.color })), ...(incLine ? [{ name: 'Entrate', color: incLine.color, line: true }] : [])])}
      ${chartSlot({ labels, fullLabels: MONTHS.map(m => `${m} ${y}`), stacks, line: incLine, totalLabel: 'Spese totali', height: 230, aria: `Spese per categoria e entrate, mese per mese, ${y}` })}
    </section>
    <section class="card"><h2>Tabella dell'anno <span class="muted small" style="font-weight:400">(euro, tocca una categoria per il dettaglio)</span></h2>
      <div class="scroll-x"><table class="year-table"><thead><tr><th>Voce</th>${MONTHS_SHORT.map(m => `<th>${m}</th>`).join('')}<th>Totale</th><th>Media</th></tr></thead>
      <tbody>${body}</tbody></table></div>
    </section>` : '<div class="empty"><strong>Nessun movimento</strong>In questo anno non ci sono dati.</div>'}`;
}

// --- Andamento di una categoria / sottocategoria ---
function matchesTarget(e, target) {
  const p = target.split(':');
  if (p[0] === 'exp') {
    if (e.type !== 'exp') return false;
    if (p[1] === 'all') return true;
    if (p[1] === 'cat') return e.cat === p[2] && !isUnclassified(e);
    if (p[1] === 'sub') return e.cat === p[2] && e.sub === p[3];
  }
  if (p[0] === 'inc') {
    if (e.type !== 'inc') return false;
    if (p[1] === 'all') return true;
    if (p[1] === 'type') return e.inc === p[2];
  }
  return false;
}
function trendInfo(target) {
  const p = target.split(':');
  if (p[0] === 'exp' && p[1] === 'cat' && catById(p[2])) return { name: catById(p[2]).name, color: catColor(catById(p[2])) };
  if (p[0] === 'exp' && p[1] === 'sub' && subOf(p[2], p[3])) return { name: `${catById(p[2]).name} › ${subOf(p[2], p[3]).name}`, color: catColor(catById(p[2])) };
  if (p[0] === 'inc' && p[1] === 'type' && incById(p[2])) return { name: incById(p[2]).name, color: 'var(--accent)' };
  if (target === 'inc:all') return { name: 'Tutte le entrate', color: 'var(--accent)' };
  return { name: 'Tutte le spese', color: 'var(--accent)' };
}
function trendTargetOptions(sel) {
  const opt = (v, l) => `<option value="${esc(v)}"${v === sel ? ' selected' : ''}>${esc(l)}</option>`;
  let h = `<optgroup label="Totali">${opt('exp:all', 'Tutte le spese')}${opt('inc:all', 'Tutte le entrate')}</optgroup>`;
  for (const c of state.categories) h += `<optgroup label="${esc(c.name)}">${opt(`exp:cat:${c.id}`, `${c.name} (tutta la categoria)`)}${c.subs.map(s => opt(`exp:sub:${c.id}:${s.id}`, s.name)).join('')}</optgroup>`;
  h += `<optgroup label="Entrate">${state.incomeTypes.map(t => opt(`inc:type:${t.id}`, t.name)).join('')}</optgroup>`;
  return h;
}

function summaryTrendHTML() {
  const tr = ui.trend;
  if (!targetValidTrend(tr.target)) tr.target = 'exp:all';
  const info = trendInfo(tr.target);
  const byMonth = new Map();
  const activeMonths = new Set();
  for (const e of state.entries) {
    const ym = e.date.slice(0, 7);
    activeMonths.add(ym);
    if (matchesTarget(e, tr.target)) byMonth.set(ym, (byMonth.get(ym) || 0) + e.amount);
  }
  const allMonths = [...activeMonths].sort();
  const first = allMonths[0], cur = thisMonth();
  const last = allMonths[allMonths.length - 1] > cur ? allMonths[allMonths.length - 1] : cur;
  const years = [];
  for (let y = +first.slice(0, 4); y <= +last.slice(0, 4); y++) years.push(y);

  const modes = [['mesi', 'Mesi'], ['anni', 'Anni'], ['confronto', 'Confronta anni']];
  let h = `<section class="card">
    <label class="field" style="margin-top:0"><span>Cosa vuoi analizzare</span><select id="trendTarget">${trendTargetOptions(tr.target)}</select></label>
    <div class="seg" style="margin-top:12px">${modes.map(([k, l]) => `<button type="button" class="${tr.mode === k ? 'on' : ''}" data-tmode="${k}">${l}</button>`).join('')}</div>`;

  if (tr.mode === 'mesi') {
    const ranges = [['12', 'Ultimi 12 mesi'], ['24', 'Ultimi 24 mesi'], ...years.slice().reverse().map(y => [`Y${y}`, `Anno ${y}`]), ['all', 'Dall\'inizio']];
    if (!ranges.some(([k]) => k === tr.range)) tr.range = '12';
    let months = [];
    if (tr.range === 'all') { for (let m = first; m <= last; m = addMonths(m, 1)) months.push(m); }
    else if (tr.range.startsWith('Y')) { for (let i = 1; i <= 12; i++) months.push(`${tr.range.slice(1)}-${pad(i)}`); }
    else { for (let i = +tr.range - 1; i >= 0; i--) months.push(addMonths(cur, -i)); }
    const vals = months.map(m => byMonth.get(m) || 0);
    const total = sum(vals);
    // Media sui mesi con dati, escluso il mese in corso (ancora incompleto) se ce ne sono altri
    let counted = months.filter(m => activeMonths.has(m) && m !== cur);
    if (!counted.length) counted = months.filter(m => activeMonths.has(m));
    const avg = counted.length ? Math.round(sum(counted.map(m => byMonth.get(m) || 0)) / counted.length) : 0;
    const maxI = vals.indexOf(Math.max(...vals));
    h += `<label class="field"><span>Periodo</span><select id="trendRange">${ranges.map(([k, l]) => `<option value="${k}"${k === tr.range ? ' selected' : ''}>${esc(l)}</option>`).join('')}</select></label></section>
      <div class="tiles">
        ${tile('Totale periodo', fmtTile(total))}
        ${tile('Media al mese', fmtTile(avg), counted.length ? `su ${counted.length} ${counted.length === 1 ? 'mese' : 'mesi'} con dati` : '')}
        ${tile('Mese più alto', total ? fmtTile(vals[maxI]) : '–', total ? monthName(months[maxI]) : '')}
      </div>
      <section class="card"><h2>${esc(info.name)}: mese per mese</h2>
        ${total ? chartSlot({ labels: months.map(m => (months.length > 12 ? monthShort(m) : MONTHS_SHORT[+m.slice(5) - 1])), fullLabels: months.map(monthName), stacks: [{ name: info.name, color: info.color, values: vals }], avg, emptyLabel: 'Nessun importo', aria: `Andamento mensile di ${info.name}` })
          : '<p class="hint">Nessun importo in questo periodo.</p>'}
      </section>
      ${total ? `<section class="card"><h2>Valori</h2><table class="sum-table"><tbody>${months.slice().reverse().map(m => {
        const v = byMonth.get(m) || 0;
        return `<tr class="${v ? '' : 'sub-row zero'}"><td>${monthName(m)}</td><td class="num">${v ? fmt(v) : '–'}</td></tr>`;
      }).join('')}</tbody></table></section>` : ''}`;
  } else if (tr.mode === 'anni') {
    const rows = years.map(y => {
      const ms = [...Array(12)].map((_, i) => `${y}-${pad(i + 1)}`);
      const tot = sum(ms.map(m => byMonth.get(m) || 0));
      let done = ms.filter(m => activeMonths.has(m) && m !== cur);
      if (!done.length) done = ms.filter(m => activeMonths.has(m));
      const n = done.length;
      return { y, tot, n, avg: n ? Math.round(sum(done.map(m => byMonth.get(m) || 0)) / n) : 0 };
    });
    h += `</section>
      <section class="card"><h2>${esc(info.name)}: anno per anno</h2>
        ${chartSlot({ labels: rows.map(r => String(r.y)), stacks: [{ name: info.name, color: info.color, values: rows.map(r => r.tot) }], emptyLabel: 'Nessun importo', aria: `Totale annuo di ${info.name}` })}
        <p class="hint small">Se un anno è incompleto, confronta la media al mese più che il totale.</p>
      </section>
      <section class="card"><table class="sum-table"><thead><tr><th>Anno</th><th class="num">Totale</th><th class="num">Media al mese</th></tr></thead><tbody>
        ${rows.slice().reverse().map(r => `<tr><td>${r.y}${r.n && r.n < 12 ? ` <span class="muted small">(${r.n} mesi)</span>` : ''}</td><td class="num">${fmt(r.tot)}</td><td class="num">${fmt(r.avg)}</td></tr>`).join('')}
      </tbody></table></section>`;
  } else {
    const shown = years.slice(-6);
    // Scala ordinata: l'anno più recente è il più marcato; i mesi senza dati (prima dell'inizio o non ancora finiti) restano vuoti
    const series = shown.map((y, k) => ({
      name: String(y),
      color: `var(--y${shown.length === 1 ? 6 : Math.round(1 + (k * 5) / (shown.length - 1))})`,
      values: [...Array(12)].map((_, i) => {
        const m = `${y}-${pad(i + 1)}`;
        return (m >= cur || m < first) && !byMonth.has(m) ? null : byMonth.get(m) || 0;
      }),
    }));
    h += `</section>
      <section class="card"><h2>${esc(info.name)}: confronto tra anni</h2>
        ${legendHTML(series.slice().reverse().map(s => ({ name: s.name, color: s.color, line: true })))}
        ${chartSlot({ kind: 'lines', labels: MONTHS_SHORT.map(m => m[0]), fullLabels: MONTHS, series, aria: `Confronto annuale di ${info.name}` })}
      </section>
      <section class="card"><div class="scroll-x"><table class="year-table"><thead><tr><th>Mese</th>${shown.slice().reverse().map(y => `<th>${y}</th>`).join('')}</tr></thead><tbody>
        ${MONTHS.map((mn, i) => `<tr><td>${mn}</td>${shown.slice().reverse().map(y => { const v = byMonth.get(`${y}-${pad(i + 1)}`) || 0; return v ? `<td>${fmtInt(v)}</td>` : '<td class="zero">–</td>'; }).join('')}</tr>`).join('')}
        <tr class="total-row"><td>Totale</td>${shown.slice().reverse().map(y => `<td>${fmtInt(sum(series.find(s => s.name === String(y)).values.map(v => v || 0)))}</td>`).join('')}</tr>
      </tbody></table></div></section>`;
  }
  return h;
}
function targetValidTrend(t) {
  const p = String(t).split(':');
  if (t === 'exp:all' || t === 'inc:all') return true;
  if (p[0] === 'exp' && p[1] === 'cat') return !!catById(p[2]);
  if (p[0] === 'exp' && p[1] === 'sub') return !!subOf(p[2], p[3]);
  if (p[0] === 'inc' && p[1] === 'type') return !!incById(p[2]);
  return false;
}

// ============================================================
// Vista: Altro (impostazioni, backup, installazione)
// ============================================================
async function startInstall() {
  if (installPrompt) {
    installPrompt.prompt();
    await installPrompt.userChoice.catch(() => {});
    installPrompt = null;
    render();
  } else {
    location.hash = '#altro';
  }
}

function installCardHTML() {
  if (isStandalone()) return `<section class="card"><h2>App installata</h2><p class="hint" style="margin:0">Stai usando l'app installata sul dispositivo.</p></section>`;
  return `<section class="card"><h2>Installa sul telefono</h2>
    ${installPrompt ? '<button type="button" class="btn primary block" data-act="install" style="margin-bottom:10px">Installa l\'app</button>' : ''}
    <p class="small" style="margin:0"><b>iPhone</b> (con Safari):</p>
    <ol class="steps small"><li>Apri questa pagina con <b>Safari</b>.</li><li>Tocca il pulsante <b>Condividi</b> (il quadrato con la freccia in su).</li><li>Scegli <b>Aggiungi alla schermata Home</b> e conferma.</li></ol>
    <p class="small" style="margin:10px 0 0"><b>Android</b> (con Chrome):</p>
    <ol class="steps small"><li>Apri questa pagina con <b>Chrome</b>.</li><li>Tocca il menu <b>⋮</b> in alto a destra.</li><li>Scegli <b>Installa app</b> (o <b>Aggiungi a schermata Home</b>).</li></ol>
  </section>`;
}

function viewSettings() {
  const last = state.settings.lastBackup;
  const catsHTML = state.categories.map(c => {
    const used = state.entries.filter(e => e.type === 'exp' && e.cat === c.id).length;
    return `<details class="cat-edit" data-cat-details="${esc(c.id)}"${ui.openCats.has(c.id) ? ' open' : ''}>
      <summary><i class="dot" style="background:${catColor(c)}"></i>${esc(c.name)} <span class="muted small">${c.subs.length} voci</span><span class="chev">${ICON.chev}</span></summary>
      <div class="edit-line"><input value="${esc(c.name)}" data-rename-cat="${esc(c.id)}" aria-label="Nome categoria"><button type="button" class="btn small danger" data-del-cat="${esc(c.id)}"${used ? ` title="${used} movimenti"` : ''}>Elimina</button></div>
      ${c.subs.map(s => `<div class="edit-line" style="padding-left:14px"><input value="${esc(s.name)}" data-rename-sub="${esc(c.id)}:${esc(s.id)}" aria-label="Nome sottocategoria"><button type="button" class="icon-btn" data-del-sub="${esc(c.id)}:${esc(s.id)}" aria-label="Elimina ${esc(s.name)}">${ICON.close}</button></div>`).join('')}
      <div class="edit-line" style="padding-left:14px"><input placeholder="Nuova sottocategoria" data-new-sub="${esc(c.id)}"><button type="button" class="btn small" data-add-sub="${esc(c.id)}">Aggiungi</button></div>
    </details>`;
  }).join('');
  const incHTML = state.incomeTypes.map(t => `<div class="edit-line"><input value="${esc(t.name)}" data-rename-inc="${esc(t.id)}" aria-label="Nome voce di entrata"><button type="button" class="icon-btn" data-del-inc="${esc(t.id)}" aria-label="Elimina ${esc(t.name)}">${ICON.close}</button></div>`).join('');
  const rules = state.rules.map((r, i) => ({ ...r, i })).filter(r => targetValid(r.target)).sort((a, b) => a.key.localeCompare(b.key));

  main.innerHTML = `${installCardHTML()}
    <section class="card"><h2>Backup e dati</h2>
      <p class="hint" style="margin:0 0 10px">I dati sono salvati solo su questo dispositivo, nel browser. ${last ? `Ultimo backup: <b>${fmtDate(last)}</b>.` : '<b>Non hai ancora fatto un backup.</b>'}</p>
      <div class="btn-row"><button type="button" class="btn primary" data-act="backup">Esporta backup</button><button type="button" class="btn" data-act="csv">Esporta per Excel (CSV)</button></div>
      <div class="btn-row" style="margin-top:8px"><button type="button" class="btn" data-act="merge">Unisci da backup</button><button type="button" class="btn" data-act="restore">Ripristina backup</button></div>
      <p class="hint small">«Unisci» aggiunge i movimenti di un altro telefono (es. le spese in contanti registrate da Sara) senza creare doppioni. «Ripristina» sostituisce tutti i dati con quelli del backup.</p>
      <input type="file" id="backupIn" accept=".json,application/json" hidden>
    </section>
    <h2 class="section-title">Categorie di spesa</h2>
    <section class="card">${catsHTML}<button type="button" class="btn block" data-act="add-cat" style="margin-top:10px">+ Nuova categoria</button></section>
    <h2 class="section-title">Voci di entrata</h2>
    <section class="card">${incHTML}<div class="edit-line"><input placeholder="Nuova voce di entrata" id="newInc"><button type="button" class="btn small" data-act="add-inc">Aggiungi</button></div></section>
    <h2 class="section-title">Regole di classificazione</h2>
    <section class="card">
      <p class="hint" style="margin:0 0 8px">L'app le impara quando correggi una categoria durante l'importazione. Puoi anche aggiungerne una: se la descrizione del movimento contiene la parola, viene classificato così.</p>
      ${rules.length ? rules.map(r => `<div class="rule-line"><div class="r-main"><code>${esc(r.key)}</code> → ${esc(targetLabel(r.target))}</div><button type="button" class="icon-btn" data-del-rule="${r.i}" aria-label="Elimina regola">${ICON.close}</button></div>`).join('') : '<p class="muted small">Nessuna regola personale per ora.</p>'}
      <div style="margin-top:12px"><input id="ruleKey" placeholder="Parola da cercare (es. ATTILIO)" style="margin-bottom:8px">
        <select id="ruleTarget" style="margin-bottom:8px">${targetOptions('', { ignore: true, emptyLabel: '— Scegli la categoria —' })}</select>
        <button type="button" class="btn block" data-act="add-rule">Aggiungi regola</button></div>
    </section>
    <section class="card"><h2>Zona pericolosa</h2>
      <button type="button" class="btn danger block" data-act="wipe">Cancella tutti i dati</button></section>
    <p class="hint small" style="text-align:center">Spese di Famiglia · versione ${APP_VERSION}</p>`;

  $$('details.cat-edit', main).forEach(d => {
    d.ontoggle = () => { d.open ? ui.openCats.add(d.dataset.catDetails) : ui.openCats.delete(d.dataset.catDetails); };
  });

  main.onclick = ev => {
    const b = ev.target.closest('button');
    if (!b) return;
    const d = b.dataset, act = d.act;
    if (act === 'install') startInstall();
    else if (act === 'backup') exportBackup();
    else if (act === 'csv') exportCSV();
    else if (act === 'merge' || act === 'restore') { backupMode = act; $('#backupIn').click(); }
    else if (act === 'add-cat') {
      const name = (prompt('Nome della nuova categoria') || '').trim();
      if (!name) return;
      const used = new Set(state.categories.map(c => c.slot));
      const slot = [1, 2, 3, 4, 5, 6, 7, 8].find(s => !used.has(s)) || 0;
      const id = uid();
      state.categories.push({ id, name, slot, subs: [] });
      ui.openCats.add(id);
      save(); render();
    } else if (d.addSub) {
      const input = $(`[data-new-sub="${CSS.escape(d.addSub)}"]`, main);
      const name = input.value.trim();
      if (!name) { input.focus(); return; }
      catById(d.addSub).subs.push({ id: uid(), name });
      save(); render();
    } else if (d.delSub) {
      const [cId, sId] = d.delSub.split(':');
      const used = state.entries.filter(e => e.cat === cId && e.sub === sId);
      const name = subOf(cId, sId).name;
      if (!confirm(used.length ? `«${name}» contiene ${used.length} movimenti, che diventeranno «da classificare». Eliminare la sottocategoria?` : `Eliminare la sottocategoria «${name}»?`)) return;
      used.forEach(e => { e.cat = null; e.sub = null; });
      const c = catById(cId);
      c.subs = c.subs.filter(s => s.id !== sId);
      state.rules = state.rules.filter(r => r.target !== `exp:${cId}:${sId}`);
      save(); render();
    } else if (d.delCat) {
      const c = catById(d.delCat);
      const used = state.entries.filter(e => e.type === 'exp' && e.cat === c.id);
      if (!confirm(used.length ? `«${c.name}» contiene ${used.length} movimenti, che diventeranno «da classificare». Eliminare la categoria?` : `Eliminare la categoria «${c.name}»?`)) return;
      used.forEach(e => { e.cat = null; e.sub = null; });
      state.categories = state.categories.filter(x => x.id !== c.id);
      state.rules = state.rules.filter(r => !r.target.startsWith(`exp:${c.id}:`));
      save(); render();
    } else if (act === 'add-inc') {
      const name = $('#newInc').value.trim();
      if (!name) return;
      state.incomeTypes.push({ id: uid(), name });
      save(); render();
    } else if (d.delInc) {
      const t = incById(d.delInc);
      const used = state.entries.filter(e => e.type === 'inc' && e.inc === t.id);
      if (!confirm(used.length ? `«${t.name}» contiene ${used.length} entrate, che diventeranno «da classificare». Eliminare la voce?` : `Eliminare la voce «${t.name}»?`)) return;
      used.forEach(e => { e.inc = null; });
      state.incomeTypes = state.incomeTypes.filter(x => x.id !== t.id);
      state.rules = state.rules.filter(r => r.target !== `inc:${t.id}`);
      save(); render();
    } else if (d.delRule !== undefined) {
      state.rules.splice(+d.delRule, 1);
      save(); render();
    } else if (act === 'add-rule') {
      const key = compact($('#ruleKey').value);
      const target = $('#ruleTarget').value;
      if (key.length < 3) { alert('Scrivi una parola di almeno 3 lettere.'); return; }
      if (!target) { alert('Scegli la categoria.'); return; }
      learnRule(key, target);
      save(); render();
      toast('Regola aggiunta: varrà dai prossimi import');
    } else if (act === 'wipe') {
      if (!confirm('Cancellare TUTTI i movimenti, le categorie e le regole da questo dispositivo?')) return;
      if (!confirm('Sei sicuro? Se non hai un backup, i dati non si possono recuperare.')) return;
      state = defaultState();
      save(); render();
      toast('Tutti i dati sono stati cancellati');
    }
  };
  main.onchange = async ev => {
    const el = ev.target, d = el.dataset;
    if (el.id === 'backupIn' && el.files[0]) {
      try { await importBackup(el.files[0], backupMode); }
      catch (err) { alert(err.message || 'File di backup non valido.'); }
      el.value = '';
      return;
    }
    const name = el.value.trim();
    if (d.renameCat && name) { catById(d.renameCat).name = name; save(); render(); }
    else if (d.renameSub && name) { const [c, s] = d.renameSub.split(':'); subOf(c, s).name = name; save(); }
    else if (d.renameInc && name) { incById(d.renameInc).name = name; save(); }
  };
  main.oninput = null;
}

async function downloadOrShare(name, text, type) {
  const file = new File([text], name, { type });
  if (isMobile && navigator.canShare && navigator.canShare({ files: [file] })) {
    try { await navigator.share({ files: [file], title: name }); return true; }
    catch (err) { if (err.name === 'AbortError') return false; }
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(file);
  a.download = name;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  return true;
}

async function exportBackup() {
  const data = JSON.stringify({ app: 'spese-famiglia', version: 1, exported: new Date().toISOString(), state });
  if (await downloadOrShare(`spese-famiglia-backup-${today()}.json`, data, 'application/json')) {
    state.settings.lastBackup = today();
    save();
    render();
    toast('Backup esportato: conservalo in un posto sicuro (Drive, email…)');
  }
}

function exportCSV() {
  const head = ['Data', 'Tipo', 'Categoria', 'Sottocategoria / voce', 'Importo', 'Pagamento', 'Nota', 'Descrizione banca'];
  const rows = [...state.entries].sort((a, b) => a.date.localeCompare(b.date)).map(e => [
    fmtDate(e.date),
    e.type === 'exp' ? 'Spesa' : 'Entrata',
    e.type === 'exp' ? catById(e.cat)?.name || 'Da classificare' : 'Entrate',
    e.type === 'exp' ? subOf(e.cat, e.sub)?.name || '' : incById(e.inc)?.name || 'Da classificare',
    (e.amount / 100).toFixed(2).replace('.', ','),
    e.type === 'exp' ? (e.method === 'contanti' ? 'Contanti' : 'Conto / carta') : '',
    e.note || '',
    e.bankDesc || '',
  ]);
  const csv = '﻿' + [head, ...rows].map(r => r.map(v => (/[;"\n\r]/.test(String(v)) ? `"${String(v).replace(/"/g, '""')}"` : v)).join(';')).join('\r\n');
  downloadOrShare(`spese-famiglia-movimenti-${today()}.csv`, csv, 'text/csv');
}

async function importBackup(file, mode) {
  const data = JSON.parse(await file.text());
  const s = data.state || data;
  if (!s || !Array.isArray(s.entries) || !Array.isArray(s.categories)) throw new Error('Questo file non è un backup di Spese di Famiglia.');
  if (mode === 'restore') {
    if (!confirm(`Il backup contiene ${s.entries.length} movimenti. Sostituire tutti i dati attuali?`)) return;
    state = migrate(s);
    save(); render();
    toast('Backup ripristinato');
    return;
  }
  const ids = new Set(state.entries.map(e => e.id));
  const keys = new Set(state.entries.map(e => e.bankKey).filter(Boolean));
  let added = 0;
  for (const e of s.entries) {
    if (ids.has(e.id) || (e.bankKey && keys.has(e.bankKey))) continue;
    state.entries.push(e);
    added++;
  }
  for (const c of s.categories) {
    const mine = catById(c.id);
    if (!mine) {
      const used = new Set(state.categories.map(x => x.slot));
      state.categories.push({ ...c, slot: used.has(c.slot) ? [1, 2, 3, 4, 5, 6, 7, 8].find(x => !used.has(x)) || 0 : c.slot });
    } else for (const sub of c.subs) if (!mine.subs.some(x => x.id === sub.id)) mine.subs.push(sub);
  }
  for (const t of s.incomeTypes || []) if (!incById(t.id)) state.incomeTypes.push(t);
  for (const r of s.rules || []) if (!state.rules.some(x => x.key === r.key)) state.rules.push(r);
  state.ignoredKeys = [...new Set([...state.ignoredKeys, ...(s.ignoredKeys || [])])];
  save(); render();
  toast(added === 1 ? 'Aggiunto 1 movimento dal backup' : `Aggiunti ${added} movimenti dal backup`);
}

// ============================================================
// Navigazione
// ============================================================
const VIEWS = { aggiungi: viewAdd, movimenti: viewList, importa: viewImport, riepilogo: viewSummary, altro: viewSettings };
const TITLES = { aggiungi: 'Nuovo movimento', movimenti: 'Movimenti', importa: 'Importa dalla banca', riepilogo: 'Riepilogo', altro: 'Altro' };

function render() {
  VIEWS[ui.tab]();
  mountCharts();
}
function route() {
  closeSheet(true);
  const tab = location.hash.slice(1);
  ui.tab = VIEWS[tab] ? tab : 'aggiungi';
  $$('.tabbar a').forEach(a => {
    const on = a.dataset.tab === ui.tab;
    a.classList.toggle('on', on);
    if (on) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
  });
  $('#title').textContent = TITLES[ui.tab];
  render();
  window.scrollTo(0, 0);
}
window.addEventListener('hashchange', route);
sheet.addEventListener('click', ev => { if (ev.target.matches('.sheet-backdrop')) closeSheet(); });

window.addEventListener('beforeinstallprompt', ev => {
  ev.preventDefault();
  installPrompt = ev;
  if (ui.tab === 'altro') render();
});
window.addEventListener('appinstalled', () => { installPrompt = null; render(); });

if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  let hadController = !!navigator.serviceWorker.controller;
  navigator.serviceWorker.addEventListener('controllerchange', () => {
    if (hadController) toast('È disponibile una nuova versione dell\'app', 'Aggiorna', () => location.reload());
    hadController = true;
  });
  window.addEventListener('load', () => navigator.serviceWorker.register('sw.js').catch(err => console.warn('Service worker non registrato', err)));
}

route();
