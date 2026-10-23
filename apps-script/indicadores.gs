/**
 * Web App que expone los indicadores de "Graficos indicadores compaginado" como JSON
 * para el dashboard indicadores.html.
 *
 * Instalado en el proyecto "SGI - Indicadores API" (cuenta coordinacion.sgi@esinel.com.ar).
 * Al cambiar el código: Implementar → Gestionar implementaciones → editar → Nueva versión,
 * así la URL /exec no cambia.
 *
 * Instalación desde cero:
 *  1. Crear un proyecto en script.google.com con la cuenta dueña de la hoja.
 *  2. Pegar este archivo completo y guardar.
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web.
 *     Ejecutar como: Yo · Quién tiene acceso: Cualquier persona.
 *  4. Copiar la URL que termina en /exec y pegarla en APPS_SCRIPT_WEBAPP_URL de indicadores.html.
 *
 * Cada pestaña cuyo nombre sea un código (COME1, INGE2, AMB4…) se lee como un indicador.
 * Se aceptan dos formatos: con ficha arriba y tabla "Ejercicio | Periodo | Puntaje | … | Objetivo",
 * o directamente "Periodo | Puntaje | … | Objetivo" desde la fila 1.
 *
 * Dashboard propio: escribir DASHBOARD en cualquier celda de las primeras 15 filas de la pestaña
 * y el link en la celda de la derecha (o como hipervínculo de esa misma celda).
 */

const SPREADSHEET_ID = '16WEwU8Dhd7uJvLDYlCIobri5WarPPVHDtXpjHdbjqSw';
const CACHE_SECONDS = 300;
const CODE_RE = /^[A-Z]+\d+$/;

const AREAS = {
  COME: 'Comercial', INGE: 'Ingeniería', COMP: 'Compras', ARMTAB: 'Producción',
  OBRA: 'Obra', RRHH: 'RRHH', AMB: 'Medio Ambiente', SAL: 'Seguridad e Higiene'
};

function doGet(e) {
  const nocache = e && e.parameter && e.parameter.nocache;
  const cache = CacheService.getScriptCache();
  let json = nocache ? null : cache.get('indicadores');
  if (!json) {
    json = JSON.stringify({ updated: new Date().toISOString(), indicadores: leerIndicadores() });
    try { cache.put('indicadores', json, CACHE_SECONDS); } catch (err) { /* > 100 KB: sin cache */ }
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function leerIndicadores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tz = ss.getSpreadsheetTimeZone();
  return ss.getSheets()
    .filter(sh => CODE_RE.test(sh.getName().trim()))
    .map(sh => leerPestana(sh, tz))
    .filter(ind => ind.periodos.length > 0);
}

function leerPestana(sh, tz) {
  const code = sh.getName().trim();
  const range = sh.getDataRange();
  const vals = range.getValues();
  const disp = range.getDisplayValues();
  const head = Math.min(15, vals.length);

  const ind = {
    codigo: code,
    area: AREAS[code.replace(/\d+$/, '')] || '',
    titulo: '', objetivo: '', metodo: '', frecuencia: '', responsable: '', acciones: '',
    dashboard: '',
    periodos: []
  };

  const t = String(disp[0][0] || '');
  if (t.indexOf('·') > -1) ind.titulo = t.split('·')[1].trim();

  const acciones = [];
  for (let r = 0; r < head; r++) {
    const label = norm(disp[r][0]);
    const val = String(disp[r][2] || '').trim();
    if (label === 'OBJETIVO') ind.objetivo = val;
    else if (label === 'METODO DE MONITOREO') ind.metodo = val;
    else if (label === 'FRECUENCIA DE MEDICION') ind.frecuencia = val;
    else if (label === 'RESPONSABLE') ind.responsable = val;
    else if (/^\d+\s*→\s*\d+$/.test(String(disp[r][0]).trim()) && val) {
      const ej = String(disp[r][0]).replace(/\s/g, '').replace('→', '->');
      const obj = String(disp[r][5] || '').trim();
      acciones.push('Ej.' + ej + ': ' + val + (obj ? ' (objetivo: ' + obj + ')' : ''));
    }
  }
  ind.acciones = acciones.join(' | ');
  ind.dashboard = buscarDashboard(sh, disp, head);

  let hr = -1, cPer = -1, cPun = -1, cObj = -1, cEj = -1;
  for (let r = 0; r < Math.min(25, vals.length) && hr < 0; r++) {
    for (let c = 0; c < disp[r].length; c++) {
      if (norm(disp[r][c]) === 'PERIODO') { hr = r; break; }
    }
  }
  if (hr < 0) return ind;
  disp[hr].forEach((h, c) => {
    const n = norm(h);
    if (n === 'PERIODO') cPer = c;
    else if (n === 'PUNTAJE') cPun = c;
    else if (n === 'OBJETIVO') cObj = c;
    else if (n === 'EJERCICIO') cEj = c;
  });

  let lastObj = null;
  for (let r = hr + 1; r < vals.length; r++) {
    const fecha = aIso(vals[r][cPer], disp[r][cPer], tz);
    if (!fecha) {
      if (!String(disp[r][cPer]).trim()) break;
      continue;
    }
    const obj = cObj > -1 ? aNum(vals[r][cObj]) : null;
    if (obj !== null) lastObj = obj;
    const ej = cEj > -1 ? parseInt(disp[r][cEj], 10) : NaN;
    ind.periodos.push([fecha, aNum(vals[r][cPun]), lastObj, isNaN(ej) ? null : ej]);
  }
  return ind;
}

function buscarDashboard(sh, disp, head) {
  for (let r = 0; r < head; r++) {
    for (let c = 0; c < disp[r].length; c++) {
      if (norm(disp[r][c]).indexOf('DASHBOARD') === -1) continue;
      const own = linkDe(sh.getRange(r + 1, c + 1));
      if (own) return own;
      for (let k = c + 1; k < disp[r].length; k++) {
        const txt = String(disp[r][k] || '').trim();
        if (/^https?:\/\//.test(txt)) return txt;
        const link = linkDe(sh.getRange(r + 1, k + 1));
        if (link) return link;
      }
    }
  }
  return '';
}

function linkDe(cell) {
  const rt = cell.getRichTextValue();
  if (rt && rt.getLinkUrl()) return rt.getLinkUrl();
  const m = String(cell.getFormula() || '').match(/HYPERLINK\("([^"]+)"/i);
  return m ? m[1] : '';
}

function norm(s) {
  return String(s || '').trim().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
}

function aNum(v) {
  if (typeof v === 'number') return Math.round(v * 100) / 100;
  const s = String(v || '').trim().replace(/\s*pts?$/i, '').replace(',', '.');
  if (!s || s === '-') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : Math.round(n * 100) / 100;
}

function aIso(v, txt, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
  const m = String(txt || '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (!m) return null;
  const y = m[3].length === 2 ? '20' + m[3] : m[3];
  return y + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}
