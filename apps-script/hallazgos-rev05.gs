/**
 * Web App — RG-06-01 Rev 05 · Hallazgos SGI
 * Lee la planilla Rev05 y expone los datos como JSON para hallazgos.html.
 * También envía avisos por mail cada lunes para vencimientos.
 *
 * INSTALACIÓN:
 *  1. Ir a script.google.com con la cuenta coordinacion.sgi@esinel.com.ar
 *  2. Crear proyecto "SGI - Hallazgos Rev05 API"
 *  3. Pegar este archivo y guardar.
 *  4. Implementar → Nueva implementación → Tipo: Aplicación web
 *     Ejecutar como: Yo · Quién tiene acceso: Cualquier persona
 *  5. Copiar la URL /exec y pegarla en serverWebAppUrl dentro de hallazgos.html
 *  6. Para los avisos de lunes: Activadores → Agregar activador
 *     Función: enviarAvisosLunes · Tipo de evento: Basado en tiempo → Semana → Lunes → 7-8 AM
 */

const SHEET_ID_REV05  = '1uzI_at19LwjTtYUekEUJoXF2W_5vu32hrgSB4866Z1E';
const SHEET_NAME      = 'Planilla';
const REMITENTE_ALIAS = 'SGI Esinel — Hallazgos';

// Días de anticipación para "próximo a vencer"
const DIAS_ALERTA_PREVIO = 30;

// Mapa nombre → correo (sincronizar con hallazgos.html si se agrega gente)
const CORREOS = {
  'thiago baira':        'coordinacion.sgi@esinel.com.ar',
  'mercedes nahas':      'mnahas@esinel.com.ar',
  'clarisa piccione':    'cpiccione@esinel.com.ar',
  'lisandro fissolo':    'lfissolo@esinel.com.ar',
  'martín diamanti':     'mdiamanti@esinel.com.ar',
  'martin diamanti':     'mdiamanti@esinel.com.ar',
  'marcelo piccione':    'mpiccione@esinel.com.ar',
  'hernan giacobbe':     'giacobbeh@esinel.com.ar',
  'hernán giacobbe':     'giacobbeh@esinel.com.ar',
  'oscar schmüth':       'oschmuth@esinel.com.ar',
  'oscar schmuth':       'oschmuth@esinel.com.ar',
  'juan monti':          'jmonti@esinel.com.ar',
  'patricia rodriguez':  'prodriguez@esinel.com.ar',
  'patricia rodríguez':  'prodriguez@esinel.com.ar',
  'daniel lértora':      'compras@esinel.com.ar',
  'daniel lertora':      'compras@esinel.com.ar',
  'salud y seguridad':   'saludyseguridad@esinel.com.ar',
  'deposito esinel':     'deposito@esinel.com.ar',
  'depósito esinel':     'deposito@esinel.com.ar',
  'sergio donda':        'sdonda@esinel.com.ar',
  'diego fernandez':     'dfernandez@esinel.com.ar',
  'diego fernández':     'dfernandez@esinel.com.ar',
  'carlos de abrantes':  'cdeabrantes@esinel.com.ar',
  'andrea chavez':       'achavez@esinel.com.ar',
  'nahuel heredia':      'nheredia@esinel.com.ar',
  'anahí gonzalez':      'imagen@esinel.com.ar',
  'anahi gonzalez':      'imagen@esinel.com.ar',
  'emiliano san martin': 'esanmartin@esinel.com.ar',
  'emiliano san martín': 'esanmartin@esinel.com.ar',
  'tomás noblia':        'tnoblia@esinel.com.ar',
  'tomas noblia':        'tnoblia@esinel.com.ar',
  'hernán giorgetti':    'hgiorgetti@esinel.com.ar',
  'hernan giorgetti':    'hgiorgetti@esinel.com.ar',
  'florencia oyuela':    'foyuela@esinel.com.ar'
};

// ── Helpers ─────────────────────────────────────────────────────────────────

function getHoja() {
  return SpreadsheetApp.openById(SHEET_ID_REV05).getSheetByName(SHEET_NAME);
}

function formatDateCell(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const d = val.getDate().toString().padStart(2,'0');
    const m = (val.getMonth()+1).toString().padStart(2,'0');
    const y = val.getFullYear();
    return `${d}/${m}/${y}`;
  }
  return String(val).trim();
}

function rowToItem(row) {
  const tipo = String(row[0]  || '').trim();
  const num  = String(row[1]  || '').trim();
  return {
    tipoHallazgo:       tipo,
    numHallazgo:        tipo && num ? tipo + '-' + num : num,
    quienDetecto:       String(row[2]  || '').trim(),
    fechaDetection:     formatDateCell(row[3]),
    descripcion:        String(row[4]  || '').trim(),
    accionInmediata:    String(row[5]  || '').trim(),
    analisisCausa:      String(row[6]  || '').trim(),
    sectorResponsable:  String(row[7]  || '').trim(),
    correspondeAccion:  String(row[8]  || '').trim(),
    tipoAccion:         String(row[9]  || '').trim(),
    numAccion:          String(row[10] || '').trim(),
    accionCorrectiva:   String(row[11] || '').trim(),
    accionPreventiva:   String(row[12] || '').trim(),
    accionMejora:       String(row[13] || '').trim(),
    evaluacionResultado:String(row[14] || '').trim(),
    responsableAccion:  String(row[15] || '').trim(),
    fechaLimite:        formatDateCell(row[16]),
    fechaRealCierre:    formatDateCell(row[17]),
    responsableCierre:  String(row[18] || '').trim(),
    evidenciaCierre:    String(row[19] || '').trim(),
    fechaVerificacion:  formatDateCell(row[20]),
    resultadoEficacia:  String(row[21] || '').trim(),
    evidenciaResultado: String(row[22] || '').trim(),
    requiereActualizar: String(row[23] || '').trim(),
    fuenteEvidencia:    String(row[24] || '').trim(),
    estado:             String(row[25] || '').trim(),
    categoria:          String(row[26] || '').trim()
  };
}

function getAllData() {
  const hoja = getHoja();
  const vals = hoja.getDataRange().getValues();
  if (vals.length <= 1) return [];
  return vals.slice(1)
    .filter(r => r[0] || r[1] || r[4])  // skip blank rows
    .map(rowToItem);
}

// ── Web App endpoints ────────────────────────────────────────────────────────

function doGet(e) {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('hallazgos_rev05');
  let data;
  if (cached) {
    data = JSON.parse(cached);
  } else {
    data = getAllData();
    cache.put('hallazgos_rev05', JSON.stringify(data), 300); // 5 min
  }
  const payload = JSON.stringify({ status: 'success', data: data });
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || 'save';

    if (action === 'save') {
      const hoja = getHoja();
      const item = body.item || {};
      const tipo = (item.tipoHallazgo || '').trim();
      const num  = (item.numHallazgo  || '').replace(tipo + '-', '').trim();
      hoja.appendRow([
        tipo, num,
        item.quienDetecto, item.fechaDetection, item.descripcion,
        item.accionInmediata, item.analisisCausa, item.sectorResponsable,
        item.correspondeAccion, item.tipoAccion, item.numAccion,
        item.accionCorrectiva, item.accionPreventiva, item.accionMejora,
        item.evaluacionResultado, item.responsableAccion,
        item.fechaLimite, item.fechaRealCierre, item.responsableCierre,
        item.evidenciaCierre, item.fechaVerificacion, item.resultadoEficacia,
        item.evidenciaResultado, item.requiereActualizar, item.fuenteEvidencia,
        item.estado, item.categoria
      ]);
      CacheService.getScriptCache().remove('hallazgos_rev05');
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'update') {
      const hoja   = getHoja();
      const item   = body.item || {};
      const target = (item.numHallazgo || '').trim();
      const vals   = hoja.getDataRange().getValues();
      for (let i = 1; i < vals.length; i++) {
        const rowCode = String(vals[i][0]).trim() + '-' + String(vals[i][1]).trim();
        if (rowCode === target) {
          const tipo = (item.tipoHallazgo || '').trim();
          const num  = target.replace(tipo + '-', '').trim();
          hoja.getRange(i + 1, 1, 1, 27).setValues([[
            tipo, num,
            item.quienDetecto, item.fechaDetection, item.descripcion,
            item.accionInmediata, item.analisisCausa, item.sectorResponsable,
            item.correspondeAccion, item.tipoAccion, item.numAccion,
            item.accionCorrectiva, item.accionPreventiva, item.accionMejora,
            item.evaluacionResultado, item.responsableAccion,
            item.fechaLimite, item.fechaRealCierre, item.responsableCierre,
            item.evidenciaCierre, item.fechaVerificacion, item.resultadoEficacia,
            item.evidenciaResultado, item.requiereActualizar, item.fuenteEvidencia,
            item.estado, item.categoria
          ]]);
          break;
        }
      }
      CacheService.getScriptCache().remove('hallazgos_rev05');
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Para google.script.run (cuando se usa como sidebar)
function getSheetDataRaw() { return getAllData(); }
function getWebAppUrl()     { return ScriptApp.getService().getUrl(); }
function getSpreadsheetUrl(){ return SpreadsheetApp.openById(SHEET_ID_REV05).getUrl(); }

// ── Avisos automáticos de lunes ──────────────────────────────────────────────
/**
 * Configurar como activador semanal:
 *   Activadores → Agregar → enviarAvisosLunes → Basado en tiempo → Semana → Lunes → 7-8 AM
 */
function enviarAvisosLunes() {
  const hoy       = new Date();
  const limite    = new Date(hoy);
  limite.setDate(hoy.getDate() + DIAS_ALERTA_PREVIO);

  const data = getAllData();

  // Agrupar por correo → lista de hallazgos a avisar
  const porCorreo = {};

  data.forEach(item => {
    // Abierto = fecha real de cierre vacía (col R)
    const frc = (item.fechaRealCierre || '').trim();
    const abierto = !frc || frc === '-';
    if (!abierto) return;

    const fechaLimite = parseDateGAS(item.fechaLimite);
    if (!fechaLimite) return;

    const vencido  = fechaLimite < hoy;
    const proximo  = !vencido && fechaLimite <= limite;
    if (!vencido && !proximo) return;

    const responsables = splitEntidades(item.responsableAccion)
      .concat(splitEntidades(item.responsableCierre));

    const unicos = [...new Set(responsables)];
    unicos.forEach(nombre => {
      const email = resolverCorreo(nombre);
      if (!email) return;
      if (!porCorreo[email]) porCorreo[email] = { nombre, vencidos: [], proximos: [] };
      if (vencido) porCorreo[email].vencidos.push(item);
      else         porCorreo[email].proximos.push(item);
    });
  });

  const fechaStr = hoy.toLocaleDateString('es-AR', { weekday:'long', day:'2-digit', month:'long', year:'numeric' });

  Object.entries(porCorreo).forEach(([email, info]) => {
    const total = info.vencidos.length + info.proximos.length;
    if (total === 0) return;

    const asunto = `[SGI Esinel] Aviso de hallazgos — ${info.vencidos.length} vencidos · ${info.proximos.length} por vencer`;

    let cuerpo = `<div style="font-family:Arial,sans-serif;font-size:14px;color:#1e293b;max-width:680px">
<div style="background:#0f172a;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0">
  <h2 style="margin:0;font-size:18px">📋 SGI Esinel — Aviso de Hallazgos</h2>
  <p style="margin:4px 0 0;font-size:12px;opacity:.7">${fechaStr}</p>
</div>
<div style="border:1px solid #e2e8f0;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px">
  <p>Hola <strong>${info.nombre}</strong>,</p>
  <p>Este es tu recordatorio semanal de hallazgos SGI que requieren tu atención:</p>`;

    if (info.vencidos.length > 0) {
      cuerpo += `<h3 style="color:#dc2626;margin-top:20px">🔴 Vencidos (${info.vencidos.length})</h3>
<table style="width:100%;border-collapse:collapse;font-size:12px">
<tr style="background:#fef2f2"><th style="padding:6px 10px;text-align:left;border:1px solid #fecaca">Código</th><th style="padding:6px 10px;text-align:left;border:1px solid #fecaca">Hallazgo</th><th style="padding:6px 10px;text-align:left;border:1px solid #fecaca">Vencía</th></tr>`;
      info.vencidos.forEach(h => {
        cuerpo += `<tr><td style="padding:6px 10px;border:1px solid #fee2e2;font-weight:bold;color:#dc2626">${h.numHallazgo}</td><td style="padding:6px 10px;border:1px solid #fee2e2">${(h.descripcion||'').substring(0,80)}${h.descripcion&&h.descripcion.length>80?'…':''}</td><td style="padding:6px 10px;border:1px solid #fee2e2;white-space:nowrap">${h.fechaLimite}</td></tr>`;
      });
      cuerpo += `</table>`;
    }

    if (info.proximos.length > 0) {
      cuerpo += `<h3 style="color:#d97706;margin-top:20px">🟡 Vencen en los próximos ${DIAS_ALERTA_PREVIO} días (${info.proximos.length})</h3>
<table style="width:100%;border-collapse:collapse;font-size:12px">
<tr style="background:#fffbeb"><th style="padding:6px 10px;text-align:left;border:1px solid #fde68a">Código</th><th style="padding:6px 10px;text-align:left;border:1px solid #fde68a">Hallazgo</th><th style="padding:6px 10px;text-align:left;border:1px solid #fde68a">Vence</th><th style="padding:6px 10px;text-align:left;border:1px solid #fde68a">Días</th></tr>`;
      info.proximos.forEach(h => {
        const fl = parseDateGAS(h.fechaLimite);
        const diasRestantes = fl ? Math.ceil((fl - hoy) / (1000*60*60*24)) : '?';
        cuerpo += `<tr><td style="padding:6px 10px;border:1px solid #fef3c7;font-weight:bold;color:#d97706">${h.numHallazgo}</td><td style="padding:6px 10px;border:1px solid #fef3c7">${(h.descripcion||'').substring(0,80)}${h.descripcion&&h.descripcion.length>80?'…':''}</td><td style="padding:6px 10px;border:1px solid #fef3c7;white-space:nowrap">${h.fechaLimite}</td><td style="padding:6px 10px;border:1px solid #fef3c7;font-weight:bold;color:#d97706;white-space:nowrap">${diasRestantes}d</td></tr>`;
      });
      cuerpo += `</table>`;
    }

    cuerpo += `<p style="margin-top:20px">Accedé al dashboard para ver el detalle y actualizar el estado:</p>
<a href="https://coordinacionsgi-prog.github.io/esinel-sgi/hallazgos.html" style="display:inline-block;background:#0284c7;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold">Ir al Dashboard de Hallazgos →</a>
<p style="margin-top:24px;font-size:11px;color:#64748b">Este mensaje fue generado automáticamente por el SGI Esinel. No responder a este correo.</p>
</div></div>`;

    MailApp.sendEmail({
      to: email,
      cc: 'coordinacion.sgi@esinel.com.ar',
      subject: asunto,
      htmlBody: cuerpo,
      name: REMITENTE_ALIAS
    });
  });

  Logger.log('Avisos enviados: ' + Object.keys(porCorreo).length + ' destinatarios');
}

// ── Helpers internos ─────────────────────────────────────────────────────────

function parseDateGAS(str) {
  if (!str || str === '-' || str === '') return null;
  if (str instanceof Date) return str;
  const s = String(str).trim();
  // dd/mm/yyyy
  const p = s.split('/');
  if (p.length === 3 && p[2].length === 4) return new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
  // yyyy-mm-dd
  const p2 = s.split('-');
  if (p2.length === 3 && p2[0].length === 4) return new Date(parseInt(p2[0]), parseInt(p2[1])-1, parseInt(p2[2]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function splitEntidades(str) {
  if (!str) return [];
  return str.split(/[,/|&]/).map(s => s.trim()).filter(Boolean);
}

function resolverCorreo(nombre) {
  if (!nombre) return null;
  const key = nombre.toLowerCase().trim();
  return CORREOS[key] || null;
}
