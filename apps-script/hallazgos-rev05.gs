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
  'florencia oyuela':    'foyuela@esinel.com.ar',
  'marisa cercos':       'mcercos@esinel.com.ar'
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
    const action = body.action || '';

    // La planilla web es de solo lectura: el alta y la edición se hacen en Google Sheets.
    if (action === 'notificar') {
      return jsonOut(notificarHallazgo(body));
    }
    return jsonOut({ status: 'error', message: 'Solo lectura: la planilla se edita en Google Sheets' });

  } catch(err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Notificación manual desde la web: nuevo hallazgo, cierre y eficacia ─────
// Solo se envía a los Responsables de Cierre (col S) de ese hallazgo, con copia a
// Coordinación SGI. El contenido lo arma el script con los datos de la planilla:
// la web solo indica qué hallazgo y qué tipo de aviso.
const TIPOS_NOTIFICACION = {
  NUEVO:    { titulo: 'Nuevo hallazgo asignado',  color: '#0284c7',
              intro: 'Se te asignó como responsable de cierre del siguiente hallazgo del Sistema de Gestión Integrado:' },
  CIERRE:   { titulo: 'Cierre del hallazgo',      color: '#059669',
              intro: 'Se registró el cierre del siguiente hallazgo del que sos responsable:' },
  EFICACIA: { titulo: 'Verificación de eficacia', color: '#7c3aed',
              intro: 'Se registró la verificación de eficacia del siguiente hallazgo del que sos responsable:' }
};
const MINUTOS_ENTRE_AVISOS = 10;   // evita envíos repetidos del mismo aviso

function notificarHallazgo(body) {
  const tipo = String(body.tipo || '').toUpperCase();
  const cfg  = TIPOS_NOTIFICACION[tipo];
  if (!cfg) return { status: 'error', message: 'Tipo de aviso inválido' };

  const num  = String(body.numHallazgo || '').trim();
  const item = getAllData().find(it => it.numHallazgo === num);
  if (!item) return { status: 'error', message: 'No se encontró el hallazgo ' + num };

  // Responsable de Cierre = columna S de la planilla
  const destinatarios = [...new Set(splitEntidades(item.responsableCierre)
    .map(resolverCorreo).filter(Boolean))];
  if (!destinatarios.length) {
    return { status: 'error', message: item.responsableCierre && item.responsableCierre !== '-'
      ? 'El Responsable de Cierre (' + item.responsableCierre + ') no tiene correo registrado'
      : 'El hallazgo no tiene Responsable de Cierre cargado en la planilla (columna S)' };
  }

  const cache = CacheService.getScriptCache();
  const clave = 'notif_' + tipo + '_' + num;
  if (cache.get(clave)) {
    return { status: 'error', message: 'Este aviso ya se envió hace menos de ' + MINUTOS_ENTRE_AVISOS + ' minutos' };
  }

  const numero = num.indexOf(item.tipoHallazgo + '-') === 0 ? num.slice(item.tipoHallazgo.length + 1) : num;
  const fila = (k, v, extra) => '<tr><td style="padding:8px 12px;font-weight:bold;border-bottom:1px solid #e2e8f0;width:38%;color:#475569">' + k +
    '</td><td style="padding:8px 12px;border-bottom:1px solid #e2e8f0;' + (extra || '') + '">' + escHtml(v || '-') + '</td></tr>';
  let filas = fila('Hallazgo', item.tipoHallazgo + ' ' + numero, 'color:' + cfg.color + ';font-weight:bold')
    + fila('Sector responsable', item.sectorResponsable)
    + fila('Fecha de detección', item.fechaDetection)
    + fila('Fecha límite de cierre', item.fechaLimite, 'color:#dc2626;font-weight:bold');
  if (tipo !== 'NUEVO') filas += fila('Fecha real de cierre', item.fechaRealCierre) + fila('Evidencia de cierre', item.evidenciaCierre);
  if (tipo === 'EFICACIA') filas += fila('Fecha de verificación', item.fechaVerificacion)
    + fila('Resultado de eficacia', item.resultadoEficacia, 'font-weight:bold;color:' + (String(item.resultadoEficacia).toUpperCase() === 'OK' ? '#059669' : '#dc2626'))
    + fila('Evidencia del resultado', item.evidenciaResultado);
  const plan = [item.accionCorrectiva, item.accionPreventiva, item.accionMejora]
    .filter(v => v && v !== '-').join('\n') || '-';

  const html = '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:650px;margin:0 auto;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;color:#334155">' +
    '<div style="background:' + cfg.color + ';padding:20px;text-align:center;color:#fff"><h2 style="margin:0;font-size:17px">' + cfg.titulo.toUpperCase() + '</h2>' +
    '<p style="margin:4px 0 0;font-size:12px">RG-06-01 Rev 05 · Sistema de Gestión Integrado</p></div>' +
    '<div style="padding:22px"><p style="font-size:13px;margin-top:0">' + cfg.intro + '</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:14px 0;font-size:12px;background:#f8fafc;border:1px solid #cbd5e1">' + filas + '</table>' +
    '<h4 style="font-size:11px;text-transform:uppercase;color:#64748b;margin:0 0 6px">Descripción</h4>' +
    '<div style="background:#f1f5f9;padding:12px;border-radius:8px;font-size:12px;white-space:pre-wrap">' + escHtml(item.descripcion) + '</div>' +
    '<h4 style="font-size:11px;text-transform:uppercase;color:#64748b;margin:14px 0 6px">Plan de acción (' + escHtml(item.tipoAccion) + ')</h4>' +
    '<div style="background:#f1f5f9;padding:12px;border-radius:8px;font-size:12px;white-space:pre-wrap">' + escHtml(plan) + '</div>' +
    '<p style="text-align:center;margin:24px 0 8px"><a href="https://coordinacionsgi-prog.github.io/esinel-sgi/hallazgos.html" style="display:inline-block;background:#0284c7;color:#fff;text-decoration:none;font-size:12px;font-weight:bold;padding:10px 20px;border-radius:8px">Ver en el Dashboard de Hallazgos</a></p>' +
    '<p style="margin-top:20px;padding-top:14px;border-top:1px solid #e2e8f0;font-size:11px;color:#94a3b8;text-align:center">Coordinación SGI · ESINEL S.R.L. · Mensaje automático, no responder.</p></div></div>';

  MailApp.sendEmail({
    to:       destinatarios.join(','),
    cc:       'coordinacion.sgi@esinel.com.ar',
    subject:  '[SGI Esinel] ' + cfg.titulo + ': ' + item.tipoHallazgo + ' ' + numero,
    htmlBody: html,
    name:     REMITENTE_ALIAS
  });
  cache.put(clave, '1', MINUTOS_ENTRE_AVISOS * 60);
  Logger.log('notificar ' + tipo + ' ' + num + ' -> ' + destinatarios.join(', ') + ' (pedido por ' + (body.solicitante || '?') + ')');
  return { status: 'success', destinatarios: destinatarios };
}

function escHtml(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
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

// ── Aviso semanal: eficacia pendiente ────────────────────────────────────────
/**
 * Configurar como activador semanal (puede ser el mismo lunes o día diferente):
 *   Activadores → Agregar → enviarAvisosEficacia → Basado en tiempo → Semana → Lunes → 7-8 AM
 *
 * Envía a coordinacion.sgi un resumen de todos los hallazgos cerrados
 * que aún no tienen resultadoEficacia (OK o NO) cargado.
 * También avisa individualmente a cada responsable si tiene pendientes.
 */
function enviarAvisosEficacia() {
  const hoy  = new Date();
  const data = getAllData();

  // Cerrados sin eficacia: fechaRealCierre no vacío y resultadoEficacia vacío
  const sinEficacia = data.filter(item => {
    const frc = (item.fechaRealCierre || '').trim();
    if (!frc || frc === '-') return false; // abierto → no aplica
    const ef  = (item.resultadoEficacia || '').trim().toUpperCase();
    return ef !== 'OK' && ef !== 'NO';
  });

  if (sinEficacia.length === 0) {
    Logger.log('enviarAvisosEficacia: sin pendientes');
    return;
  }

  const fechaStr = hoy.toLocaleDateString('es-AR', { weekday:'long', day:'2-digit', month:'long', year:'numeric' });

  // ── 1. Resumen global a coordinacion.sgi ────────────────────────────────
  let tablaGlobal = `<table style="width:100%;border-collapse:collapse;font-size:12px">
<tr style="background:#f0f9ff"><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Código</th><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Descripción</th><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Cierre</th><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Responsable</th></tr>`;
  sinEficacia.forEach(h => {
    tablaGlobal += `<tr>
      <td style="padding:6px 10px;border:1px solid #e0f2fe;font-weight:bold;color:#0284c7">${h.numHallazgo}</td>
      <td style="padding:6px 10px;border:1px solid #e0f2fe">${(h.descripcion||'').substring(0,80)}${h.descripcion&&h.descripcion.length>80?'…':''}</td>
      <td style="padding:6px 10px;border:1px solid #e0f2fe;white-space:nowrap">${h.fechaRealCierre}</td>
      <td style="padding:6px 10px;border:1px solid #e0f2fe">${h.responsableAccion||h.responsableCierre||'—'}</td>
    </tr>`;
  });
  tablaGlobal += `</table>`;

  const cuerpoCoord = `<div style="font-family:Arial,sans-serif;font-size:14px;color:#1e293b;max-width:720px">
<div style="background:#0f172a;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0">
  <h2 style="margin:0;font-size:18px">📋 SGI Esinel — Eficacia Pendiente</h2>
  <p style="margin:4px 0 0;font-size:12px;opacity:.7">${fechaStr}</p>
</div>
<div style="border:1px solid #e2e8f0;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px">
  <p>Resumen semanal de hallazgos <strong>cerrados sin verificación de eficacia</strong>:</p>
  <p style="font-size:20px;font-weight:bold;color:#0284c7">${sinEficacia.length} hallazgos</p>
  ${tablaGlobal}
  <p style="margin-top:20px">Ver detalle en el dashboard:</p>
  <a href="https://coordinacionsgi-prog.github.io/esinel-sgi/hallazgos.html" style="display:inline-block;background:#0284c7;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold">Ir al Dashboard →</a>
  <p style="margin-top:24px;font-size:11px;color:#64748b">Generado automáticamente por SGI Esinel.</p>
</div></div>`;

  MailApp.sendEmail({
    to:       'coordinacion.sgi@esinel.com.ar',
    subject:  `[SGI Esinel] ${sinEficacia.length} hallazgos cerrados sin eficacia verificada`,
    htmlBody: cuerpoCoord,
    name:     REMITENTE_ALIAS
  });

  // ── 2. Aviso individual a cada responsable ───────────────────────────────
  const porCorreo = {};
  sinEficacia.forEach(item => {
    const responsables = splitEntidades(item.responsableAccion)
      .concat(splitEntidades(item.responsableCierre));
    [...new Set(responsables)].forEach(nombre => {
      const email = resolverCorreo(nombre);
      if (!email || email === 'coordinacion.sgi@esinel.com.ar') return;
      if (!porCorreo[email]) porCorreo[email] = { nombre, items: [] };
      porCorreo[email].items.push(item);
    });
  });

  Object.entries(porCorreo).forEach(([email, info]) => {
    let tabla = `<table style="width:100%;border-collapse:collapse;font-size:12px">
<tr style="background:#f0f9ff"><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Código</th><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Descripción</th><th style="padding:6px 10px;text-align:left;border:1px solid #bae6fd">Fecha Cierre</th></tr>`;
    info.items.forEach(h => {
      tabla += `<tr>
        <td style="padding:6px 10px;border:1px solid #e0f2fe;font-weight:bold;color:#0284c7">${h.numHallazgo}</td>
        <td style="padding:6px 10px;border:1px solid #e0f2fe">${(h.descripcion||'').substring(0,80)}${h.descripcion&&h.descripcion.length>80?'…':''}</td>
        <td style="padding:6px 10px;border:1px solid #e0f2fe;white-space:nowrap">${h.fechaRealCierre}</td>
      </tr>`;
    });
    tabla += `</table>`;

    const cuerpo = `<div style="font-family:Arial,sans-serif;font-size:14px;color:#1e293b;max-width:680px">
<div style="background:#0f172a;color:#fff;padding:20px 24px;border-radius:8px 8px 0 0">
  <h2 style="margin:0;font-size:18px">📋 SGI Esinel — Verificación de Eficacia Pendiente</h2>
  <p style="margin:4px 0 0;font-size:12px;opacity:.7">${fechaStr}</p>
</div>
<div style="border:1px solid #e2e8f0;border-top:none;padding:20px 24px;border-radius:0 0 8px 8px">
  <p>Hola <strong>${info.nombre}</strong>,</p>
  <p>Los siguientes hallazgos de tu sector están cerrados pero aún <strong>no tienen la eficacia verificada</strong> (falta completar el resultado: OK o NO):</p>
  ${tabla}
  <p style="margin-top:16px">Por favor, ingresá al dashboard para registrar el resultado de la verificación de eficacia:</p>
  <a href="https://coordinacionsgi-prog.github.io/esinel-sgi/hallazgos.html" style="display:inline-block;background:#0284c7;color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:bold">Ir al Dashboard →</a>
  <p style="margin-top:24px;font-size:11px;color:#64748b">Generado automáticamente por SGI Esinel. No responder a este correo.</p>
</div></div>`;

    MailApp.sendEmail({
      to:       email,
      cc:       'coordinacion.sgi@esinel.com.ar',
      subject:  `[SGI Esinel] Tenés ${info.items.length} hallazgo${info.items.length>1?'s':''} cerrado${info.items.length>1?'s':''} sin verificación de eficacia`,
      htmlBody: cuerpo,
      name:     REMITENTE_ALIAS
    });
  });

  Logger.log('enviarAvisosEficacia: ' + sinEficacia.length + ' pendientes, ' + Object.keys(porCorreo).length + ' destinatarios individuales');
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

// Reconoce nombres como "Lisandro Fissolo", "M. PICCIONE", "PICCIONE C.", "FISSOLO" o "Salud y Seguridad".
function normNombre(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[@.]/g, ' ').replace(/\s+/g, ' ').trim();
}
function buscarCorreo(nombre, mapa) {
  const n = normNombre(nombre);
  if (!n || n === '-') return null;
  const claves = Object.keys(mapa);
  const norm = {};
  claves.forEach(k => { norm[normNombre(k)] = mapa[k]; });
  if (norm[n]) return norm[n];
  const ks = Object.keys(norm);
  if (n.length >= 5) {
    const porTexto = [...new Set(ks.filter(k => k.includes(n) || n.includes(k)).map(k => norm[k]))];
    if (porTexto.length === 1) return porTexto[0];
  }
  // Apellido solo o inicial + apellido
  const toks = n.split(' ');
  const cands = ks.filter(k => {
    const kt = k.split(' ');
    const apellido = kt[kt.length - 1];
    if (!toks.includes(apellido)) return false;
    const resto = toks.filter(t => t !== apellido);
    return resto.every(t => t.length === 1 ? kt[0].startsWith(t) : kt.includes(t));
  });
  const correos = [...new Set(cands.map(k => norm[k]))];
  return correos.length === 1 ? correos[0] : null;   // si es ambiguo (ej. "PICCIONE" solo), no adivina
}
function separarNombres(str) {
  return String(str || '').split(/[,/|&;]|\s-\s|-(?=[A-ZÁÉÍÓÚÑ])/).map(s => s.trim()).filter(s => s && s !== '-');
}

function splitEntidades(str) { return separarNombres(str); }

function resolverCorreo(nombre) {
  return buscarCorreo(nombre, CORREOS);
}
