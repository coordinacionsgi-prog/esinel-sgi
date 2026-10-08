/**
 * I-AMB2 · Generación de Consumos — Dashboard automático
 * ---------------------------------------------------------------
 * Reemplaza TODO el código anterior del proyecto de Apps Script de la
 * planilla "I-AMB2 - Consumos EJ 35".
 *
 * Fuente de verdad: las hojas "HASTA dd-mm-aa". El período se toma del
 * NOMBRE de la hoja (no de las celdas DESDE/HASTA, que suelen quedar mal
 * al duplicar la hoja). Cada hoja tiene 3 bloques idénticos, uno por sede,
 * desplazados 19 columnas entre sí.
 *
 * Menú "Indicador AMB2" en la planilla:
 *   - Abrir dashboard
 *   - Actualizar ficha AMB2 (agrega a la ficha los trimestres que falten)
 *   - Corregir encabezados de períodos (fechas y Nº de trimestre)
 *   - Crear hoja del próximo trimestre
 *   - Revisar datos
 *   - Exportar HTML ahora
 *   - Activar actualización automática (cada hora: ficha + HTML exportado)
 *
 * Web App: Implementar > Nueva implementación > Aplicación web.
 *   /exec               -> dashboard
 *   /exec?formato=json  -> datos en JSON
 */

// ---------------------------------------------------------------- Config

var CFG = {
  // Se usa solo si el script no está dentro de la planilla (proyecto independiente).
  PLANILLA_ID: '1kjJgtY3Q9WT7-8gFpaQw09JvAtyNPUXxkkUyQj2d7_U',
  FICHA: 'AMB2',
  ARCHIVO_HTML: 'I-AMB2 Dashboard.html',   // copia exportada en la carpeta de la planilla
  PATRON_HOJA: /^\s*HASTA\s+(\d{1,2})-(\d{1,2})-(\d{2,4})\s*$/i,
  OBJETIVO_DEFAULT: 8,
  // Bloques de las hojas trimestrales (columna del rótulo "RESMAS:" etc. = C, V, AO)
  SEDES: [
    { id: 'REM', nombre: 'Remedios',   desp: 0,  despFicha: 0  },
    { id: 'STO', nombre: 'Santo Tomé', desp: 19, despFicha: 17 },
    { id: 'MOS', nombre: 'Mosconi',    desp: 38, despFicha: 33 }
  ],
  // Posiciones (fila, columna) del bloque Remedios; las demás sedes suman "desp".
  CELDA: {
    trimTxt:  [8, 3],   // C8  "2do"
    ejTxt:    [8, 4],   // D8  "TRIMESTRE DEL EJERCICIO 35"
    desde:    [9, 3],   // C9
    hasta:    [9, 6],   // F9
    diasHab:  [11, 4],  // D11
    nomina:   [12, 4],  // D12
    horas:    [13, 4],  // D13
    filaItem: [16, 17, 18], // resmas, plotter, energía
    rotulo:   3,  // C
    valor:    4,  // D
    nota:     9,  // I
    peso:     10, // J
    min:      12, // L
    max:      13, // M
    resultado:[20, 4]   // D20
  },
  ITEMS: [
    { id: 'resmas',  nombre: 'Resmas',          unidad: 'u'   },
    { id: 'plotter', nombre: 'Rollos plotter',  unidad: 'u'   },
    { id: 'energia', nombre: 'Energía eléctrica', unidad: 'kWh' }
  ],
  // Ficha AMB2 (bloque Remedios; las demás suman despFicha)
  FICHA_COL: { desde: 1, al: 2, hasta: 3, resultado: 4, decisiones: 10 },
  FICHA_FILA_INICIO: 23,
  ORDINAL: ['1er', '2do', '3er', '4to']
};

// ---------------------------------------------------------------- Menú / Web

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Indicador AMB2')
    .addItem('Abrir dashboard', 'abrirDashboard')
    .addSeparator()
    .addItem('Actualizar ficha AMB2', 'menuActualizarFicha')
    .addItem('Corregir encabezados de períodos', 'menuCorregirEncabezados')
    .addItem('Crear hoja del próximo trimestre', 'menuNuevoTrimestre')
    .addItem('Revisar datos', 'menuRevisar')
    .addItem('Exportar HTML ahora', 'menuExportarHtml')
    .addSeparator()
    .addItem('Activar actualización automática', 'activarAutomatico')
    .addToUi();
}

function doGet(e) {
  if (e && e.parameter && e.parameter.formato === 'json') {
    return ContentService.createTextOutput(JSON.stringify(obtenerDatos()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutput(htmlBase_())
    .setTitle('I-AMB2 · Consumos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function abrirDashboard() {
  var html = HtmlService.createHtmlOutput(htmlBase_())
    .setWidth(1200).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'I-AMB2 · Generación de Consumos');
}

// ---------------------------------------------------------------- Lectura

/** Devuelve todo lo que necesita el dashboard. Lo llama Dashboard.html. */
function obtenerDatos() {
  var ss = libro_();
  var trimestres = leerTrimestres_(ss);
  var objetivos = leerObjetivos_(ss);
  var alertas = revisar_(ss, trimestres);
  // google.script.run no admite objetos Date: las celdas crudas solo se usan para revisar.
  trimestres.forEach(function (t) {
    CFG.SEDES.forEach(function (s) { if (t.sedes[s.id]) delete t.sedes[s.id].celdas; });
  });

  trimestres.forEach(function (t) {
    var suma = 0, pesos = 0, simples = [];
    CFG.SEDES.forEach(function (s) {
      var d = t.sedes[s.id];
      if (!d || d.resultado === null) return;
      simples.push(d.resultado);
      if (d.nomina > 0) { suma += d.resultado * d.nomina; pesos += d.nomina; }
    });
    t.general = pesos > 0 ? redondear_(suma / pesos, 2)
      : (simples.length ? redondear_(promedio_(simples), 2) : null);
    t.objetivo = objetivos.general;
    t.cumple = t.general === null ? null : t.general >= t.objetivo;
  });

  var ejercicios = {};
  trimestres.forEach(function (t) {
    if (t.general === null) return;
    (ejercicios[t.ej] = ejercicios[t.ej] || []).push(t.general);
  });
  var resumenEj = Object.keys(ejercicios).map(Number).sort().map(function (ej) {
    return { ej: ej, promedio: redondear_(promedio_(ejercicios[ej]), 2), trimestres: ejercicios[ej].length };
  });

  return {
    planilla: ss.getName(),
    url: ss.getUrl(),
    generado: Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm'),
    sedes: CFG.SEDES.map(function (s) { return { id: s.id, nombre: s.nombre, objetivo: objetivos[s.id] }; }),
    items: CFG.ITEMS,
    objetivo: objetivos.general,
    trimestres: trimestres,
    ejercicios: resumenEj,
    alertas: alertas
  };
}

/** Lee todas las hojas "HASTA dd-mm-aa", ordenadas de la más vieja a la más nueva. */
function leerTrimestres_(ss) {
  var out = [];
  ss.getSheets().forEach(function (sh) {
    var per = periodoDesdeNombre_(sh.getName());
    if (!per) return;
    var v = sh.getRange(1, 1, 30, 56).getValues();
    var t = {
      hoja: sh.getName(),
      inicio: iso_(per.inicio),
      fin: iso_(per.fin),
      ej: per.ej,
      trim: per.trim,
      etiqueta: 'T' + per.trim + ' Ej.' + per.ej,
      rango: rangoTexto_(per.inicio, per.fin),
      sedes: {}
    };
    CFG.SEDES.forEach(function (s) { t.sedes[s.id] = leerBloque_(v, s.desp); });
    out.push(t);
  });
  out.sort(function (a, b) { return a.fin < b.fin ? -1 : 1; });
  return out;
}

function leerBloque_(v, desp) {
  var C = CFG.CELDA;
  var cel = function (f, c) { return v[f - 1][c - 1 + desp]; };
  var items = {};
  var vacio = true;
  CFG.ITEMS.forEach(function (it, i) {
    var f = C.filaItem[i];
    var val = num_(cel(f, C.valor));
    if (val !== null) vacio = false;
    items[it.id] = {
      rotulo: String(cel(f, C.rotulo) || ''),
      valor: val,
      nota: num_(cel(f, C.nota)),
      peso: num_(cel(f, C.peso)),
      min: num_(cel(f, C.min)),
      max: num_(cel(f, C.max))
    };
  });
  var res = num_(cel(C.resultado[0], C.resultado[1]));
  if (vacio) return null;   // sede sin consumos cargados ese trimestre
  var nomina = num_(cel(C.nomina[0], C.nomina[1]));
  var horas = num_(cel(C.horas[0], C.horas[1]));
  var kwh = items.energia.valor;
  return {
    nomina: nomina,
    diasHab: num_(cel(C.diasHab[0], C.diasHab[1])),
    horas: horas,
    items: items,
    resultado: res === null ? null : redondear_(res, 2),
    kwhPorPersona: (kwh !== null && nomina) ? redondear_(kwh / nomina, 1) : null,
    celdas: {
      desde: cel(C.desde[0], C.desde[1]),
      hasta: cel(C.hasta[0], C.hasta[1]),
      trimTxt: String(cel(C.trimTxt[0], C.trimTxt[1]) || ''),
      ejTxt: String(cel(C.ejTxt[0], C.ejTxt[1]) || '')
    }
  };
}

/** Objetivos: celda a la derecha de "Objetivo" en la ficha, por sede. */
function leerObjetivos_(ss) {
  var obj = { general: CFG.OBJETIVO_DEFAULT };
  CFG.SEDES.forEach(function (s) { obj[s.id] = CFG.OBJETIVO_DEFAULT; });
  var sh = ss.getSheetByName(CFG.FICHA);
  if (!sh) return obj;
  var fila = filaObjetivo_(sh);
  if (!fila) return obj;
  CFG.SEDES.forEach(function (s, i) {
    var val = num_(sh.getRange(fila, CFG.FICHA_COL.resultado + s.despFicha).getValue());
    if (val !== null) { obj[s.id] = val; if (i === 0) obj.general = val; }
  });
  return obj;
}

// ---------------------------------------------------------------- Revisión

function revisar_(ss, trimestres) {
  var al = [];
  var push = function (nivel, hoja, msg) { al.push({ nivel: nivel, hoja: hoja, msg: msg }); };

  ss.getSheets().forEach(function (sh) {
    var n = sh.getName();
    if (/^\s*HASTA/i.test(n) && !periodoDesdeNombre_(n)) {
      push('error', n, 'El nombre no tiene el formato "HASTA dd-mm-aa": la hoja se ignora.');
    }
  });

  var prev = {};
  trimestres.forEach(function (t) {
    var fin = new Date(t.fin + 'T00:00:00');
    var ini = new Date(t.inicio + 'T00:00:00');
    CFG.SEDES.forEach(function (s) {
      var d = t.sedes[s.id];
      if (!d) return;
      var donde = s.nombre;
      var desde = d.celdas.desde instanceof Date ? d.celdas.desde : null;
      var hasta = d.celdas.hasta instanceof Date ? d.celdas.hasta : null;
      if (desde && hasta && hasta < desde) {
        push('error', t.hoja, donde + ': HASTA (' + fmt_(hasta) + ') es anterior a DESDE (' + fmt_(desde) + ').');
      } else if ((desde && !mismoDia_(desde, ini)) || (hasta && !mismoDia_(hasta, fin))) {
        push('aviso', t.hoja, donde + ': las fechas DESDE/HASTA (' + fmt_(desde) + ' – ' + fmt_(hasta) +
          ') no coinciden con el trimestre de la hoja (' + t.rango + ').');
      }
      if (d.celdas.ejTxt && d.celdas.ejTxt.indexOf(String(t.ej)) < 0) {
        push('aviso', t.hoja, donde + ': dice "' + d.celdas.ejTxt + '" pero corresponde al Ejercicio ' + t.ej + '.');
      }
      if (d.resultado === null) push('error', t.hoja, donde + ': falta el RESULTADO.');
      CFG.ITEMS.forEach(function (it) {
        if (d.items[it.id].valor === null) push('aviso', t.hoja, donde + ': falta el dato de ' + it.nombre + '.');
      });
      var p = prev[s.id];
      if (p && p.nomina && d.nomina && Math.abs(d.nomina - p.nomina) / p.nomina > 0.5) {
        push('aviso', t.hoja, donde + ': la nómina pasó de ' + p.nomina + ' a ' + d.nomina +
          ' personas. Verificar (pondera el resultado general).');
      }
      prev[s.id] = d;
    });
  });

  var ficha = ss.getSheetByName(CFG.FICHA);
  if (ficha) {
    var enFicha = hojasEnFicha_(ficha);
    trimestres.forEach(function (t) {
      if (!enFicha[t.hoja]) push('aviso', CFG.FICHA, 'El trimestre ' + t.rango + ' (hoja "' + t.hoja +
        '") no figura en la ficha. Usar "Actualizar ficha AMB2".');
    });
  } else {
    push('error', CFG.FICHA, 'No existe la hoja de ficha "' + CFG.FICHA + '".');
  }
  return al;
}

function menuRevisar() {
  var ss = libro_();
  var al = revisar_(ss, leerTrimestres_(ss));
  var ui = SpreadsheetApp.getUi();
  if (!al.length) { ui.alert('Revisión de datos', 'Todo en orden.', ui.ButtonSet.OK); return; }
  ui.alert('Revisión de datos (' + al.length + ')', al.map(function (a) {
    return (a.nivel === 'error' ? '✖ ' : '⚠ ') + '[' + a.hoja.trim() + '] ' + a.msg;
  }).join('\n'), ui.ButtonSet.OK);
}

// ---------------------------------------------------------------- Ficha AMB2

/** Agrega a la ficha los trimestres nuevos que todavía no figuran. */
function actualizarFicha() {
  var ss = libro_();
  var sh = ss.getSheetByName(CFG.FICHA);
  if (!sh) throw new Error('No existe la hoja "' + CFG.FICHA + '".');
  var enFicha = hojasEnFicha_(sh);
  var ultimoFin = ultimoFinEnFicha_(sh, enFicha);
  var agregados = [];

  leerTrimestres_(ss).forEach(function (t) {
    if (enFicha[t.hoja]) return;
    if (ultimoFin && t.fin <= ultimoFin) return;   // trimestres viejos faltantes: solo se avisan
    var filaObj = filaObjetivo_(sh);
    if (!filaObj) throw new Error('No encuentro la fila "Objetivo" en la ficha.');
    var nueva = filaObj;                 // se inserta justo antes de "Objetivo"
    sh.insertRowBefore(filaObj);
    var ancho = sh.getLastColumn();
    sh.getRange(nueva - 1, 1, 1, ancho).copyTo(sh.getRange(nueva, 1, 1, ancho),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    var ini = new Date(t.inicio + 'T00:00:00'), fin = new Date(t.fin + 'T00:00:00');
    var C = CFG.FICHA_COL, R = CFG.CELDA.resultado;
    CFG.SEDES.forEach(function (s) {
      var hoja = "'" + t.hoja.replace(/'/g, "''") + "'!";
      var col = colLetra_(CFG.CELDA.valor + s.desp);
      var datos = hoja + col + CFG.CELDA.filaItem[0] + ':' + col + CFG.CELDA.filaItem[2];
      var ref = hoja + colLetra_(R[1] + s.desp) + R[0];
      var f = s.despFicha;
      sh.getRange(nueva, C.desde + f).setValue(ini);
      sh.getRange(nueva, C.al + f).setValue('al');
      sh.getRange(nueva, C.hasta + f).setValue(fin);
      // Vacío mientras no haya consumos cargados (si no, las notas dan 10 con celdas en blanco)
      sh.getRange(nueva, C.resultado + f).setFormula('=IF(COUNT(' + datos + ')=0,"",IFERROR(' + ref + ',""))');
      sh.getRange(nueva, C.decisiones + f).clearContent();
    });
    agregados.push(t.rango);
    ultimoFin = t.fin;
  });
  SpreadsheetApp.flush();
  return agregados;
}

function menuActualizarFicha() {
  var ag = actualizarFicha();
  SpreadsheetApp.getUi().alert('Ficha AMB2', ag.length
    ? 'Se agregaron ' + ag.length + ' trimestre(s):\n' + ag.join('\n')
    : 'La ficha ya estaba al día.', SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Mapa { nombreHoja: fila } de las hojas que la ficha ya referencia. */
function hojasEnFicha_(sh) {
  var mapa = {};
  var filaObj = filaObjetivo_(sh) || sh.getLastRow();
  var n = filaObj - CFG.FICHA_FILA_INICIO;
  if (n <= 0) return mapa;
  var cols = CFG.SEDES.map(function (s) { return CFG.FICHA_COL.resultado + s.despFicha; });
  cols.forEach(function (c) {
    sh.getRange(CFG.FICHA_FILA_INICIO, c, n, 1).getFormulas().forEach(function (r, i) {
      var m = String(r[0]).match(/'((?:[^']|'')+)'!/) || String(r[0]).match(/=\+?([^'!(]+)!/);
      if (m) mapa[m[1].replace(/''/g, "'")] = CFG.FICHA_FILA_INICIO + i;
    });
  });
  return mapa;
}

function ultimoFinEnFicha_(sh, enFicha) {
  var ult = null;
  Object.keys(enFicha).forEach(function (h) {
    var p = periodoDesdeNombre_(h);
    if (p && (!ult || iso_(p.fin) > ult)) ult = iso_(p.fin);
  });
  return ult;
}

function filaObjetivo_(sh) {
  var col = CFG.FICHA_COL.hasta;
  var desde = CFG.FICHA_FILA_INICIO;
  var n = Math.max(sh.getLastRow() - desde + 1, 1);
  var v = sh.getRange(desde, col, n, 1).getValues();
  for (var i = 0; i < v.length; i++) {
    if (String(v[i][0]).trim().toLowerCase() === 'objetivo') return desde + i;
  }
  return null;
}

// ---------------------------------------------------------------- Encabezados

/** Pone en las hojas trimestrales (o solo en "soloHoja") las fechas y el Nº de trimestre según su nombre. */
function corregirEncabezados(soloHoja) {
  var ss = libro_();
  var C = CFG.CELDA, cambios = 0;
  ss.getSheets().forEach(function (sh) {
    if (typeof soloHoja === 'string' && sh.getName() !== soloHoja) return;
    var p = periodoDesdeNombre_(sh.getName());
    if (!p) return;
    var v = sh.getRange(1, 1, 30, 56).getValues();
    CFG.SEDES.forEach(function (s) {
      if (!leerBloque_(v, s.desp)) return;
      var set = function (pos, val) {
        var actual = v[pos[0] - 1][pos[1] - 1 + s.desp];
        var igual = (actual instanceof Date && val instanceof Date) ? mismoDia_(actual, val)
          : String(actual).trim().toUpperCase() === String(val).toUpperCase();
        if (!igual) { sh.getRange(pos[0], pos[1] + s.desp).setValue(val); cambios++; }
      };
      set(C.trimTxt, CFG.ORDINAL[p.trim - 1]);
      set(C.ejTxt, 'TRIMESTRE DEL EJERCICIO ' + p.ej);
      set(C.desde, p.inicio);
      set(C.hasta, p.fin);
    });
  });
  return cambios;
}

function menuCorregirEncabezados() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.alert('Corregir encabezados',
    'Se van a reescribir, en cada hoja "HASTA ...", las celdas de trimestre, ejercicio, DESDE y HASTA ' +
    'de las 3 sedes según el nombre de la hoja. Los consumos NO se tocan. ¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;
  ui.alert('Corregir encabezados', 'Celdas corregidas: ' + corregirEncabezados(), ui.ButtonSet.OK);
}

// ---------------------------------------------------------------- Nuevo trimestre

/** Duplica la última hoja trimestral como la del trimestre siguiente, con consumos en blanco. */
function nuevoTrimestre() {
  var ss = libro_();
  var ts = leerTrimestres_(ss);
  if (!ts.length) throw new Error('No hay hojas "HASTA ..." para usar de modelo.');
  var ult = ts[ts.length - 1];
  var finUlt = new Date(ult.fin + 'T00:00:00');
  var finNuevo = new Date(finUlt.getFullYear(), finUlt.getMonth() + 4, 0);   // fin del trimestre siguiente
  var nombre = 'HASTA ' + Utilities.formatDate(finNuevo, ss.getSpreadsheetTimeZone(), 'dd-MM-yy');
  if (ss.getSheetByName(nombre)) throw new Error('Ya existe la hoja "' + nombre + '".');

  var modelo = ss.getSheetByName(ult.hoja);
  var sh = modelo.copyTo(ss).setName(nombre);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(modelo.getIndex());   // queda antes que la anterior, como el resto
  var C = CFG.CELDA;
  // Fechas y trimestre primero (con los consumos del modelo, para que detecte qué sedes usar)
  corregirEncabezados(nombre);
  CFG.SEDES.forEach(function (s) {
    C.filaItem.forEach(function (f) { sh.getRange(f, C.valor + s.desp).clearContent(); });
  });
  actualizarFicha();
  return nombre;
}

function menuNuevoTrimestre() {
  var ui = SpreadsheetApp.getUi();
  var n = nuevoTrimestre();
  ui.alert('Nuevo trimestre', 'Se creó la hoja "' + n + '" y se agregó a la ficha.\n' +
    'Cargá resmas, rollos y energía de cada sede (y revisá nómina y días hábiles).\n' +
    'Ojo: si la energía de Remedios viene por IMPORTRANGE, actualizá el rango de filas.', ui.ButtonSet.OK);
}

// ---------------------------------------------------------------- HTML

/** HTML del dashboard: embebido (versión de un solo archivo) o archivo "Dashboard". */
function htmlBase_() {
  return (typeof DASHBOARD_HTML === 'string') ? DASHBOARD_HTML
    : HtmlService.createHtmlOutputFromFile('Dashboard').getContent();
}

/** Guarda en la carpeta de la planilla una copia del dashboard con los datos actuales. */
function exportarHtml() {
  var ss = libro_();
  var datos = JSON.stringify(obtenerDatos()).replace(/</g, '\\u003c');
  var html = htmlBase_().replace('<script src="https://cdn.jsdelivr.net',
    '<script>window.__AMB2_DATA__ = ' + datos + ';</script>\n<script src="https://cdn.jsdelivr.net');
  var carpeta = DriveApp.getFileById(ss.getId()).getParents().next();
  var it = carpeta.getFilesByName(CFG.ARCHIVO_HTML);
  var archivo = it.hasNext() ? it.next() : null;
  if (archivo) archivo.setContent(html);
  else archivo = carpeta.createFile(CFG.ARCHIVO_HTML, html, MimeType.HTML);
  return archivo.getUrl();
}

function menuExportarHtml() {
  var url = exportarHtml();
  SpreadsheetApp.getUi().alert('HTML exportado', 'Se actualizó "' + CFG.ARCHIVO_HTML + '":\n' + url,
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ---------------------------------------------------------------- Automatización

function activarAutomatico() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'tareaAutomatica') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('tareaAutomatica').timeBased().everyHours(1).create();
  tareaAutomatica();
  SpreadsheetApp.getUi().alert('Listo: cada hora se pone al día la ficha AMB2 y se exporta "' +
    CFG.ARCHIVO_HTML + '" con los datos actuales.');
}

function tareaAutomatica() {
  actualizarFicha();
  exportarHtml();
}

// ---------------------------------------------------------------- Utilidades

function libro_() {
  return SpreadsheetApp.getActive() || SpreadsheetApp.openById(CFG.PLANILLA_ID);
}

/** "HASTA 30-09-26" -> período del trimestre que termina ese día. Ejercicio: abril a marzo. */
function periodoDesdeNombre_(nombre) {
  var m = String(nombre).match(CFG.PATRON_HOJA);
  if (!m) return null;
  var d = +m[1], mes = +m[2], a = +m[3];
  if (a < 100) a += 2000;
  if (mes < 1 || mes > 12) return null;
  var fin = new Date(a, mes, 0);   // último día del mes (corrige 30-03, 30-12, etc.)
  if (Math.abs(d - fin.getDate()) > 3) return null;
  var inicio = new Date(a, mes - 3, 1);
  var ej = mes >= 4 ? a - 1991 : a - 1992;
  var trim = mes >= 4 ? Math.ceil((mes - 3) / 3) : 4;
  return { inicio: inicio, fin: fin, ej: ej, trim: trim };
}

function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return isFinite(v) ? v : null;
  var t = String(v).trim();
  if (!t) return null;
  if (t.indexOf(',') >= 0) t = t.replace(/\./g, '').replace(',', '.');   // 1.234,5
  var n = Number(t);
  return isFinite(n) ? n : null;
}
function redondear_(n, d) { var f = Math.pow(10, d); return Math.round(n * f) / f; }
function promedio_(a) { return a.reduce(function (s, x) { return s + x; }, 0) / a.length; }
function iso_(d) { return d.getFullYear() + '-' + pad_(d.getMonth() + 1) + '-' + pad_(d.getDate()); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function fmt_(d) { return d ? pad_(d.getDate()) + '/' + pad_(d.getMonth() + 1) + '/' + d.getFullYear() : '—'; }
function mismoDia_(a, b) { return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
function colLetra_(c) { var s = ''; while (c > 0) { var r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = (c - r - 1) / 26; } return s; }
function rangoTexto_(ini, fin) {
  var M = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  return M[ini.getMonth()] + '–' + M[fin.getMonth()] + ' ' + fin.getFullYear();
}

// ---------------------------------------------------------------- Página del dashboard
// (generado a partir de Dashboard.html)
var DASHBOARD_HTML = String.raw`<!DOCTYPE html>
<html lang="es">
<head>
<base target="_top">
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>SGI · Indicador I-AMB2</title>
<script src="https://cdn.tailwindcss.com/3.4.16"></script>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script>
  tailwind.config = { theme: { extend: { fontFamily: { sans: ['Plus Jakarta Sans', 'system-ui', 'sans-serif'] } } } };
</script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  body { background: #f1f5f9; }
  .card { background: #fff; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(15,23,42,.05); border-radius: 1rem; }
  .pill { display: inline-flex; align-items: center; gap: .25rem; padding: .125rem .5rem; border-radius: 9999px; font-size: 10px; font-weight: 800; border: 1px solid transparent; white-space: nowrap; }
  .pill.ok { background: #d1fae5; color: #065f46; border-color: #a7f3d0; }
  .pill.bad { background: #ffe4e6; color: #9f1239; border-color: #fecdd3; }
  .pill.na { background: #f1f5f9; color: #475569; border-color: #e2e8f0; }
  .sede-btn { padding: .375rem .75rem; border-radius: .75rem; font-size: 12px; font-weight: 700; background: #f1f5f9; color: #334155; transition: background .15s; }
  .sede-btn:hover { background: #e2e8f0; }
  .sede-btn.on { background: #0f172a; color: #fff; }
  .chart-box { position: relative; height: 260px; }
  .spin { animation: spin 1s linear infinite; }
  @keyframes spin { to { transform: rotate(360deg); } }
  table.det th { font-size: 10px; text-transform: uppercase; letter-spacing: .05em; }
  table.det td, table.det th { padding: .7rem .75rem; text-align: center; white-space: nowrap; }
  table.det td:first-child, table.det th:first-child { text-align: left; padding-left: 1.25rem; }
  .dim { opacity: .35; }
  @media print { .no-print { display: none !important; } body { background: #fff; } }
</style>
</head>
<body class="text-slate-800 font-sans antialiased min-h-screen">

<header class="bg-white border-b border-slate-200 sticky top-0 z-30">
  <div class="max-w-7xl mx-auto px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
    <div class="flex items-center gap-3 min-w-0">
      <div class="w-10 h-10 rounded-xl bg-slate-900 flex items-center justify-center text-lg shrink-0">🍃</div>
      <div class="min-w-0">
        <div class="flex flex-wrap items-center gap-2">
          <h1 class="text-base font-extrabold text-slate-900 leading-snug">SGI · Indicador I-AMB2 <span class="text-slate-400 font-semibold">Generación de consumos</span></h1>
          <span class="pill ok">ISO 14001</span>
        </div>
        <p class="text-xs text-slate-500 truncate" id="sub">Leyendo la planilla…</p>
      </div>
    </div>
    <div class="flex flex-wrap items-center gap-2 no-print">
      <div class="flex items-center bg-slate-100 p-1 rounded-xl border border-slate-200 text-xs">
        <label for="fPer" class="px-2 font-bold text-slate-600">📅 Trimestre:</label>
        <select id="fPer" class="bg-white text-slate-900 font-bold py-1.5 px-2 rounded-lg border border-slate-300 text-xs cursor-pointer focus:outline-none focus:ring-2 focus:ring-blue-500"></select>
      </div>
      <button id="btnRefresh" type="button" class="px-3.5 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-extrabold flex items-center gap-1.5 shadow-sm disabled:opacity-60">
        <span id="icoRefresh">⟳</span> Actualizar
      </button>
      <button id="btnHelp" type="button" class="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold border border-slate-300">? Ayuda</button>
    </div>
  </div>
</header>

<div class="max-w-7xl mx-auto px-4 sm:px-6 pt-4 no-print">
  <div class="card p-2.5 flex flex-wrap items-center justify-between gap-2">
    <div class="flex flex-wrap items-center gap-2 ml-1">
      <span class="text-xs font-bold text-slate-600">Vista por sede:</span>
      <span id="lblNomina" class="text-[11px] font-semibold text-slate-500 bg-slate-100 px-2 py-0.5 rounded">Nómina total: —</span>
    </div>
    <div class="flex flex-wrap gap-1.5" id="sedeBtns"></div>
  </div>
</div>

<main class="max-w-7xl mx-auto px-4 sm:px-6 py-5 space-y-5" id="app">
  <div class="card p-10 text-center text-sm text-slate-500">Leyendo hojas trimestrales…</div>
</main>

<div id="modalHelp" class="hidden fixed inset-0 z-50 bg-slate-900/60 flex items-center justify-center p-4">
  <div class="bg-white rounded-3xl max-w-xl w-full p-6 shadow-2xl border border-slate-200 space-y-4">
    <div class="flex justify-between items-start border-b border-slate-100 pb-3">
      <div>
        <h3 class="text-base font-extrabold text-slate-900">Cómo funciona este dashboard</h3>
        <p class="text-xs text-slate-500">De dónde salen los datos y cómo se actualiza</p>
      </div>
      <button type="button" data-close class="w-8 h-8 rounded-lg hover:bg-slate-100 text-slate-400 hover:text-slate-700">✕</button>
    </div>
    <div class="space-y-3 text-xs text-slate-600 leading-relaxed">
      <div class="p-3 bg-blue-50/70 rounded-xl border border-blue-200">
        <b class="text-blue-900 block mb-1">1. Lee la planilla en el momento</b>
        Cada vez que abrís esta página (o tocás <b>Actualizar</b>) se leen todas las hojas <b>HASTA dd-mm-aa</b>. Si cambiás un dato en la planilla, alcanza con actualizar.
      </div>
      <div class="p-3 bg-emerald-50/70 rounded-xl border border-emerald-200">
        <b class="text-emerald-900 block mb-1">2. Cálculo</b>
        Cada sede tiene su resultado (celdas D20, W20 y AP20). El <b>global</b> es el promedio de las 3 sedes ponderado por su nómina. Meta: el valor de "Objetivo" en la ficha AMB2.
      </div>
      <div class="p-3 bg-slate-50 rounded-xl border border-slate-200">
        <b class="text-slate-900 block mb-1">3. Trimestre nuevo</b>
        Duplicá la última hoja <b>HASTA …</b> (o usá el menú <b>Indicador AMB2 → Crear hoja del próximo trimestre</b>) y cargá los consumos. El período se toma del nombre de la hoja.
      </div>
    </div>
    <div class="pt-2 flex justify-end gap-2 border-t border-slate-100">
      <button type="button" id="btnHelpRefresh" class="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold">Actualizar ahora</button>
      <button type="button" data-close class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold">Cerrar</button>
    </div>
  </div>
</div>

<script>
var DATA = null, charts = {}, filtroSede = 'ALL';
var SEDE_UI = {
  REM: { color: '#2563eb', borde: 'border-l-blue-600', txt: 'text-blue-700', val: 'text-blue-900', sub: 'CABA' },
  STO: { color: '#0d9488', borde: 'border-l-teal-600', txt: 'text-teal-700', val: 'text-teal-900', sub: 'Santa Fe' },
  MOS: { color: '#7c3aed', borde: 'border-l-purple-600', txt: 'text-purple-700', val: 'text-purple-900', sub: 'Mosconi' }
};
var ITEM_UI = {
  energia: { ico: '⚡', color: 'amber', titulo: 'Consumo de energía eléctrica', unidad: 'kWh' },
  resmas:  { ico: '📄', color: 'blue', titulo: 'Consumo de resmas de papel', unidad: 'resmas' },
  plotter: { ico: '🖨️', color: 'purple', titulo: 'Consumo de rollos de plotter', unidad: 'rollos' }
};
var ORDEN_ITEMS = ['energia', 'plotter', 'resmas'];
var MAX_TRIM_GRAF = 8;

function n(v, d) {
  if (v === null || v === undefined || isNaN(v)) return '—';
  return Number(v).toLocaleString('es-AR', { minimumFractionDigits: d || 0, maximumFractionDigits: d === undefined ? 2 : d });
}
function pct(v) { return v === null || v === undefined ? '—' : Math.round(v * 100) + '%'; }
function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function el(id) { return document.getElementById(id); }

function cargar() {
  el('btnRefresh').disabled = true;
  el('icoRefresh').className = 'spin';
  if (window.google && google.script && google.script.run) {
    google.script.run.withSuccessHandler(recibir).withFailureHandler(function (e) {
      el('app').innerHTML = '<div class="card p-10 text-center text-sm text-rose-700">No se pudieron leer los datos: ' + esc(e.message || e) + '</div>';
      fin();
    }).obtenerDatos();
  } else if (window.__AMB2_DATA__) {
    recibir(window.__AMB2_DATA__);
  }
}
function fin() { el('btnRefresh').disabled = false; el('icoRefresh').className = ''; }

function recibir(d) {
  DATA = d;
  fin();
  el('sub').innerHTML = '<a class="hover:underline" href="' + esc(d.url) + '" target="_blank" rel="noopener">' + esc(d.planilla) +
    '</a> · actualizado ' + esc(d.generado) + ' · Remedios, Santo Tomé y Mosconi';
  var sel = el('fPer'), prev = sel.value;
  var conDato = d.trimestres.filter(function (t) { return t.general !== null; });
  sel.innerHTML = conDato.slice().reverse().map(function (t, i) {
    return '<option value="' + esc(t.hoja) + '">' + esc(t.rango) + ' · ' + esc(t.etiqueta) + (i === 0 ? ' (último)' : '') + '</option>';
  }).join('');
  if (prev && sel.querySelector('option[value="' + prev.replace(/"/g, '') + '"]')) sel.value = prev;
  el('sedeBtns').innerHTML = '<button type="button" class="sede-btn" data-s="ALL">🏢 Global (todas)</button>' +
    d.sedes.map(function (s) {
      return '<button type="button" class="sede-btn" data-s="' + s.id + '"><span style="color:' + SEDE_UI[s.id].color + '">●</span> ' + esc(s.nombre) + '</button>';
    }).join('');
  el('sedeBtns').querySelectorAll('button').forEach(function (b) {
    b.onclick = function () { filtroSede = b.dataset.s; render(); };
  });
  render();
}

function render() {
  if (!DATA) return;
  el('sedeBtns').querySelectorAll('button').forEach(function (b) { b.classList.toggle('on', b.dataset.s === filtroSede); });
  var conDato = DATA.trimestres.filter(function (t) { return t.general !== null; });
  if (!conDato.length) {
    el('app').innerHTML = '<div class="card p-10 text-center text-sm text-slate-500">No hay hojas "HASTA dd-mm-aa" con resultados.</div>';
    return;
  }
  var hoja = el('fPer').value;
  var idx = conDato.map(function (t) { return t.hoja; }).indexOf(hoja);
  if (idx < 0) idx = conDato.length - 1;
  var cur = conDato[idx], ant = idx > 0 ? conDato[idx - 1] : null;
  var serie = conDato.slice(Math.max(0, idx - MAX_TRIM_GRAF + 1), idx + 1);
  var resEj = DATA.ejercicios.filter(function (e) { return e.ej === cur.ej; })[0];

  var nomTot = 0;
  DATA.sedes.forEach(function (s) { var d = cur.sedes[s.id]; if (d && d.nomina) nomTot += d.nomina; });
  el('lblNomina').textContent = 'Nómina total: ' + n(nomTot, 0) + ' personas · ' + cur.rango;

  var h = '';
  // --- Tarjetas
  h += '<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">';
  h += tarjetaGlobal(cur, ant, nomTot, resEj);
  DATA.sedes.forEach(function (s) { h += tarjetaSede(s, cur, ant, nomTot); });
  h += '</div>';

  // --- Tabla de desglose
  h += '<div class="card overflow-hidden">' +
    '<div class="px-5 py-4 border-b border-slate-200 bg-slate-50/70 flex flex-wrap justify-between items-center gap-2">' +
      '<div><h2 class="text-sm font-extrabold text-slate-900">📊 Desglose del trimestre · ' + esc(cur.rango) + '</h2>' +
      '<p class="text-xs text-slate-500">Consumo real de cada sede y su nota (2 a 10) según la escala de la hoja ' + esc(cur.hoja.trim()) + '</p></div>' +
      '<div class="flex flex-wrap items-center gap-1.5 text-[11px] font-bold">' + pesosHTML(cur) + '</div>' +
    '</div><div class="overflow-x-auto">' + tablaHTML(cur, nomTot) + '</div></div>';

  // --- Gráficos
  h += '<div class="grid grid-cols-1 lg:grid-cols-2 gap-5">';
  h += graficoCard('cRes', '📈', 'Evolución de la nota vs. meta', 'Global ponderado por nómina y nota de cada sede · últimos ' + serie.length + ' trimestres',
    '<span class="pill bad">Meta: ' + n(cur.objetivo, 2) + '</span>');
  ORDEN_ITEMS.forEach(function (k) {
    var u = ITEM_UI[k], it = DATA.items.filter(function (x) { return x.id === k; })[0];
    var peso = pesoItem(cur, k);
    h += graficoCard('c_' + k, u.ico, u.titulo + ' (' + esc(it.unidad) + ')', 'Por sede y por trimestre',
      peso !== null ? '<span class="pill na">Peso: ' + pct(peso) + '</span>' : '');
  });
  h += '</div>';

  // --- Revisión de datos
  h += revisionHTML(conDato);
  h += '<p class="text-[11px] text-slate-400 text-center pb-4">El período sale del nombre de cada hoja (ejercicio de abril a marzo). El global pondera cada sede por su nómina.</p>';

  el('app').innerHTML = h;
  dibujar(serie, cur);
}

function pill(v, obj) {
  if (v === null || v === undefined) return '<span class="pill na">Sin dato</span>';
  return v >= obj ? '<span class="pill ok">Cumple</span>' : '<span class="pill bad">Desvío</span>';
}
function delta(a, b) {
  if (a === null || a === undefined || b === null || b === undefined) return '<span class="text-slate-400">sin trimestre anterior</span>';
  var d = a - b;
  if (Math.abs(d) < 0.005) return '<span class="text-slate-500 font-bold">= trimestre anterior</span>';
  return '<span class="font-bold ' + (d > 0 ? 'text-emerald-600' : 'text-rose-600') + '">' + (d > 0 ? '▲ +' : '▼ ') + n(d, 2) + ' vs. anterior</span>';
}
function pesoItem(t, k) {
  for (var i = 0; i < DATA.sedes.length; i++) {
    var d = t.sedes[DATA.sedes[i].id];
    if (d && d.items[k] && d.items[k].peso !== null) return d.items[k].peso;
  }
  return null;
}

function tarjetaGlobal(cur, ant, nomTot, resEj) {
  var ok = cur.general >= cur.objetivo;
  var dif = cur.general - cur.objetivo;
  var dim = filtroSede !== 'ALL' ? ' dim' : '';
  return '<div class="card p-5 border-2 ' + (ok ? 'border-emerald-300' : 'border-rose-300') + dim + '">' +
    '<div class="flex justify-between items-start mb-2"><span class="text-[11px] font-extrabold uppercase tracking-wider text-slate-600">🏢 Global empresa</span>' +
      pill(cur.general, cur.objetivo) + '</div>' +
    '<div class="flex items-baseline gap-2 mb-1"><span class="text-3xl font-extrabold text-slate-900">' + n(cur.general, 2) + '</span>' +
      '<span class="text-xs font-bold text-slate-400">/ 10 pts</span></div>' +
    '<p class="text-xs text-slate-600 font-semibold">Ponderado por nómina: ' + n(nomTot, 0) + ' personas</p>' +
    '<p class="text-[11px] mt-1">' + delta(cur.general, ant && ant.general) + '</p>' +
    '<div class="mt-3 pt-2.5 border-t border-slate-100 text-[11px] flex justify-between items-center text-slate-500 font-semibold gap-2">' +
      '<span>Meta: <b>≥ ' + n(cur.objetivo, 2) + '</b></span>' +
      '<span class="font-bold ' + (dif >= 0 ? 'text-emerald-600' : 'text-rose-600') + '">' + (dif >= 0 ? '+' : '') + n(dif, 2) + ' vs. meta</span></div>' +
    (resEj ? '<div class="mt-1.5 text-[11px] text-slate-500 font-semibold">Promedio Ej. ' + resEj.ej + ': <b class="text-slate-800">' + n(resEj.promedio, 2) +
      '</b> (' + resEj.trimestres + ' de 4 trim.)</div>' : '') +
  '</div>';
}

function tarjetaSede(s, cur, ant, nomTot) {
  var u = SEDE_UI[s.id], d = cur.sedes[s.id];
  var dim = (filtroSede !== 'ALL' && filtroSede !== s.id) ? ' dim' : '';
  var cab = '<div class="flex justify-between items-start mb-2"><div><span class="text-[11px] font-extrabold uppercase tracking-wider ' + u.txt + ' block">' +
    esc(s.nombre) + '</span><span class="text-[10px] text-slate-400 font-bold">' +
    (d ? 'Nómina: ' + n(d.nomina, 0) + ' pers. · ' + (nomTot ? n(d.nomina / nomTot * 100, 1) : '—') + '% del global' : u.sub) + '</span></div>' +
    pill(d ? d.resultado : null, s.objetivo) + '</div>';
  if (!d) {
    return '<div class="card p-5 border-l-4 ' + u.borde + dim + '">' + cab + '<p class="text-sm text-slate-400 mt-4">Sin consumos cargados este trimestre.</p></div>';
  }
  var pa = ant && ant.sedes[s.id];
  var it = d.items;
  return '<div class="card p-5 border-l-4 ' + u.borde + dim + '">' + cab +
    '<div class="flex items-baseline gap-2 mb-1"><span class="text-3xl font-extrabold ' + u.val + '">' + n(d.resultado, 2) + '</span>' +
      '<span class="text-xs font-bold text-slate-400">/ 10 pts</span></div>' +
    '<p class="text-xs text-slate-600 font-medium">⚡ ' + n(it.energia.nota, 0) + ' · 🖨️ ' + n(it.plotter.nota, 0) + ' · 📄 ' + n(it.resmas.nota, 0) + ' <span class="text-slate-400">(notas)</span></p>' +
    '<p class="text-[11px] mt-1">' + delta(d.resultado, pa ? pa.resultado : null) + '</p>' +
    '<div class="mt-3 pt-2.5 border-t border-slate-100 text-[11px] flex justify-between items-center text-slate-500 font-semibold gap-2">' +
      '<span>Energía: <b class="text-slate-700">' + n(it.energia.valor, 0) + '</b> kWh</span>' +
      '<span class="font-bold ' + u.txt + '">' + n(d.kwhPorPersona, 0) + ' kWh/pers.</span></div>' +
  '</div>';
}

function pesosHTML(cur) {
  var cls = { amber: 'bg-amber-50 text-amber-800 border-amber-200', purple: 'bg-purple-50 text-purple-800 border-purple-200', blue: 'bg-blue-50 text-blue-800 border-blue-200' };
  return ORDEN_ITEMS.map(function (k) {
    var p = pesoItem(cur, k), u = ITEM_UI[k];
    var nombre = DATA.items.filter(function (x) { return x.id === k; })[0].nombre;
    return '<span class="px-2 py-0.5 rounded border ' + cls[u.color] + '">' + esc(nombre) + ': ' + pct(p) + '</span>';
  }).join('');
}

function notaCls(v) {
  if (v === null || v === undefined) return 'text-slate-400';
  return v >= 8 ? 'text-emerald-700' : v >= 6 ? 'text-amber-700' : 'text-rose-700';
}

function tablaHTML(cur, nomTot) {
  var h = '<table class="det w-full text-xs border-collapse"><thead><tr class="bg-slate-100/80 text-slate-700 font-bold border-b border-slate-200">' +
    '<th>Sede</th><th>Nómina</th>' +
    '<th class="bg-amber-50/60 text-amber-900">Energía</th><th class="bg-amber-50/60 text-amber-900">Nota</th>' +
    '<th class="bg-purple-50/60 text-purple-900">Plotter</th><th class="bg-purple-50/60 text-purple-900">Nota</th>' +
    '<th class="bg-blue-50/60 text-blue-900">Resmas</th><th class="bg-blue-50/60 text-blue-900">Nota</th>' +
    '<th class="text-slate-900">Nota sede</th><th>Aporte al global</th></tr></thead><tbody class="divide-y divide-slate-200 font-medium text-slate-700">';
  DATA.sedes.forEach(function (s) {
    if (filtroSede !== 'ALL' && filtroSede !== s.id) return;
    var u = SEDE_UI[s.id], d = cur.sedes[s.id];
    if (!d) {
      h += '<tr><td class="font-bold ' + u.txt + '"><span style="color:' + u.color + '">●</span> ' + esc(s.nombre) + '</td><td colspan="9" class="text-slate-400">Sin consumos cargados</td></tr>';
      return;
    }
    var it = d.items;
    var aporte = nomTot && d.resultado !== null ? d.resultado * d.nomina / nomTot : null;
    h += '<tr class="hover:bg-slate-50">' +
      '<td class="font-bold ' + u.txt + '"><span style="color:' + u.color + '">●</span> ' + esc(s.nombre) + '</td>' +
      '<td class="font-bold bg-slate-50/50">' + n(d.nomina, 0) + ' pers.</td>' +
      '<td class="font-bold text-slate-800 bg-amber-50/30">' + n(it.energia.valor, 0) + ' kWh</td>' +
      '<td class="font-extrabold bg-amber-50/30 ' + notaCls(it.energia.nota) + '">' + n(it.energia.nota, 0) + '</td>' +
      '<td class="font-bold text-slate-800 bg-purple-50/30">' + n(it.plotter.valor) + ' rollos</td>' +
      '<td class="font-extrabold bg-purple-50/30 ' + notaCls(it.plotter.nota) + '">' + n(it.plotter.nota, 0) + '</td>' +
      '<td class="font-bold text-slate-800 bg-blue-50/30">' + n(it.resmas.valor) + ' resmas</td>' +
      '<td class="font-extrabold bg-blue-50/30 ' + notaCls(it.resmas.nota) + '">' + n(it.resmas.nota, 0) + '</td>' +
      '<td class="font-extrabold text-sm ' + (d.resultado !== null && d.resultado < s.objetivo ? 'text-rose-700' : 'text-emerald-700') + '">' + n(d.resultado, 2) + '</td>' +
      '<td class="font-mono text-slate-600">' + n(aporte, 3) + '</td></tr>';
  });
  if (filtroSede === 'ALL') {
    h += '<tr class="bg-slate-50 font-extrabold text-slate-900"><td>🏢 Global</td><td>' + n(nomTot, 0) + ' pers.</td><td colspan="6"></td>' +
      '<td class="text-sm ' + (cur.cumple === false ? 'text-rose-700' : 'text-emerald-700') + '">' + n(cur.general, 2) + '</td><td class="font-mono">' + n(cur.general, 3) + '</td></tr>';
  }
  return h + '</tbody></table>';
}

function graficoCard(id, ico, titulo, sub, extra) {
  return '<div class="card p-5"><div class="mb-3 flex justify-between items-start gap-2"><div>' +
    '<h3 class="text-sm font-extrabold text-slate-900">' + ico + ' ' + titulo + '</h3>' +
    '<p class="text-xs text-slate-500">' + sub + '</p></div>' + (extra || '') + '</div>' +
    '<div class="chart-box"><canvas id="' + id + '"></canvas></div></div>';
}

function revisionHTML(conDato) {
  var als = DATA.alertas.slice().sort(function (a, b) { return (a.nivel === 'error' ? 0 : 1) - (b.nivel === 'error' ? 0 : 1); });
  var h = '<div class="card p-5"><h3 class="text-sm font-extrabold text-slate-900">🔎 Revisión de datos</h3>';
  if (!als.length) return h + '<p class="text-xs text-emerald-700 font-semibold mt-1">✓ Sin inconsistencias en la planilla.</p></div>';
  h += '<p class="text-xs text-slate-500 mb-3">' + als.length + ' punto(s) para revisar en la planilla</p><ul class="space-y-2">';
  als.forEach(function (a) {
    var err = a.nivel === 'error';
    h += '<li class="flex gap-2 text-xs p-2.5 rounded-xl border ' + (err ? 'bg-rose-50 border-rose-200 text-rose-900' : 'bg-amber-50 border-amber-200 text-amber-900') + '">' +
      '<span class="font-extrabold">' + (err ? '✖' : '!') + '</span><span>' + esc(a.msg) +
      ' <span class="text-slate-500 font-semibold">· ' + esc(a.hoja.trim()) + '</span></span></li>';
  });
  return h + '</ul></div>';
}

function opts(min, max, step) {
  return {
    responsive: true, maintainAspectRatio: false,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { position: 'bottom', labels: { usePointStyle: true, boxWidth: 8, font: { family: 'Plus Jakarta Sans', size: 11, weight: 600 }, color: '#475569' } },
      tooltip: { backgroundColor: '#0f172a', padding: 10, titleFont: { family: 'Plus Jakarta Sans' }, bodyFont: { family: 'Plus Jakarta Sans' },
        callbacks: { label: function (c) { return ' ' + c.dataset.label + ': ' + (c.raw === null ? '—' : n(c.raw, max === 10.5 ? 2 : undefined)); } } }
    },
    scales: {
      x: { grid: { display: false }, ticks: { color: '#64748b', font: { family: 'Plus Jakarta Sans', size: 11 } } },
      y: { min: min, max: max, beginAtZero: min === 0, grid: { color: '#f1f5f9' }, border: { display: false },
        ticks: { color: '#64748b', stepSize: step, font: { family: 'Plus Jakarta Sans', size: 11 }, callback: function (v) { return n(v, 0); } } }
    }
  };
}

function dibujar(serie, cur) {
  Object.keys(charts).forEach(function (k) { charts[k].destroy(); });
  charts = {};
  var labels = serie.map(function (t) { return t.rango; });
  var sedes = DATA.sedes.filter(function (s) { return filtroSede === 'ALL' || filtroSede === s.id; });

  var ds = [];
  if (filtroSede === 'ALL') {
    ds.push({ label: 'Global empresa', data: serie.map(function (t) { return t.general; }), borderColor: '#059669', backgroundColor: '#059669',
      borderWidth: 3.5, tension: 0.25, pointRadius: 5, spanGaps: true });
  }
  sedes.forEach(function (s) {
    var c = SEDE_UI[s.id].color;
    ds.push({ label: s.nombre, data: serie.map(function (t) { var d = t.sedes[s.id]; return d ? d.resultado : null; }),
      borderColor: c, backgroundColor: c, borderWidth: 2, tension: 0.25, pointRadius: 4, spanGaps: true,
      borderDash: filtroSede === 'ALL' ? [4, 4] : [] });
  });
  ds.push({ label: 'Meta', data: serie.map(function (t) { return t.objetivo; }), borderColor: '#ef4444', borderWidth: 1.5,
    borderDash: [6, 6], pointRadius: 0, pointHoverRadius: 0 });
  var vals = [];
  ds.forEach(function (x) { x.data.forEach(function (v) { if (v !== null) vals.push(v); }); });
  var yMin = Math.max(0, Math.floor(Math.min.apply(null, vals)) - 1);
  charts.res = new Chart(el('cRes'), { type: 'line', data: { labels: labels, datasets: ds }, options: opts(yMin, 10.5, 1) });

  ORDEN_ITEMS.forEach(function (k) {
    var dsK = sedes.map(function (s) {
      return { label: s.nombre, data: serie.map(function (t) { var d = t.sedes[s.id]; return d ? d.items[k].valor : null; }),
        backgroundColor: SEDE_UI[s.id].color, borderRadius: 5, maxBarThickness: 34 };
    });
    charts[k] = new Chart(el('c_' + k), { type: 'bar', data: { labels: labels, datasets: dsK }, options: opts(0, undefined, undefined) });
  });
}

el('fPer').addEventListener('change', render);
el('btnRefresh').addEventListener('click', cargar);
el('btnHelp').addEventListener('click', function () { el('modalHelp').classList.remove('hidden'); });
el('btnHelpRefresh').addEventListener('click', function () { el('modalHelp').classList.add('hidden'); cargar(); });
document.querySelectorAll('[data-close]').forEach(function (b) { b.addEventListener('click', function () { el('modalHelp').classList.add('hidden'); }); });
el('modalHelp').addEventListener('click', function (e) { if (e.target === el('modalHelp')) el('modalHelp').classList.add('hidden'); });
document.addEventListener('keydown', function (e) { if (e.key === 'Escape') el('modalHelp').classList.add('hidden'); });
cargar();
</script>
</body>
</html>
`;
