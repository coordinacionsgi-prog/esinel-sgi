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
<title>I-AMB2 · Consumos</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {
    color-scheme: light;
    --page: #f9f9f7; --surface: #fcfcfb; --ink: #0b0b0b; --ink-2: #52514e; --muted: #898781;
    --grid: #e1e0d9; --axis: #c3c2b7; --border: rgba(11,11,11,0.10);
    --s1: #2a78d6; --s2: #eb6834; --s3: #1baf7a; --gen: #0b0b0b;
    --good: #0ca30c; --good-ink: #006300; --warn: #fab219; --crit: #d03b3b; --crit-ink: #b42323;
    --chip: #f0efec;
  }
  @media (prefers-color-scheme: dark) {
    :root:not([data-theme="light"]) {
      color-scheme: dark;
      --page: #0d0d0d; --surface: #1a1a19; --ink: #ffffff; --ink-2: #c3c2b7; --muted: #898781;
      --grid: #2c2c2a; --axis: #383835; --border: rgba(255,255,255,0.10);
      --s1: #3987e5; --s2: #d95926; --s3: #199e70; --gen: #ffffff;
      --good-ink: #0ca30c; --crit-ink: #e66767; --chip: #383835;
    }
  }
  * { box-sizing: border-box; }
  body { margin: 0; background: var(--page); color: var(--ink);
    font: 14px/1.45 system-ui, -apple-system, "Segoe UI", Roboto, sans-serif; }
  .wrap { max-width: 1240px; margin: 0 auto; padding: 20px 16px 40px; }
  header { display: flex; flex-wrap: wrap; gap: 12px; align-items: flex-end; justify-content: space-between; margin-bottom: 18px; }
  h1 { font-size: 20px; margin: 0; letter-spacing: -0.01em; }
  .sub { color: var(--ink-2); font-size: 12px; margin-top: 2px; }
  .sub a { color: inherit; }
  .controls { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
  select, button { font: inherit; font-size: 13px; color: var(--ink); background: var(--surface);
    border: 1px solid var(--border); border-radius: 8px; padding: 6px 10px; cursor: pointer; }
  button:hover, select:hover { border-color: var(--axis); }
  .card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px; }
  .kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 12px; }
  .kpi .lbl { font-size: 12px; color: var(--ink-2); display: flex; align-items: center; gap: 6px; }
  .kpi .val { font-size: 30px; font-weight: 700; margin: 4px 0 2px; letter-spacing: -0.02em; }
  .kpi .val small { font-size: 13px; font-weight: 500; color: var(--muted); }
  .kpi .meta { font-size: 12px; color: var(--ink-2); display: flex; gap: 8px; flex-wrap: wrap; align-items: center; }
  .kpi.hero { border-top: 3px solid var(--gen); }
  .dot { width: 10px; height: 10px; border-radius: 50%; display: inline-block; flex: none; }
  .badge { display: inline-flex; align-items: center; gap: 4px; font-size: 11px; font-weight: 600;
    padding: 1px 8px; border-radius: 999px; background: var(--chip); color: var(--ink); }
  .badge::before { content: ""; width: 7px; height: 7px; border-radius: 50%; background: var(--muted); }
  .badge.ok::before { background: var(--good); }
  .badge.bad::before { background: var(--crit); }
  .delta.up { color: var(--good-ink); } .delta.down { color: var(--crit-ink); }
  .grid2 { display: grid; grid-template-columns: 1.4fr 1fr; gap: 12px; margin-bottom: 12px; }
  @media (max-width: 860px) { .grid2 { grid-template-columns: 1fr; } }
  .card h2 { font-size: 14px; margin: 0 0 2px; }
  .card .hint { font-size: 12px; color: var(--muted); margin: 0 0 10px; }
  .card-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; flex-wrap: wrap; }
  .tabs { display: inline-flex; background: var(--chip); border-radius: 8px; padding: 2px; }
  .tabs button { border: 0; background: transparent; padding: 4px 10px; border-radius: 6px; font-size: 12px; }
  .tabs button.on { background: var(--surface); box-shadow: 0 1px 2px rgba(0,0,0,.12); font-weight: 600; }
  .chart-box { position: relative; height: 300px; }
  .legend { display: flex; gap: 14px; flex-wrap: wrap; font-size: 12px; color: var(--ink-2); margin-bottom: 8px; }
  .legend span { display: inline-flex; align-items: center; gap: 6px; }
  .legend .ln { width: 16px; height: 0; border-top: 2px solid; }
  .legend .ln.dash { border-top-style: dashed; }
  .alerts { list-style: none; margin: 0; padding: 0; max-height: 300px; overflow: auto; }
  .alerts li { display: flex; gap: 8px; padding: 8px 0; border-bottom: 1px solid var(--grid); font-size: 12.5px; }
  .alerts li:last-child { border-bottom: 0; }
  .alerts .ic { font-weight: 700; width: 16px; flex: none; text-align: center; }
  .alerts .error .ic { color: var(--crit-ink); } .alerts .aviso .ic { color: #b07a00; }
  .alerts .hoja { color: var(--muted); font-size: 11px; display: block; }
  .table-wrap { overflow-x: auto; }
  table { width: 100%; border-collapse: collapse; font-size: 12.5px; font-variant-numeric: tabular-nums; }
  th, td { padding: 7px 8px; text-align: right; border-bottom: 1px solid var(--grid); white-space: nowrap; }
  th { color: var(--ink-2); font-weight: 600; font-size: 11.5px; position: sticky; top: 0; background: var(--surface); }
  th:first-child, td:first-child, th:nth-child(2), td:nth-child(2) { text-align: left; }
  tr.grp td { background: var(--chip); font-weight: 600; text-align: left; }
  td.res { font-weight: 700; }
  td.bad { color: var(--crit-ink); }
  .empty, .loading { padding: 60px 20px; text-align: center; color: var(--ink-2); }
  .foot { margin-top: 14px; font-size: 11.5px; color: var(--muted); }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <div>
      <h1>I-AMB2 · Generación de consumos</h1>
      <div class="sub" id="sub">Cargando datos de la planilla…</div>
    </div>
    <div class="controls">
      <label class="sub" for="fEj">Ejercicio</label>
      <select id="fEj"></select>
      <button id="btnRefresh" type="button">↻ Actualizar</button>
    </div>
  </header>

  <div id="app"><div class="card loading">Leyendo hojas trimestrales…</div></div>
</div>

<script>
var DATA = null, charts = {}, consumoActivo = 'energia';
var SEDE_VAR = { REM: '--s1', STO: '--s2', MOS: '--s3' };

function css(v) { return getComputedStyle(document.documentElement).getPropertyValue(v).trim(); }
function n(v, d) {
  if (v === null || v === undefined) return '—';
  return Number(v).toLocaleString('es-AR', { minimumFractionDigits: d || 0, maximumFractionDigits: d === undefined ? 2 : d });
}
function esc(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

function cargar() {
  document.getElementById('btnRefresh').disabled = true;
  if (window.google && google.script && google.script.run) {
    google.script.run.withSuccessHandler(recibir).withFailureHandler(function (e) {
      document.getElementById('app').innerHTML = '<div class="card empty">No se pudieron leer los datos: ' + esc(e.message || e) + '</div>';
      document.getElementById('btnRefresh').disabled = false;
    }).obtenerDatos();
  } else if (window.__AMB2_DATA__) {
    recibir(window.__AMB2_DATA__);
  }
}

function recibir(d) {
  DATA = d;
  document.getElementById('btnRefresh').disabled = false;
  document.getElementById('sub').innerHTML = '<a href="' + esc(d.url) + '" target="_blank" rel="noopener">' + esc(d.planilla) +
    '</a> · actualizado ' + esc(d.generado) + ' · objetivo ' + n(d.objetivo, 1) + ' pts';
  var sel = document.getElementById('fEj');
  var ejs = d.ejercicios.map(function (e) { return e.ej; }).reverse();
  var prev = sel.value;
  sel.innerHTML = '<option value="todos">Todos</option>' + ejs.map(function (e) {
    return '<option value="' + e + '">Ej. ' + e + ' (' + (1991 + e) + '/' + String(1992 + e).slice(2) + ')</option>';
  }).join('');
  sel.value = prev && sel.querySelector('option[value="' + prev + '"]') ? prev : (ejs.length ? String(ejs[0]) : 'todos');
  render();
}

function filtrados() {
  var v = document.getElementById('fEj').value;
  return DATA.trimestres.filter(function (t) { return v === 'todos' || String(t.ej) === v; });
}

function render() {
  if (!DATA) return;
  var ts = filtrados();
  if (!DATA.trimestres.length) {
    document.getElementById('app').innerHTML = '<div class="card empty">No hay hojas con nombre "HASTA dd-mm-aa".</div>';
    return;
  }
  var conDato = ts.filter(function (t) { return t.general !== null; });
  var ult = conDato[conDato.length - 1];
  var idxGlobal = ult ? DATA.trimestres.indexOf(ult) : -1;
  var ant = null;
  for (var i = idxGlobal - 1; i >= 0; i--) if (DATA.trimestres[i].general !== null) { ant = DATA.trimestres[i]; break; }
  var ejSel = document.getElementById('fEj').value;
  var resEj = DATA.ejercicios.filter(function (e) { return String(e.ej) === ejSel; })[0];

  var h = '';
  // KPIs
  h += '<div class="kpis">';
  if (ult) {
    h += kpi('hero', '<span class="dot" style="background:var(--gen)"></span>Resultado general · ' + esc(ult.etiqueta),
      n(ult.general, 2), '/ 10', badge(ult.general, ult.objetivo) + delta(ult.general, ant && ant.general) +
      '<span>' + esc(ult.rango) + '</span>');
    DATA.sedes.forEach(function (s) {
      var d = ult.sedes[s.id], pa = null;
      for (var j = idxGlobal - 1; j >= 0; j--) { var x = DATA.trimestres[j].sedes[s.id]; if (x && x.resultado !== null) { pa = x; break; } }
      h += kpi('', '<span class="dot" style="background:var(' + SEDE_VAR[s.id] + ')"></span>' + esc(s.nombre),
        d ? n(d.resultado, 1) : '—', d ? '/ 10' : '',
        d ? badge(d.resultado, s.objetivo) + delta(d.resultado, pa && pa.resultado) +
            '<span>' + n(d.items.energia.valor, 0) + ' kWh</span>' : '<span>Sin datos este trimestre</span>');
    });
    if (resEj) {
      h += kpi('', 'Promedio Ej. ' + resEj.ej, n(resEj.promedio, 2), '/ 10',
        badge(resEj.promedio, DATA.objetivo) + '<span>' + resEj.trimestres + ' de 4 trimestres</span>');
    }
  } else {
    h += '<div class="card empty">Sin resultados en el ejercicio seleccionado.</div>';
  }
  h += '</div>';

  // Gráfico resultados + alertas
  h += '<div class="grid2">';
  h += '<div class="card"><h2>Resultado por trimestre</h2><p class="hint">General ponderado por nómina y por sede · línea punteada = objetivo</p>' +
       legendHTML(true) + '<div class="chart-box"><canvas id="cRes" role="img" aria-label="Resultado por trimestre"></canvas></div></div>';
  var hojas = {}; ts.forEach(function (t) { hojas[t.hoja] = true; });
  var als = DATA.alertas.filter(function (a) { return hojas[a.hoja] || !/^\s*HASTA/i.test(a.hoja); })
    .sort(function (a, b) { return (a.nivel === 'error' ? 0 : 1) - (b.nivel === 'error' ? 0 : 1); });
  var otras = DATA.alertas.length - als.length;
  h += '<div class="card"><h2>Revisión de datos</h2><p class="hint">' +
       (als.length ? als.length + ' punto(s) a revisar' : 'Sin inconsistencias en este período') +
       (otras ? ' · ' + otras + ' más en otros ejercicios' : '') + '</p>' + alertasHTML(als) + '</div>';
  h += '</div>';

  // Consumos
  h += '<div class="card" style="margin-bottom:12px"><div class="card-head"><div><h2>Consumos por sede</h2>' +
       '<p class="hint" id="hintCons"></p></div><div class="tabs" id="tabs">' +
       DATA.items.map(function (it) { return '<button type="button" data-k="' + it.id + '" class="' + (it.id === consumoActivo ? 'on' : '') + '">' + esc(it.nombre) + '</button>'; }).join('') +
       '</div></div>' + legendHTML(false) + '<div class="chart-box"><canvas id="cCons" role="img" aria-label="Consumos por sede"></canvas></div></div>';

  // Tabla
  h += '<div class="card"><h2>Detalle por trimestre</h2><p class="hint">Valores leídos de cada hoja "HASTA …" · nota 2 a 10 por ítem · resultado = suma ponderada</p>' +
       '<div class="table-wrap">' + tablaHTML(ts) + '</div></div>';
  h += '<div class="foot">El período sale del nombre de cada hoja (ejercicio de abril a marzo). El resultado general pondera cada sede por su nómina.</div>';

  document.getElementById('app').innerHTML = h;
  document.querySelectorAll('#tabs button').forEach(function (b) {
    b.onclick = function () { consumoActivo = b.dataset.k; render(); };
  });
  dibujar(ts);
}

function kpi(cls, lbl, val, unit, meta) {
  return '<div class="card kpi ' + cls + '"><div class="lbl">' + lbl + '</div><div class="val">' + val +
    (unit ? ' <small>' + unit + '</small>' : '') + '</div><div class="meta">' + meta + '</div></div>';
}
function badge(v, obj) {
  if (v === null || v === undefined) return '<span class="badge">Sin dato</span>';
  return v >= obj ? '<span class="badge ok">Cumple</span>' : '<span class="badge bad">Desvío</span>';
}
function delta(a, b) {
  if (a === null || b === null || b === undefined) return '';
  var d = a - b;
  if (Math.abs(d) < 0.005) return '<span class="delta">= anterior</span>';
  return '<span class="delta ' + (d > 0 ? 'up' : 'down') + '">' + (d > 0 ? '▲ +' : '▼ ') + n(d, 2) + '</span>';
}
function legendHTML(conGeneral) {
  var h = '<div class="legend">';
  if (conGeneral) h += '<span><i class="ln" style="border-color:var(--gen);border-top-width:3px"></i>General</span>';
  DATA.sedes.forEach(function (s) { h += '<span><i class="dot" style="background:var(' + SEDE_VAR[s.id] + ')"></i>' + esc(s.nombre) + '</span>'; });
  if (conGeneral) h += '<span><i class="ln dash" style="border-color:var(--muted)"></i>Objetivo</span>';
  return h + '</div>';
}
function alertasHTML(als) {
  if (!als.length) return '<p class="sub">✓ Todo en orden.</p>';
  return '<ul class="alerts">' + als.map(function (a) {
    return '<li class="' + a.nivel + '"><span class="ic">' + (a.nivel === 'error' ? '✖' : '!') + '</span><span>' +
      esc(a.msg) + '<span class="hoja">' + esc(a.hoja.trim()) + '</span></span></li>';
  }).join('') + '</ul>';
}
function tablaHTML(ts) {
  var h = '<table><thead><tr><th>Trimestre</th><th>Sede</th><th>Nómina</th><th>Resmas</th><th>Rollos</th><th>Energía kWh</th>' +
          '<th>kWh/persona</th><th>Notas (R/P/E)</th><th>Resultado</th></tr></thead><tbody>';
  ts.slice().reverse().forEach(function (t) {
    h += '<tr class="grp"><td colspan="8">' + esc(t.etiqueta) + ' · ' + esc(t.rango) + ' <span class="sub">(' + esc(t.hoja.trim()) + ')</span></td>' +
         '<td style="text-align:right" class="' + (t.cumple === false ? 'bad' : '') + '">' + n(t.general, 2) + '</td></tr>';
    DATA.sedes.forEach(function (s) {
      var d = t.sedes[s.id];
      if (!d) return;
      var it = d.items;
      h += '<tr><td></td><td><span class="dot" style="background:var(' + SEDE_VAR[s.id] + ');margin-right:6px"></span>' + esc(s.nombre) + '</td>' +
        '<td>' + n(d.nomina, 0) + '</td><td>' + n(it.resmas.valor) + '</td><td>' + n(it.plotter.valor) + '</td>' +
        '<td>' + n(it.energia.valor, 0) + '</td><td>' + n(d.kwhPorPersona, 1) + '</td>' +
        '<td>' + [it.resmas.nota, it.plotter.nota, it.energia.nota].map(function (x) { return n(x, 0); }).join(' / ') + '</td>' +
        '<td class="res ' + (d.resultado !== null && d.resultado < s.objetivo ? 'bad' : '') + '">' + n(d.resultado, 2) + '</td></tr>';
    });
  });
  return h + '</tbody></table>';
}

function baseOpts(yTitle, yMin, yMax) {
  var grid = css('--grid'), muted = css('--muted'), axis = css('--axis');
  return {
    responsive: true, maintainAspectRatio: false, animation: false,
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: css('--surface'), titleColor: css('--ink'), bodyColor: css('--ink-2'),
        borderColor: axis, borderWidth: 1, padding: 10, boxPadding: 4, usePointStyle: true,
        callbacks: { label: function (c) { return ' ' + c.dataset.label + ': ' + (c.raw === null ? '—' : n(c.raw, yMax === 10 ? 2 : 1)); } }
      }
    },
    scales: {
      x: { grid: { display: false }, border: { color: axis }, ticks: { color: muted, font: { size: 11 } } },
      y: { min: yMin, max: yMax, beginAtZero: true, grid: { color: grid }, border: { display: false },
           ticks: { color: muted, font: { size: 11 }, stepSize: yMax === 10 ? 1 : undefined,
                    callback: function (v) { return n(v, 0); } },
           title: { display: !!yTitle, text: yTitle, color: muted, font: { size: 11 } } }
    }
  };
}

function dibujar(ts) {
  Object.keys(charts).forEach(function (k) { charts[k].destroy(); });
  charts = {};
  var labels = ts.map(function (t) { return ts.length > 6 ? t.rango : [t.etiqueta, t.rango]; });
  var surface = css('--surface');

  var dsRes = DATA.sedes.map(function (s) {
    var c = css(SEDE_VAR[s.id]);
    return { label: s.nombre, data: ts.map(function (t) { var d = t.sedes[s.id]; return d ? d.resultado : null; }),
      borderColor: c, backgroundColor: c, borderWidth: 2, pointRadius: 4, pointHoverRadius: 6,
      pointBorderColor: surface, pointBorderWidth: 2, spanGaps: true, tension: 0 };
  });
  dsRes.unshift({ label: 'General', data: ts.map(function (t) { return t.general; }), borderColor: css('--gen'),
    backgroundColor: css('--gen'), borderWidth: 3, pointRadius: 4, pointHoverRadius: 6, pointBorderColor: surface,
    pointBorderWidth: 2, spanGaps: true, tension: 0 });
  dsRes.push({ label: 'Objetivo', data: ts.map(function (t) { return t.objetivo; }), borderColor: css('--muted'),
    borderDash: [5, 4], borderWidth: 1.5, pointRadius: 0, pointHoverRadius: 0, fill: false });
  var vals = [];
  dsRes.forEach(function (d) { d.data.forEach(function (v) { if (v !== null) vals.push(v); }); });
  var yMin = vals.length ? Math.max(0, Math.floor(Math.min.apply(null, vals)) - 1) : 0;
  charts.res = new Chart(document.getElementById('cRes'), {
    type: 'line', data: { labels: labels, datasets: dsRes }, options: baseOpts('Puntos', yMin, 10)
  });

  var it = DATA.items.filter(function (x) { return x.id === consumoActivo; })[0];
  document.getElementById('hintCons').textContent = it.nombre + ' (' + it.unidad + ') por trimestre';
  var dsCons = DATA.sedes.map(function (s) {
    var c = css(SEDE_VAR[s.id]);
    return { label: s.nombre, data: ts.map(function (t) { var d = t.sedes[s.id]; return d ? d.items[consumoActivo].valor : null; }),
      backgroundColor: c, borderColor: surface, borderWidth: { left: 1, right: 1 }, borderRadius: { topLeft: 4, topRight: 4 },
      borderSkipped: 'bottom', maxBarThickness: 28, categoryPercentage: 0.7, barPercentage: 0.92 };
  });
  var oc = baseOpts(it.unidad === 'kWh' ? 'kWh' : 'Unidades', 0, undefined);
  delete oc.scales.y.max;
  charts.cons = new Chart(document.getElementById('cCons'), { type: 'bar', data: { labels: labels, datasets: dsCons }, options: oc });
}

document.getElementById('fEj').addEventListener('change', render);
document.getElementById('btnRefresh').addEventListener('click', cargar);
if (window.matchMedia) window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function () { if (DATA) render(); });
cargar();
</script>
</body>
</html>
`;
