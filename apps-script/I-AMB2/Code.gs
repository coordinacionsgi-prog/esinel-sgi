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
 *   - Activar actualización automática (diaria)
 *
 * Web App: Implementar > Nueva implementación > Aplicación web.
 *   /exec               -> dashboard
 *   /exec?formato=json  -> datos en JSON
 */

// ---------------------------------------------------------------- Config

var CFG = {
  FICHA: 'AMB2',
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
    .addSeparator()
    .addItem('Activar actualización automática (diaria)', 'activarAutomatico')
    .addToUi();
}

function doGet(e) {
  if (e && e.parameter && e.parameter.formato === 'json') {
    return ContentService.createTextOutput(JSON.stringify(obtenerDatos()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  return HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setTitle('I-AMB2 · Consumos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function abrirDashboard() {
  var html = HtmlService.createTemplateFromFile('Dashboard').evaluate()
    .setWidth(1200).setHeight(820);
  SpreadsheetApp.getUi().showModalDialog(html, 'I-AMB2 · Generación de Consumos');
}

// ---------------------------------------------------------------- Lectura

/** Devuelve todo lo que necesita el dashboard. Lo llama Dashboard.html. */
function obtenerDatos() {
  var ss = SpreadsheetApp.getActive();
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
  var ss = SpreadsheetApp.getActive();
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
  var ss = SpreadsheetApp.getActive();
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
  var ss = SpreadsheetApp.getActive();
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
  var ss = SpreadsheetApp.getActive();
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

// ---------------------------------------------------------------- Automatización

function activarAutomatico() {
  ScriptApp.getProjectTriggers().forEach(function (tr) {
    if (tr.getHandlerFunction() === 'tareaAutomatica') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('tareaAutomatica').timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Listo: todos los días a las 7 h la ficha AMB2 se pone al día sola.');
}

function tareaAutomatica() {
  actualizarFicha();
}

// ---------------------------------------------------------------- Utilidades

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
