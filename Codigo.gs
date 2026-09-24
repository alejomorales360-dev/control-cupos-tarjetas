// ════════════════════════════════════════════════════════════════
// Control de Cupos Temporales de Tarjetas — Codigo.gs
// ════════════════════════════════════════════════════════════════
//
// Registra clientes con su cupo base y los aumentos temporales de cupo
// (monto + rango de fechas). Todos los días revisa qué aumentos están por
// vencer, vencen hoy o ya vencieron sin haberse cortado, y envía alertas
// por correo y (opcional) Telegram. Opcionalmente crea un evento en
// Google Calendar en la fecha de fin de cada aumento.
//
// ════════ ESTRUCTURA DE HOJAS (se crean solas con setup()) ════════
//
// CLIENTES
//   A:ID  B:NOMBRE  C:DOCUMENTO  D:BANCO  E:TARJETA_ULT4  F:CUPO_BASE
//   G:EMAIL  H:TELEFONO  I:NOTAS  J:ACTIVO  K:CREADO
//
// AUMENTOS
//   A:ID  B:ID_CLIENTE  C:CLIENTE  D:MONTO  E:FECHA_INICIO  F:FECHA_FIN
//   G:MOTIVO  H:ESTADO (ACTIVO | CORTADO | ANULADO)  I:FECHA_CORTE
//   J:NOTA_CORTE  K:ULTIMO_AVISO  L:EVENTO_CALENDAR  M:CREADO
//
// HISTORIAL
//   A:FECHA  B:ACCION  C:DETALLE  D:USUARIO
//
// ════════ CONFIGURACIÓN (Script Properties, editable desde la app) ════════
//   SPREADSHEET_ID   ID de la planilla (lo crea setup())
//   EMAIL_ALERTAS    correos destino, separados por coma
//   DIAS_AVISO       días de anticipación para el aviso previo (default 3)
//   HORA_ALERTA      hora del día (0-23) en que corre la revisión (default 8)
//   TELEGRAM_TOKEN   token del bot de Telegram (opcional)
//   TELEGRAM_CHAT_ID chat donde el bot envía las alertas (opcional)
//   USAR_CALENDAR    'true' para crear eventos en Google Calendar
//
// ════════════════════════════════════════════════════════════════

const APP_NAME     = 'Control de Cupos';
const SH_CLIENTES  = 'CLIENTES';
const SH_AUMENTOS  = 'AUMENTOS';
const SH_HISTORIAL = 'HISTORIAL';

const HEADERS = {
  CLIENTES:  ['ID','NOMBRE','DOCUMENTO','BANCO','TARJETA_ULT4','CUPO_BASE','EMAIL','TELEFONO','NOTAS','ACTIVO','CREADO'],
  AUMENTOS:  ['ID','ID_CLIENTE','CLIENTE','MONTO','FECHA_INICIO','FECHA_FIN','MOTIVO','ESTADO','FECHA_CORTE','NOTA_CORTE','ULTIMO_AVISO','EVENTO_CALENDAR','CREADO'],
  HISTORIAL: ['FECHA','ACCION','DETALLE','USUARIO']
};

const ESTADO_ACTIVO  = 'ACTIVO';
const ESTADO_CORTADO = 'CORTADO';
const ESTADO_ANULADO = 'ANULADO';

const TRIGGER_FN = 'revisarVencimientos';

// ════════════════════════════════════════════════════════════════
// INSTALACIÓN
// ════════════════════════════════════════════════════════════════

/**
 * Ejecutar UNA VEZ desde el editor de Apps Script.
 * Crea la planilla (si no existe), las hojas y el disparador diario.
 */
function setup() {
  const ss = getSS_();
  ensureSheets_(ss);
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('EMAIL_ALERTAS')) {
    props.setProperty('EMAIL_ALERTAS', Session.getEffectiveUser().getEmail());
  }
  if (!props.getProperty('DIAS_AVISO'))  props.setProperty('DIAS_AVISO', '3');
  if (!props.getProperty('HORA_ALERTA')) props.setProperty('HORA_ALERTA', '8');
  instalarTrigger_(Number(props.getProperty('HORA_ALERTA')));
  log_('SETUP', 'Planilla: ' + ss.getUrl());
  Logger.log('Listo. Planilla: ' + ss.getUrl());
  Logger.log('Alertas a: ' + props.getProperty('EMAIL_ALERTAS'));
}

function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* se recrea abajo */ }
  }
  const ss = SpreadsheetApp.create('Control de Cupos de Tarjetas');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  ensureSheets_(ss);
  const hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);
  return ss;
}

function ensureSheets_(ss) {
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      const h = HEADERS[name];
      sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
  });
}

function getSheet_(name) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);
  if (!sh) { ensureSheets_(ss); sh = ss.getSheetByName(name); }
  return sh;
}

function instalarTrigger_(hora) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === TRIGGER_FN) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger(TRIGGER_FN).timeBased().everyDays(1).atHour(hora).create();
}

// ════════════════════════════════════════════════════════════════
// WEB APP
// ════════════════════════════════════════════════════════════════

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ════════════════════════════════════════════════════════════════
// FECHAS (todo se compara como texto 'yyyy-MM-dd' en la zona del script)
// ════════════════════════════════════════════════════════════════

function tz_() { return Session.getScriptTimeZone(); }
function hoyISO_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function pad2_(n) { return ('0' + n).slice(-2); }

function aISO_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})$/); // dd-mm-yyyy
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  return '';
}

// Mediodía para que ningún cambio de zona horaria mueva el día.
function isoAFecha_(iso) {
  const p = iso.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2], 12, 0, 0);
}

function isoADMY_(iso) {
  if (!iso) return '';
  const p = iso.split('-');
  return p[2] + '-' + p[1] + '-' + p[0];
}

function diasEntre_(desdeISO, hastaISO) {
  const a = desdeISO.split('-').map(Number), b = hastaISO.split('-').map(Number);
  return Math.round((Date.UTC(b[0], b[1] - 1, b[2]) - Date.UTC(a[0], a[1] - 1, a[2])) / 86400000);
}

// ════════════════════════════════════════════════════════════════
// LECTURA / ESCRITURA GENÉRICA
// ════════════════════════════════════════════════════════════════

function leer_(name) {
  const sh = getSheet_(name);
  const values = sh.getDataRange().getValues();
  const h = values.shift() || [];
  return values
    .map(function (row, i) {
      const o = { _row: i + 2 };
      h.forEach(function (k, j) { o[k] = row[j]; });
      return o;
    })
    .filter(function (o) { return o.ID !== '' || name === SH_HISTORIAL; });
}

function escribirFila_(name, obj, rowNum) {
  const sh = getSheet_(name);
  const h = HEADERS[name];
  const row = h.map(function (k) { return obj[k] === undefined ? '' : obj[k]; });
  if (rowNum) sh.getRange(rowNum, 1, 1, h.length).setValues([row]);
  else { sh.appendRow(row); rowNum = sh.getLastRow(); }
  formatearFila_(sh, name, rowNum);
  return rowNum;
}

function formatearFila_(sh, name, rowNum) {
  const h = HEADERS[name];
  h.forEach(function (k, j) {
    const c = sh.getRange(rowNum, j + 1);
    if (/^FECHA_|^ULTIMO_AVISO$/.test(k)) c.setNumberFormat('dd-mm-yyyy');
    else if (k === 'CREADO' || (name === SH_HISTORIAL && k === 'FECHA')) c.setNumberFormat('dd-mm-yyyy hh:mm');
    else if (k === 'MONTO' || k === 'CUPO_BASE') c.setNumberFormat('#,##0.00');
    else if (k === 'TARJETA_ULT4' || k === 'DOCUMENTO' || k === 'TELEFONO') c.setNumberFormat('@');
  });
}

function nuevoId_(prefijo) {
  return prefijo + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 8).toUpperCase();
}

function conLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function log_(accion, detalle) {
  try {
    let user = '';
    try { user = Session.getActiveUser().getEmail(); } catch (e) {}
    getSheet_(SH_HISTORIAL).appendRow([new Date(), accion, detalle, user]);
  } catch (e) { Logger.log('No se pudo registrar historial: ' + e); }
}

function num_(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v || '').replace(/[^\d.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function money_(n) {
  return '$' + Number(n || 0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ════════════════════════════════════════════════════════════════
// LÓGICA DE ESTADOS
// ════════════════════════════════════════════════════════════════

/**
 * Estado calculado de un aumento para un día dado:
 *   PROGRAMADO  aún no empieza
 *   VIGENTE     en curso, lejos del fin
 *   POR_VENCER  termina dentro de DIAS_AVISO días
 *   VENCE_HOY   hoy es la fecha de fin → hacer el trámite de corte
 *   VENCIDO     ya pasó la fecha de fin y NO se ha marcado como cortado
 *   CORTADO / ANULADO  cerrados
 */
function estadoCalculado_(a, hoy, diasAviso) {
  if (a.ESTADO === ESTADO_CORTADO) return 'CORTADO';
  if (a.ESTADO === ESTADO_ANULADO) return 'ANULADO';
  const ini = aISO_(a.FECHA_INICIO), fin = aISO_(a.FECHA_FIN);
  if (hoy < ini) return 'PROGRAMADO';
  const d = diasEntre_(hoy, fin);
  if (d < 0) return 'VENCIDO';
  if (d === 0) return 'VENCE_HOY';
  if (d <= diasAviso) return 'POR_VENCER';
  return 'VIGENTE';
}

function aumentoDTO_(a, hoy, diasAviso, clientesById) {
  const c = clientesById[a.ID_CLIENTE] || {};
  const fin = aISO_(a.FECHA_FIN);
  return {
    id: a.ID,
    idCliente: a.ID_CLIENTE,
    cliente: c.NOMBRE || a.CLIENTE,
    banco: c.BANCO || '',
    tarjeta: c.TARJETA_ULT4 ? String(c.TARJETA_ULT4) : '',
    cupoBase: num_(c.CUPO_BASE),
    monto: num_(a.MONTO),
    inicio: aISO_(a.FECHA_INICIO),
    fin: fin,
    motivo: a.MOTIVO || '',
    estado: a.ESTADO,
    estadoCalc: estadoCalculado_(a, hoy, diasAviso),
    diasRestantes: fin ? diasEntre_(hoy, fin) : null,
    fechaCorte: aISO_(a.FECHA_CORTE),
    notaCorte: a.NOTA_CORTE || '',
    ultimoAviso: aISO_(a.ULTIMO_AVISO),
    tieneEvento: !!a.EVENTO_CALENDAR
  };
}

// ════════════════════════════════════════════════════════════════
// API PARA LA INTERFAZ (google.script.run)
// ════════════════════════════════════════════════════════════════

function obtenerDatos() {
  const hoy = hoyISO_();
  const diasAviso = getConfig_().diasAviso;
  const clientes = leer_(SH_CLIENTES);
  const byId = {};
  clientes.forEach(function (c) { byId[c.ID] = c; });

  const aumentos = leer_(SH_AUMENTOS).map(function (a) { return aumentoDTO_(a, hoy, diasAviso, byId); });

  const enCurso = ['VIGENTE', 'POR_VENCER', 'VENCE_HOY'];
  const clientesDTO = clientes.map(function (c) {
    const extra = aumentos
      .filter(function (a) { return a.idCliente === c.ID && enCurso.indexOf(a.estadoCalc) >= 0; })
      .reduce(function (s, a) { return s + a.monto; }, 0);
    const pendientes = aumentos.filter(function (a) { return a.idCliente === c.ID && a.estadoCalc === 'VENCIDO'; }).length;
    return {
      id: c.ID, nombre: c.NOMBRE, documento: String(c.DOCUMENTO || ''), banco: c.BANCO || '',
      tarjeta: String(c.TARJETA_ULT4 || ''), cupoBase: num_(c.CUPO_BASE), email: c.EMAIL || '',
      telefono: String(c.TELEFONO || ''), notas: c.NOTAS || '', activo: c.ACTIVO !== false && c.ACTIVO !== 'NO',
      aumentoVigente: extra, cupoVigente: num_(c.CUPO_BASE) + extra, cortesPendientes: pendientes
    };
  });

  const cuenta = function (e) { return aumentos.filter(function (a) { return a.estadoCalc === e; }).length; };
  return {
    hoy: hoy,
    diasAviso: diasAviso,
    clientes: clientesDTO,
    aumentos: aumentos,
    resumen: {
      vencidos: cuenta('VENCIDO'),
      venceHoy: cuenta('VENCE_HOY'),
      porVencer: cuenta('POR_VENCER'),
      vigentes: cuenta('VIGENTE'),
      programados: cuenta('PROGRAMADO')
    },
    planillaUrl: getSS_().getUrl()
  };
}

function guardarCliente(c) {
  return conLock_(function () {
    const nombre = String(c.nombre || '').trim();
    if (!nombre) throw new Error('El nombre es obligatorio.');
    const ult4 = String(c.tarjeta || '').replace(/\D/g, '').slice(-4);
    const cupo = num_(c.cupoBase);
    if (cupo < 0) throw new Error('El cupo base no puede ser negativo.');

    let existente = null;
    if (c.id) {
      existente = leer_(SH_CLIENTES).filter(function (x) { return x.ID === c.id; })[0];
      if (!existente) throw new Error('Cliente no encontrado.');
    }
    const obj = {
      ID: existente ? existente.ID : nuevoId_('C'),
      NOMBRE: nombre,
      DOCUMENTO: String(c.documento || '').trim(),
      BANCO: String(c.banco || '').trim(),
      TARJETA_ULT4: ult4,
      CUPO_BASE: cupo,
      EMAIL: String(c.email || '').trim(),
      TELEFONO: String(c.telefono || '').trim(),
      NOTAS: String(c.notas || '').trim(),
      ACTIVO: c.activo === false ? 'NO' : 'SI',
      CREADO: existente ? existente.CREADO : new Date()
    };
    escribirFila_(SH_CLIENTES, obj, existente ? existente._row : null);

    // Mantener el nombre copiado en AUMENTOS al día
    if (existente && existente.NOMBRE !== nombre) {
      const sh = getSheet_(SH_AUMENTOS);
      const col = HEADERS.AUMENTOS.indexOf('CLIENTE') + 1;
      leer_(SH_AUMENTOS).forEach(function (a) {
        if (a.ID_CLIENTE === obj.ID) sh.getRange(a._row, col).setValue(nombre);
      });
    }
    log_(existente ? 'CLIENTE_EDITADO' : 'CLIENTE_CREADO', obj.NOMBRE + ' · cupo base ' + money_(cupo));
    return obj.ID;
  });
}

function guardarAumento(a) {
  return conLock_(function () {
    const cliente = leer_(SH_CLIENTES).filter(function (x) { return x.ID === a.idCliente; })[0];
    if (!cliente) throw new Error('Selecciona un cliente válido.');
    const monto = num_(a.monto);
    if (!(monto > 0)) throw new Error('El monto del aumento debe ser mayor a 0.');
    const ini = aISO_(a.inicio), fin = aISO_(a.fin);
    if (!ini || !fin) throw new Error('Las fechas de inicio y fin son obligatorias.');
    if (fin < ini) throw new Error('La fecha de fin no puede ser anterior a la de inicio.');

    let existente = null;
    if (a.id) {
      existente = leer_(SH_AUMENTOS).filter(function (x) { return x.ID === a.id; })[0];
      if (!existente) throw new Error('Aumento no encontrado.');
      if (existente.ESTADO !== ESTADO_ACTIVO) throw new Error('Solo se pueden editar aumentos activos.');
    }

    const obj = {
      ID: existente ? existente.ID : nuevoId_('A'),
      ID_CLIENTE: cliente.ID,
      CLIENTE: cliente.NOMBRE,
      MONTO: monto,
      FECHA_INICIO: isoAFecha_(ini),
      FECHA_FIN: isoAFecha_(fin),
      MOTIVO: String(a.motivo || '').trim(),
      ESTADO: ESTADO_ACTIVO,
      FECHA_CORTE: '',
      NOTA_CORTE: '',
      ULTIMO_AVISO: existente ? existente.ULTIMO_AVISO : '',
      EVENTO_CALENDAR: existente ? existente.EVENTO_CALENDAR : '',
      CREADO: existente ? existente.CREADO : new Date()
    };

    if (getConfig_().usarCalendar) {
      if (obj.EVENTO_CALENDAR) borrarEvento_(obj.EVENTO_CALENDAR);
      obj.EVENTO_CALENDAR = crearEvento_(obj, cliente);
    }

    escribirFila_(SH_AUMENTOS, obj, existente ? existente._row : null);
    log_(existente ? 'AUMENTO_EDITADO' : 'AUMENTO_CREADO',
      cliente.NOMBRE + ' · +' + money_(monto) + ' del ' + isoADMY_(ini) + ' al ' + isoADMY_(fin));
    return obj.ID;
  });
}

function marcarCortado(id, nota) {
  return cerrarAumento_(id, ESTADO_CORTADO, nota);
}

function anularAumento(id, nota) {
  return cerrarAumento_(id, ESTADO_ANULADO, nota);
}

function reabrirAumento(id) {
  return conLock_(function () {
    const a = leer_(SH_AUMENTOS).filter(function (x) { return x.ID === id; })[0];
    if (!a) throw new Error('Aumento no encontrado.');
    a.ESTADO = ESTADO_ACTIVO;
    a.FECHA_CORTE = '';
    a.NOTA_CORTE = '';
    escribirFila_(SH_AUMENTOS, a, a._row);
    log_('AUMENTO_REABIERTO', a.CLIENTE + ' · ' + a.ID);
    return true;
  });
}

function cerrarAumento_(id, estado, nota) {
  return conLock_(function () {
    const a = leer_(SH_AUMENTOS).filter(function (x) { return x.ID === id; })[0];
    if (!a) throw new Error('Aumento no encontrado.');
    if (a.ESTADO !== ESTADO_ACTIVO) throw new Error('Este aumento ya está ' + a.ESTADO.toLowerCase() + '.');
    a.ESTADO = estado;
    a.FECHA_CORTE = new Date();
    a.NOTA_CORTE = String(nota || '').trim();
    if (a.EVENTO_CALENDAR) {
      if (estado === ESTADO_ANULADO) { borrarEvento_(a.EVENTO_CALENDAR); a.EVENTO_CALENDAR = ''; }
      else marcarEventoCortado_(a.EVENTO_CALENDAR);
    }
    escribirFila_(SH_AUMENTOS, a, a._row);
    log_(estado === ESTADO_CORTADO ? 'CUPO_CORTADO' : 'AUMENTO_ANULADO',
      a.CLIENTE + ' · +' + money_(num_(a.MONTO)) + (a.NOTA_CORTE ? ' · ' + a.NOTA_CORTE : ''));
    return true;
  });
}

// ════════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ════════════════════════════════════════════════════════════════

function getConfig_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  return {
    email: p.EMAIL_ALERTAS || '',
    diasAviso: Math.max(0, parseInt(p.DIAS_AVISO || '3', 10) || 0),
    hora: Math.min(23, Math.max(0, parseInt(p.HORA_ALERTA || '8', 10) || 0)),
    telegramToken: p.TELEGRAM_TOKEN || '',
    telegramChatId: p.TELEGRAM_CHAT_ID || '',
    usarCalendar: p.USAR_CALENDAR === 'true'
  };
}

function obtenerConfig() {
  const c = getConfig_();
  const trigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === TRIGGER_FN; });
  return {
    email: c.email,
    diasAviso: c.diasAviso,
    hora: c.hora,
    telegramConfigurado: !!c.telegramToken,
    telegramChatId: c.telegramChatId,
    usarCalendar: c.usarCalendar,
    triggerActivo: trigger,
    zonaHoraria: tz_()
  };
}

function guardarConfig(cfg) {
  const props = PropertiesService.getScriptProperties();
  const emails = String(cfg.email || '').split(/[,;\s]+/).filter(String);
  emails.forEach(function (e) { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('Correo inválido: ' + e); });
  props.setProperty('EMAIL_ALERTAS', emails.join(','));
  props.setProperty('DIAS_AVISO', String(Math.max(0, parseInt(cfg.diasAviso, 10) || 0)));
  const hora = Math.min(23, Math.max(0, parseInt(cfg.hora, 10) || 0));
  props.setProperty('HORA_ALERTA', String(hora));
  props.setProperty('USAR_CALENDAR', cfg.usarCalendar ? 'true' : 'false');
  if (cfg.telegramToken) props.setProperty('TELEGRAM_TOKEN', String(cfg.telegramToken).trim());
  if (cfg.borrarTelegram) { props.deleteProperty('TELEGRAM_TOKEN'); props.deleteProperty('TELEGRAM_CHAT_ID'); }
  if (cfg.telegramChatId !== undefined && !cfg.borrarTelegram) props.setProperty('TELEGRAM_CHAT_ID', String(cfg.telegramChatId).trim());
  instalarTrigger_(hora);
  log_('CONFIG', 'Alertas: ' + emails.join(',') + ' · aviso ' + cfg.diasAviso + ' días · hora ' + hora);
  return obtenerConfig();
}

// ════════════════════════════════════════════════════════════════
// REVISIÓN DIARIA Y ALERTAS
// ════════════════════════════════════════════════════════════════

/** Lo ejecuta el disparador diario. También se puede correr a mano. */
function revisarVencimientos() {
  const hoy = hoyISO_();
  const cfg = getConfig_();
  const clientes = leer_(SH_CLIENTES);
  const byId = {};
  clientes.forEach(function (c) { byId[c.ID] = c; });

  const filas = leer_(SH_AUMENTOS);
  const items = filas
    .filter(function (a) { return a.ESTADO === ESTADO_ACTIVO; })
    .map(function (a) { return aumentoDTO_(a, hoy, cfg.diasAviso, byId); });

  const porFin = function (x, y) { return x.fin < y.fin ? -1 : x.fin > y.fin ? 1 : 0; };
  const grupos = {
    vencidos:  items.filter(function (a) { return a.estadoCalc === 'VENCIDO'; }).sort(porFin),
    venceHoy:  items.filter(function (a) { return a.estadoCalc === 'VENCE_HOY'; }).sort(porFin),
    porVencer: items.filter(function (a) { return a.estadoCalc === 'POR_VENCER'; }).sort(porFin),
    inicianHoy: items.filter(function (a) { return a.inicio === hoy; })
  };

  const total = grupos.vencidos.length + grupos.venceHoy.length + grupos.porVencer.length;
  if (total === 0) {
    Logger.log('Sin alertas para ' + hoy);
    return { enviado: false, total: 0 };
  }

  const resultado = enviarAlertas_(grupos, hoy, cfg);

  // Marcar ULTIMO_AVISO en los aumentos notificados
  const sh = getSheet_(SH_AUMENTOS);
  const col = HEADERS.AUMENTOS.indexOf('ULTIMO_AVISO') + 1;
  const notificados = {};
  grupos.vencidos.concat(grupos.venceHoy, grupos.porVencer).forEach(function (a) { notificados[a.id] = true; });
  filas.forEach(function (f) {
    if (notificados[f.ID]) sh.getRange(f._row, col).setValue(isoAFecha_(hoy)).setNumberFormat('dd-mm-yyyy');
  });

  log_('ALERTA_ENVIADA', total + ' aumento(s) · ' + resultado.canales.join(', '));
  return { enviado: true, total: total, canales: resultado.canales, errores: resultado.errores };
}

/** Botón "Revisar ahora" de la interfaz. */
function revisarAhora() {
  return revisarVencimientos();
}

/** Botón "Enviar prueba" de la interfaz. */
function enviarPrueba() {
  const cfg = getConfig_();
  const hoy = hoyISO_();
  const ejemplo = {
    id: 'PRUEBA', cliente: 'Cliente de prueba', banco: 'Banco', tarjeta: '1234', cupoBase: 5000,
    monto: 2000, inicio: hoy, fin: hoy, diasRestantes: 0, estadoCalc: 'VENCE_HOY', motivo: 'Mensaje de prueba'
  };
  const r = enviarAlertas_({ vencidos: [], venceHoy: [ejemplo], porVencer: [], inicianHoy: [] }, hoy, cfg, true);
  log_('PRUEBA_ALERTA', r.canales.join(', ') + (r.errores.length ? ' · errores: ' + r.errores.join(' | ') : ''));
  return r;
}

function enviarAlertas_(g, hoy, cfg, esPrueba) {
  const canales = [], errores = [];
  const urgentes = g.vencidos.length + g.venceHoy.length;
  const asunto = (esPrueba ? '[PRUEBA] ' : '') +
    (urgentes ? '⚠️ ' + urgentes + ' cupo(s) por cortar hoy' : '🔔 ' + g.porVencer.length + ' cupo(s) próximos a vencer') +
    ' · ' + isoADMY_(hoy);

  if (cfg.email) {
    try {
      MailApp.sendEmail({ to: cfg.email, subject: asunto, htmlBody: htmlAlerta_(g, hoy, cfg), body: textoAlerta_(g, hoy, false), name: APP_NAME });
      canales.push('email');
    } catch (e) { errores.push('Email: ' + e.message); }
  }
  if (cfg.telegramToken && cfg.telegramChatId) {
    try {
      enviarTelegram_(cfg, (esPrueba ? '<b>[PRUEBA]</b>\n' : '') + textoAlerta_(g, hoy, true));
      canales.push('telegram');
    } catch (e) { errores.push('Telegram: ' + e.message); }
  }
  if (!canales.length && !errores.length) errores.push('No hay ningún canal de alerta configurado.');
  return { canales: canales, errores: errores };
}

function lineaAumento_(a) {
  const tarjeta = [a.banco, a.tarjeta ? '****' + a.tarjeta : ''].filter(String).join(' ');
  return a.cliente + (tarjeta ? ' (' + tarjeta + ')' : '') +
    ' · +' + money_(a.monto) + ' del ' + isoADMY_(a.inicio) + ' al ' + isoADMY_(a.fin) +
    ' · cupo debe volver a ' + money_(a.cupoBase);
}

function escHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function textoAlerta_(g, hoy, html) {
  const b = function (s) { return html ? '<b>' + escHtml_(s) + '</b>' : s.toUpperCase(); };
  const e = function (s) { return html ? escHtml_(s) : s; };
  const out = [b('Control de cupos · ' + isoADMY_(hoy)), ''];
  if (g.vencidos.length) {
    out.push('🔴 ' + b('VENCIDOS — CORTE PENDIENTE'));
    g.vencidos.forEach(function (a) { out.push('• ' + e(lineaAumento_(a)) + ' · venció hace ' + (-a.diasRestantes) + ' día(s)'); });
    out.push('');
  }
  if (g.venceHoy.length) {
    out.push('🟠 ' + b('VENCEN HOY — hacer trámite con el banco'));
    g.venceHoy.forEach(function (a) { out.push('• ' + e(lineaAumento_(a))); });
    out.push('');
  }
  if (g.porVencer.length) {
    out.push('🟡 ' + b('PRÓXIMOS A VENCER'));
    g.porVencer.forEach(function (a) { out.push('• ' + e(lineaAumento_(a)) + ' · faltan ' + a.diasRestantes + ' día(s)'); });
    out.push('');
  }
  if (g.inicianHoy.length) {
    out.push('🟢 ' + b('INICIAN HOY'));
    g.inicianHoy.forEach(function (a) { out.push('• ' + e(lineaAumento_(a))); });
    out.push('');
  }
  out.push(e('Cuando hagas el corte, márcalo como "Cortado" en la app para dejar de recibir este aviso.'));
  const url = urlApp_();
  if (url) out.push(url);
  return out.join('\n');
}

function htmlAlerta_(g, hoy, cfg) {
  const seccion = function (titulo, color, lista, extra) {
    if (!lista.length) return '';
    const filas = lista.map(function (a) {
      const tarjeta = [a.banco, a.tarjeta ? '****' + a.tarjeta : ''].filter(String).join(' ');
      return '<tr>' +
        '<td style="padding:8px;border-bottom:1px solid #eee"><b>' + escHtml_(a.cliente) + '</b><br><span style="color:#666;font-size:12px">' + escHtml_(tarjeta) + '</span></td>' +
        '<td style="padding:8px;border-bottom:1px solid #eee;text-align:right">+' + money_(a.monto) + '</td>' +
        '<td style="padding:8px;border-bottom:1px solid #eee">' + isoADMY_(a.inicio) + ' → <b>' + isoADMY_(a.fin) + '</b></td>' +
        '<td style="padding:8px;border-bottom:1px solid #eee;text-align:right">' + money_(a.cupoBase) + '</td>' +
        '<td style="padding:8px;border-bottom:1px solid #eee;color:#666">' + escHtml_(extra(a)) + '</td>' +
        '</tr>';
    }).join('');
    return '<h3 style="margin:24px 0 8px;color:' + color + '">' + titulo + '</h3>' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px">' +
      '<tr style="background:#f5f5f5;text-align:left"><th style="padding:8px">Cliente</th><th style="padding:8px;text-align:right">Aumento</th><th style="padding:8px">Vigencia</th><th style="padding:8px;text-align:right">Cupo base</th><th style="padding:8px"></th></tr>' +
      filas + '</table>';
  };
  const url = urlApp_();
  return '<div style="font-family:Arial,sans-serif;max-width:720px;color:#111">' +
    '<h2 style="margin:0">Control de cupos · ' + isoADMY_(hoy) + '</h2>' +
    seccion('🔴 Vencidos — corte pendiente', '#b91c1c', g.vencidos, function (a) { return 'venció hace ' + (-a.diasRestantes) + ' día(s)'; }) +
    seccion('🟠 Vencen hoy — hacer trámite con el banco', '#c2410c', g.venceHoy, function () { return 'hoy'; }) +
    seccion('🟡 Próximos a vencer (' + cfg.diasAviso + ' días)', '#a16207', g.porVencer, function (a) { return 'faltan ' + a.diasRestantes + ' día(s)'; }) +
    seccion('🟢 Inician hoy', '#15803d', g.inicianHoy, function () { return ''; }) +
    '<p style="margin-top:24px;color:#444">Cuando hagas el corte con el banco, márcalo como <b>Cortado</b> en la app para dejar de recibir el aviso.</p>' +
    (url ? '<p><a href="' + url + '" style="background:#1f2937;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none">Abrir la app</a></p>' : '') +
    '</div>';
}

function urlApp_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; }
}

// ════════════════════════════════════════════════════════════════
// TELEGRAM (opcional)
// ════════════════════════════════════════════════════════════════

function enviarTelegram_(cfg, texto) {
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + cfg.telegramToken + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ chat_id: cfg.telegramChatId, text: texto, parse_mode: 'HTML', disable_web_page_preview: true }),
    muteHttpExceptions: true
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) throw new Error(body.description || ('HTTP ' + res.getResponseCode()));
}

/**
 * Busca el chat_id del último mensaje que recibió el bot.
 * Pasos: abre tu bot en Telegram, envíale cualquier mensaje y luego pulsa
 * "Detectar chat" en la app.
 */
function detectarChatTelegram() {
  const cfg = getConfig_();
  if (!cfg.telegramToken) throw new Error('Primero guarda el token del bot.');
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + cfg.telegramToken + '/getUpdates', { muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) throw new Error('Telegram: ' + (body.description || 'token inválido'));
  const ups = (body.result || []).filter(function (u) { return u.message && u.message.chat; });
  if (!ups.length) throw new Error('El bot no ha recibido mensajes. Escríbele algo en Telegram y vuelve a intentar.');
  const chat = ups[ups.length - 1].message.chat;
  PropertiesService.getScriptProperties().setProperty('TELEGRAM_CHAT_ID', String(chat.id));
  return { chatId: String(chat.id), nombre: chat.title || [chat.first_name, chat.last_name].filter(String).join(' ') || chat.username || '' };
}

// ════════════════════════════════════════════════════════════════
// GOOGLE CALENDAR (opcional)
// ════════════════════════════════════════════════════════════════

function crearEvento_(aum, cliente) {
  try {
    const fin = isoAFecha_(aISO_(aum.FECHA_FIN));
    const tarjeta = [cliente.BANCO, cliente.TARJETA_ULT4 ? '****' + cliente.TARJETA_ULT4 : ''].filter(String).join(' ');
    const ev = CalendarApp.getDefaultCalendar().createAllDayEvent(
      '⚠️ Cortar cupo: ' + cliente.NOMBRE + ' +' + money_(aum.MONTO),
      fin,
      {
        description: 'Aumento temporal de cupo que vence hoy.\n\n' +
          'Cliente: ' + cliente.NOMBRE + '\n' +
          (tarjeta ? 'Tarjeta: ' + tarjeta + '\n' : '') +
          'Aumento: +' + money_(aum.MONTO) + '\n' +
          'Vigencia: ' + isoADMY_(aISO_(aum.FECHA_INICIO)) + ' al ' + isoADMY_(aISO_(aum.FECHA_FIN)) + '\n' +
          'El cupo debe volver a: ' + money_(num_(cliente.CUPO_BASE)) + '\n' +
          (aum.MOTIVO ? 'Motivo: ' + aum.MOTIVO + '\n' : '') +
          '\nID: ' + aum.ID
      }
    );
    ev.removeAllReminders();
    ev.addPopupReminder(15 * 60);  // 09:00 del día anterior
    ev.addPopupReminder(0);        // al comenzar el día de vencimiento
    return ev.getId();
  } catch (e) {
    Logger.log('No se pudo crear el evento: ' + e);
    return '';
  }
}

function borrarEvento_(eventId) {
  try {
    const ev = CalendarApp.getDefaultCalendar().getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (e) { Logger.log('No se pudo borrar el evento: ' + e); }
}

function marcarEventoCortado_(eventId) {
  try {
    const ev = CalendarApp.getDefaultCalendar().getEventById(eventId);
    if (ev) {
      ev.setTitle(ev.getTitle().replace(/^⚠️ Cortar cupo/, '✅ Cupo cortado'));
      ev.removeAllReminders();
    }
  } catch (e) { Logger.log('No se pudo actualizar el evento: ' + e); }
}
