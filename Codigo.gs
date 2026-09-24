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
const APP_URL      = 'https://alejomorales360-dev.github.io/control-cupos-tarjetas/';
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
  exigirDueno_();
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
  if (!props.getProperty('LOGIN_HASH')) generarClaveAcceso();
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

/**
 * Cambia la planilla donde se guardan los datos (URL completa o ID).
 * Crea en ella las hojas CLIENTES, AUMENTOS e HISTORIAL si no existen;
 * no borra ni modifica otras hojas que ya tenga.
 */
function cambiarPlanilla(token, urlOId) {
  exigirSesion_(token);
  const txt = String(urlOId || '').trim();
  const m = txt.match(/\/d\/([a-zA-Z0-9_-]{20,})/) || txt.match(/^([a-zA-Z0-9_-]{20,})$/);
  if (!m) throw new Error('Pega la URL completa de la planilla (https://docs.google.com/spreadsheets/d/…).');
  let ss;
  try { ss = SpreadsheetApp.openById(m[1]); }
  catch (e) { throw new Error('No pude abrir esa planilla. Revisa la URL y que sea de la misma cuenta de Google que el proyecto de Apps Script.'); }
  ensureSheets_(ss);
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() === 0 && h.getLastColumn() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  log_('PLANILLA_CAMBIADA', ss.getName() + ' · ' + ss.getUrl());
  return { nombre: ss.getName(), url: ss.getUrl() };
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

// La interfaz vive en GitHub Pages (APP_URL). Este proyecto solo expone
// una API JSON por POST: { fn: 'nombreFuncion', args: [...] }.
// Se envía como text/plain para que el navegador no haga preflight CORS.

function doGet() {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;text-align:center;margin-top:20vh">' +
    '<p>Control de Cupos se abre en <a href="' + APP_URL + '" target="_top">' + APP_URL + '</a></p></div>'
  ).setTitle(APP_NAME).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  // Solo estas funciones se pueden llamar desde fuera. Todas, salvo el
  // login, validan el token de sesión que llega como primer argumento.
  const API = {
    iniciarSesion: iniciarSesion,
    cerrarSesion: cerrarSesion,
    cambiarCredenciales: cambiarCredenciales,
    obtenerDatos: obtenerDatos,
    guardarCliente: guardarCliente,
    guardarAumento: guardarAumento,
    marcarCortado: marcarCortado,
    anularAumento: anularAumento,
    reabrirAumento: reabrirAumento,
    obtenerConfig: obtenerConfig,
    guardarConfig: guardarConfig,
    revisarAhora: revisarAhora,
    enviarPrueba: enviarPrueba,
    detectarChatTelegram: detectarChatTelegram,
    importarDatos: importarDatos,
    cambiarPlanilla: cambiarPlanilla
  };
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return jsonOut_({ ok: false, error: 'Solicitud inválida.' }); }
  const fn = req && Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
  if (!fn) return jsonOut_({ ok: false, error: 'Operación no permitida.' });
  try {
    return jsonOut_({ ok: true, data: fn.apply(null, Array.isArray(req.args) ? req.args : []) });
  } catch (err) {
    return jsonOut_({ ok: false, error: err && err.message ? err.message : String(err) });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════
// LOGIN (usuario + clave propios de la app)
// ════════════════════════════════════════════════════════════════
//
// La web app se implementa con acceso "Cualquier persona" para poder entrar
// desde cualquier equipo. La página en sí no contiene datos: todo lo que
// lee o modifica datos exige un token de sesión válido.
//
//   - La clave se guarda como hash SHA-256 iterado con sal (nunca en texto).
//   - Las sesiones se guardan como hash del token, con vencimiento.
//   - Tras MAX_FALLOS intentos fallidos se bloquea el login BLOQUEO_MIN minutos.
//   - La clave inicial se genera ejecutando generarClaveAcceso() en el editor.

const SESION_HORAS_CORTA = 12;
const SESION_DIAS_LARGA  = 30;
const MAX_FALLOS         = 8;
const BLOQUEO_MIN        = 15;
const HASH_ITER          = 1000;

/** Verdadero solo cuando la función la ejecuta el dueño (p. ej. desde el editor). */
function esDueno_() {
  let activo = '', dueno = '';
  try { activo = (Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  try { dueno = (Session.getEffectiveUser().getEmail() || '').toLowerCase(); } catch (e) {}
  return !!activo && activo === dueno;
}

function exigirDueno_() {
  if (!esDueno_()) throw new Error('Esta función solo se puede ejecutar desde el editor de Apps Script.');
}

function sha256_(texto) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, texto, Utilities.Charset.UTF_8));
}

function hashClave_(clave, sal) {
  let h = sal + ':' + clave;
  for (let i = 0; i < HASH_ITER; i++) h = sha256_(h + ':' + sal);
  return h;
}

function claveAleatoria_(largo) {
  const abc = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let out = '';
  while (out.length < largo) {
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid()).forEach(function (b) {
      if (out.length < largo) out += abc[(b + 256) % abc.length];
    });
  }
  return out;
}

function guardarCredenciales_(usuario, clave) {
  const sal = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    LOGIN_USUARIO: String(usuario).trim().toLowerCase(),
    LOGIN_SAL: sal,
    LOGIN_HASH: hashClave_(clave, sal)
  });
  borrarSesiones_();
}

/**
 * Ejecutar desde el editor para crear (o resetear si la olvidaste) la clave.
 * La clave nueva aparece en el Registro de ejecución. Cierra todas las sesiones.
 */
function generarClaveAcceso() {
  exigirDueno_();
  const props = PropertiesService.getScriptProperties();
  const usuario = props.getProperty('LOGIN_USUARIO') || 'admin';
  const clave = claveAleatoria_(12);
  guardarCredenciales_(usuario, clave);
  CacheService.getScriptCache().remove('LOGIN_FALLOS');
  log_('CLAVE_GENERADA', 'Clave regenerada desde el editor');
  Logger.log('Usuario: ' + usuario);
  Logger.log('Clave:   ' + clave);
  Logger.log('Entra a la app y cámbiala en Configuración → Acceso.');
}

function iniciarSesion(usuario, clave, recordar) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('LOGIN_FALLOS') || 0);
  if (fallos >= MAX_FALLOS) throw new Error('Demasiados intentos fallidos. Espera ' + BLOQUEO_MIN + ' minutos.');

  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.LOGIN_HASH) throw new Error('La clave aún no está configurada. Ejecuta generarClaveAcceso() en el editor de Apps Script.');

  const okUsuario = String(usuario || '').trim().toLowerCase() === p.LOGIN_USUARIO;
  const okClave = hashClave_(String(clave || ''), p.LOGIN_SAL) === p.LOGIN_HASH;
  if (!okUsuario || !okClave) {
    cache.put('LOGIN_FALLOS', String(fallos + 1), BLOQUEO_MIN * 60);
    Utilities.sleep(1000);
    log_('LOGIN_FALLIDO', 'Usuario: ' + String(usuario || '').slice(0, 60));
    throw new Error('Usuario o clave incorrectos.');
  }
  cache.remove('LOGIN_FALLOS');

  limpiarSesionesVencidas_();
  const token = Utilities.getUuid() + Utilities.getUuid();
  const vence = Date.now() + (recordar ? SESION_DIAS_LARGA * 86400000 : SESION_HORAS_CORTA * 3600000);
  PropertiesService.getScriptProperties().setProperty('SES_' + sha256_(token), String(vence));
  log_('LOGIN', recordar ? 'Sesión de ' + SESION_DIAS_LARGA + ' días' : 'Sesión de ' + SESION_HORAS_CORTA + ' horas');
  return { token: token, vence: vence, usuario: p.LOGIN_USUARIO };
}

function cerrarSesion(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('SES_' + sha256_(String(token)));
  return true;
}

function exigirSesion_(token) {
  if (!token) throw new Error('SESION_INVALIDA');
  const key = 'SES_' + sha256_(String(token));
  const props = PropertiesService.getScriptProperties();
  const vence = Number(props.getProperty(key) || 0);
  if (!vence || vence < Date.now()) {
    if (vence) props.deleteProperty(key);
    throw new Error('SESION_INVALIDA');
  }
}

function limpiarSesionesVencidas_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties(), ahora = Date.now();
  Object.keys(all).forEach(function (k) {
    if (k.indexOf('SES_') === 0 && Number(all[k]) < ahora) props.deleteProperty(k);
  });
}

function borrarSesiones_() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).forEach(function (k) { if (k.indexOf('SES_') === 0) props.deleteProperty(k); });
}

function cambiarCredenciales(token, actual, nuevoUsuario, nuevaClave) {
  exigirSesion_(token);
  const p = PropertiesService.getScriptProperties().getProperties();
  if (hashClave_(String(actual || ''), p.LOGIN_SAL) !== p.LOGIN_HASH) throw new Error('La clave actual no es correcta.');
  const usuario = String(nuevoUsuario || '').trim().toLowerCase() || p.LOGIN_USUARIO;
  if (!/^[a-z0-9._@-]{3,60}$/.test(usuario)) throw new Error('Usuario inválido (3 a 60 caracteres: letras, números, . _ - @).');
  const clave = String(nuevaClave || '');
  if (clave.length < 8) throw new Error('La nueva clave debe tener al menos 8 caracteres.');
  guardarCredenciales_(usuario, clave);
  log_('CLAVE_CAMBIADA', 'Usuario: ' + usuario);
  return true; // todas las sesiones quedan cerradas; hay que volver a entrar
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
  let t = String(v || '').replace(/[^\d.,\-]/g, '');
  if (/,\d{1,2}$/.test(t) || /^\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, '').replace(',', '.'); // 2.000,50 · 2.000
  else t = t.replace(/,/g, '');                                                                          // 2,000.50 · 2,000
  const n = Number(t);
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

function obtenerDatos(token) {
  exigirSesion_(token);
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
      tarjeta: String(c.TARJETA_ULT4 || ''), cupoBase: c.CUPO_BASE === '' ? null : num_(c.CUPO_BASE), email: c.EMAIL || '',
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

/**
 * Importación masiva (desde Excel). Cada fila:
 *   { nombre, cupoBase?, banco?, tarjeta?, documento?, monto?, inicio?, fin? }
 * - Cliente nuevo → se crea. Si ya existe (mismo nombre, sin distinguir
 *   mayúsculas/tildes) → se reutiliza y se completan los datos que vengan.
 * - Si la fila trae monto + inicio + fin → se crea el aumento, salvo que ya
 *   exista uno activo idéntico (así reimportar el mismo archivo no duplica).
 * Con soloValidar=true no escribe nada: devuelve la vista previa.
 */
function importarDatos(token, filas, soloValidar) {
  exigirSesion_(token);
  if (!Array.isArray(filas) || !filas.length) throw new Error('No hay filas para importar.');
  if (filas.length > 2000) throw new Error('Máximo 2000 filas por importación.');
  return conLock_(function () {
    const clave = function (n) { return String(n || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim().toUpperCase(); };
    const clientes = leer_(SH_CLIENTES);
    const porNombre = {};
    clientes.forEach(function (c) { porNombre[clave(c.NOMBRE)] = c; });
    const aumentos = leer_(SH_AUMENTOS);
    const firmaAum = function (idCliente, monto, ini, fin) { return [idCliente, Number(monto).toFixed(2), ini, fin].join('|'); };
    const existentes = {};
    aumentos.forEach(function (a) {
      if (a.ESTADO === ESTADO_ACTIVO) existentes[firmaAum(a.ID_CLIENTE, num_(a.MONTO), aISO_(a.FECHA_INICIO), aISO_(a.FECHA_FIN))] = true;
    });

    const res = { clientesNuevos: 0, clientesActualizados: 0, aumentosNuevos: 0, omitidos: 0, errores: [], filas: [] };
    const nuevosClientes = [], actualizaciones = [], nuevosAumentos = [];
    const vistos = {};

    filas.forEach(function (f, i) {
      const fila = f.fila || i + 2;
      const nombre = String(f.nombre || '').replace(/\s+/g, ' ').trim();
      const k = clave(nombre);
      const out = { fila: fila, nombre: nombre, cliente: '', aumento: '' };
      try {
        if (!nombre) throw new Error('Falta el nombre.');
        const cupo = f.cupoBase === '' || f.cupoBase == null ? null : num_(f.cupoBase);
        if (cupo != null && cupo < 0) throw new Error('Cupo base negativo.');
        const ult4 = String(f.tarjeta || '').replace(/\D/g, '').slice(-4);

        const tieneAum = f.monto !== '' && f.monto != null || f.inicio || f.fin;
        let ini = '', fin = '', monto = 0;
        if (tieneAum) {
          monto = num_(f.monto); ini = aISO_(f.inicio); fin = aISO_(f.fin);
          if (!(monto > 0)) throw new Error('Monto solicitado inválido.');
          if (!ini || !fin) throw new Error('Faltan fechas del aumento (o tienen formato inválido).');
          if (fin < ini) throw new Error('La fecha de término es anterior a la de aumento.');
        }

        let cli = porNombre[k] || vistos[k];
        if (!cli) {
          cli = {
            ID: nuevoId_('C'), NOMBRE: nombre, DOCUMENTO: String(f.documento || '').trim(), BANCO: String(f.banco || '').trim(),
            TARJETA_ULT4: ult4, CUPO_BASE: cupo == null ? '' : cupo, EMAIL: '', TELEFONO: '', NOTAS: '', ACTIVO: 'SI', CREADO: new Date()
          };
          vistos[k] = cli; nuevosClientes.push(cli); res.clientesNuevos++; out.cliente = 'nuevo';
        } else {
          const cambios = {};
          if (cupo != null && num_(cli.CUPO_BASE) !== cupo) cambios.CUPO_BASE = cupo;
          if (f.banco && !cli.BANCO) cambios.BANCO = String(f.banco).trim();
          if (ult4 && !cli.TARJETA_ULT4) cambios.TARJETA_ULT4 = ult4;
          if (f.documento && !cli.DOCUMENTO) cambios.DOCUMENTO = String(f.documento).trim();
          if (Object.keys(cambios).length && cli._row) {
            Object.keys(cambios).forEach(function (c) { cli[c] = cambios[c]; });
            if (actualizaciones.indexOf(cli) < 0) { actualizaciones.push(cli); res.clientesActualizados++; }
            out.cliente = 'actualizado';
          } else {
            out.cliente = 'existente';
          }
        }

        if (tieneAum) {
          const firma = firmaAum(cli.ID, monto, ini, fin);
          if (existentes[firma]) { out.aumento = 'ya existía'; res.omitidos++; }
          else {
            existentes[firma] = true;
            nuevosAumentos.push({
              ID: nuevoId_('A'), ID_CLIENTE: cli.ID, CLIENTE: cli.NOMBRE, MONTO: monto,
              FECHA_INICIO: isoAFecha_(ini), FECHA_FIN: isoAFecha_(fin), MOTIVO: String(f.motivo || 'Importado desde Excel').trim(),
              ESTADO: ESTADO_ACTIVO, FECHA_CORTE: '', NOTA_CORTE: '', ULTIMO_AVISO: '', EVENTO_CALENDAR: '', CREADO: new Date(),
              _cliente: cli
            });
            res.aumentosNuevos++; out.aumento = 'nuevo +' + money_(monto) + ' del ' + isoADMY_(ini) + ' al ' + isoADMY_(fin);
          }
        }
      } catch (e) {
        res.errores.push({ fila: fila, nombre: nombre, error: e.message });
        out.error = e.message;
      }
      res.filas.push(out);
    });

    if (soloValidar) return res;

    const usarCal = getConfig_().usarCalendar;
    nuevosAumentos.forEach(function (a) { if (usarCal) a.EVENTO_CALENDAR = crearEvento_(a, a._cliente); });
    agregarFilas_(SH_CLIENTES, nuevosClientes);
    actualizaciones.forEach(function (c) { escribirFila_(SH_CLIENTES, c, c._row); });
    agregarFilas_(SH_AUMENTOS, nuevosAumentos);
    const cfgImp = getConfig_();
    if (nuevosAumentos.length && cfgImp.icsInvitacion && cfgImp.email) {
      // Un solo correo con todos los eventos (cuida la cuota diaria de correos)
      try { enviarInvitacionIcs_(cfgImp, nuevosAumentos); } catch (e) { Logger.log('No se pudo enviar la invitación: ' + e); }
    }
    log_('IMPORTACION', res.clientesNuevos + ' clientes nuevos, ' + res.clientesActualizados + ' actualizados, ' +
      res.aumentosNuevos + ' aumentos nuevos, ' + res.errores.length + ' filas con error');
    delete res.filas;
    return res;
  });
}

/** Escribe muchas filas de una vez (mucho más rápido que appendRow en bucle). */
function agregarFilas_(name, objs) {
  if (!objs.length) return;
  const sh = getSheet_(name);
  const h = HEADERS[name];
  const start = sh.getLastRow() + 1;
  sh.getRange(start, 1, objs.length, h.length).setValues(objs.map(function (o) {
    return h.map(function (k) { return o[k] === undefined ? '' : o[k]; });
  }));
  h.forEach(function (k, j) {
    const r = sh.getRange(start, j + 1, objs.length, 1);
    if (/^FECHA_|^ULTIMO_AVISO$/.test(k)) r.setNumberFormat('dd-mm-yyyy');
    else if (k === 'CREADO') r.setNumberFormat('dd-mm-yyyy hh:mm');
    else if (k === 'MONTO' || k === 'CUPO_BASE') r.setNumberFormat('#,##0.00');
  });
}

function guardarCliente(token, c) {
  exigirSesion_(token);
  return conLock_(function () {
    const nombre = String(c.nombre || '').trim();
    if (!nombre) throw new Error('El nombre es obligatorio.');
    const ult4 = String(c.tarjeta || '').replace(/\D/g, '').slice(-4);
    const cupo = c.cupoBase === '' || c.cupoBase == null ? '' : num_(c.cupoBase);
    if (cupo !== '' && cupo < 0) throw new Error('El cupo base no puede ser negativo.');

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
    log_(existente ? 'CLIENTE_EDITADO' : 'CLIENTE_CREADO', obj.NOMBRE + (cupo === '' ? '' : ' · cupo base ' + money_(cupo)));
    return obj.ID;
  });
}

function guardarAumento(token, a) {
  exigirSesion_(token);
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
    const cfgAum = getConfig_();
    if (cfgAum.icsInvitacion && cfgAum.email) {
      try { enviarInvitacionIcs_(cfgAum, [obj]); } catch (e) { Logger.log('No se pudo enviar la invitación: ' + e); }
    }
    log_(existente ? 'AUMENTO_EDITADO' : 'AUMENTO_CREADO',
      cliente.NOMBRE + ' · +' + money_(monto) + ' del ' + isoADMY_(ini) + ' al ' + isoADMY_(fin));
    return obj.ID;
  });
}

function marcarCortado(token, id, nota) {
  exigirSesion_(token);
  return cerrarAumento_(id, ESTADO_CORTADO, nota);
}

function anularAumento(token, id, nota) {
  exigirSesion_(token);
  return cerrarAumento_(id, ESTADO_ANULADO, nota);
}

function reabrirAumento(token, id) {
  exigirSesion_(token);
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
    usarCalendar: p.USAR_CALENDAR === 'true',
    icsInvitacion: p.ICS_INVITACION === 'true',
    wspTelefono: p.WSP_TELEFONO || '',
    wspApiKey: p.WSP_APIKEY || '',
    teamsWebhook: p.TEAMS_WEBHOOK || ''
  };
}

function obtenerConfig(token) {
  exigirSesion_(token);
  const c = getConfig_();
  const trigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === TRIGGER_FN; });
  return {
    email: c.email,
    diasAviso: c.diasAviso,
    hora: c.hora,
    telegramConfigurado: !!c.telegramToken,
    telegramChatId: c.telegramChatId,
    usarCalendar: c.usarCalendar,
    icsInvitacion: c.icsInvitacion,
    wspTelefono: c.wspTelefono,
    wspConfigurado: !!(c.wspTelefono && c.wspApiKey),
    teamsConfigurado: !!c.teamsWebhook,
    triggerActivo: trigger,
    planilla: (function () { try { const ss = getSS_(); return { nombre: ss.getName(), url: ss.getUrl() }; } catch (e) { return null; } })(),
    cuotaCorreo: (function () { try { return MailApp.getRemainingDailyQuota(); } catch (e) { return null; } })(),
    zonaHoraria: tz_()
  };
}

function guardarConfig(token, cfg) {
  exigirSesion_(token);
  const props = PropertiesService.getScriptProperties();
  const emails = String(cfg.email || '').split(/[,;\s]+/).filter(String);
  emails.forEach(function (e) { if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) throw new Error('Correo inválido: ' + e); });
  props.setProperty('EMAIL_ALERTAS', emails.join(','));
  props.setProperty('DIAS_AVISO', String(Math.max(0, parseInt(cfg.diasAviso, 10) || 0)));
  const hora = Math.min(23, Math.max(0, parseInt(cfg.hora, 10) || 0));
  props.setProperty('HORA_ALERTA', String(hora));
  props.setProperty('USAR_CALENDAR', cfg.usarCalendar ? 'true' : 'false');
  props.setProperty('ICS_INVITACION', cfg.icsInvitacion ? 'true' : 'false');

  // WhatsApp (CallMeBot)
  if (cfg.borrarWsp) { props.deleteProperty('WSP_TELEFONO'); props.deleteProperty('WSP_APIKEY'); }
  else {
    if (cfg.wspTelefono !== undefined) {
      const tel = String(cfg.wspTelefono || '').replace(/[^\d+]/g, '');
      if (tel && !/^\+?\d{8,15}$/.test(tel)) throw new Error('Teléfono de WhatsApp inválido. Usa el formato internacional, ej: +56912345678');
      if (tel) props.setProperty('WSP_TELEFONO', tel.charAt(0) === '+' ? tel : '+' + tel); else props.deleteProperty('WSP_TELEFONO');
    }
    if (cfg.wspApiKey) props.setProperty('WSP_APIKEY', String(cfg.wspApiKey).trim());
  }

  // Microsoft Teams (webhook de Workflows)
  if (cfg.borrarTeams) props.deleteProperty('TEAMS_WEBHOOK');
  else if (cfg.teamsWebhook) {
    const url = String(cfg.teamsWebhook).trim();
    if (!/^https:\/\/[^\s]+$/.test(url)) throw new Error('La URL del webhook de Teams debe empezar con https://');
    props.setProperty('TEAMS_WEBHOOK', url);
  }

  if (cfg.telegramToken) props.setProperty('TELEGRAM_TOKEN', String(cfg.telegramToken).trim());
  if (cfg.borrarTelegram) { props.deleteProperty('TELEGRAM_TOKEN'); props.deleteProperty('TELEGRAM_CHAT_ID'); }
  if (cfg.telegramChatId !== undefined && !cfg.borrarTelegram) props.setProperty('TELEGRAM_CHAT_ID', String(cfg.telegramChatId).trim());
  instalarTrigger_(hora);
  log_('CONFIG', 'Alertas: ' + emails.join(',') + ' · aviso ' + cfg.diasAviso + ' días · hora ' + hora);
  return obtenerConfig(token);
}

// ════════════════════════════════════════════════════════════════
// REVISIÓN DIARIA Y ALERTAS
// ════════════════════════════════════════════════════════════════

/** Lo ejecuta el disparador diario (una vez cada 30 min como máximo). */
function revisarVencimientos() {
  // Es pública para que la llame el disparador diario; el freno evita que
  // alguien la invoque repetidamente desde fuera para generar correos.
  const cache = CacheService.getScriptCache();
  if (cache.get('REVISION_RECIENTE')) return { enviado: false, total: 0, omitido: true };
  cache.put('REVISION_RECIENTE', '1', 30 * 60);
  return revisarVencimientos_();
}

function revisarVencimientos_() {
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
function revisarAhora(token) {
  exigirSesion_(token);
  return revisarVencimientos_();
}

/** Botón "Enviar prueba" de la interfaz. */
function enviarPrueba(token, canal) {
  exigirSesion_(token);
  const cfg = getConfig_();
  const hoy = hoyISO_();
  const ejemplo = {
    id: 'PRUEBA', cliente: 'Cliente de prueba', banco: 'Banco', tarjeta: '1234', cupoBase: 5000,
    monto: 2000, inicio: hoy, fin: hoy, diasRestantes: 0, estadoCalc: 'VENCE_HOY', motivo: 'Mensaje de prueba'
  };
  if (canal === 'ics') {
    if (!cfg.email) throw new Error('Configura primero un correo.');
    enviarInvitacionIcs_(cfg, [{ ID: 'PRUEBA-' + Date.now(), CLIENTE: 'Cliente de prueba', MONTO: 2000, FECHA_INICIO: hoy, FECHA_FIN: hoy, MOTIVO: 'Prueba' }], true);
    return { canales: ['invitación de calendario'], errores: [] };
  }
  const r = enviarAlertas_({ vencidos: [], venceHoy: [ejemplo], porVencer: [], inicianHoy: [] }, hoy, cfg, true, canal);
  log_('PRUEBA_ALERTA', r.canales.join(', ') + (r.errores.length ? ' · errores: ' + r.errores.join(' | ') : ''));
  return r;
}

function enviarAlertas_(g, hoy, cfg, esPrueba, soloCanal) {
  const usar = function (c) { return !soloCanal || soloCanal === c; };
  const canales = [], errores = [];
  const urgentes = g.vencidos.length + g.venceHoy.length;
  const asunto = (esPrueba ? '[PRUEBA] ' : '') +
    (urgentes ? '⚠️ ' + urgentes + ' cupo(s) por cortar hoy' : '🔔 ' + g.porVencer.length + ' cupo(s) próximos a vencer') +
    ' · ' + isoADMY_(hoy);

  if (soloCanal === 'email' && !cfg.email) errores.push('Correo: no hay correo configurado.');
  if (soloCanal === 'whatsapp' && !(cfg.wspTelefono && cfg.wspApiKey)) errores.push('WhatsApp: falta teléfono o API key.');
  if (soloCanal === 'teams' && !cfg.teamsWebhook) errores.push('Teams: falta la URL del webhook.');
  if (soloCanal === 'telegram' && !(cfg.telegramToken && cfg.telegramChatId)) errores.push('Telegram: falta token o chat.');

  if (usar('email') && cfg.email) {
    try {
      MailApp.sendEmail({ to: cfg.email, subject: asunto, htmlBody: htmlAlerta_(g, hoy, cfg), body: textoAlerta_(g, hoy, false), name: APP_NAME });
      canales.push('email');
    } catch (e) { errores.push('Email: ' + e.message); }
  }
  if (usar('whatsapp') && cfg.wspTelefono && cfg.wspApiKey) {
    try {
      enviarWhatsApp_(cfg, (esPrueba ? '*[PRUEBA]*\n' : '') + textoAlerta_(g, hoy, 'wsp'));
      canales.push('whatsapp');
    } catch (e) { errores.push('WhatsApp: ' + e.message); }
  }
  if (usar('teams') && cfg.teamsWebhook) {
    try {
      enviarTeams_(cfg, g, hoy, esPrueba);
      canales.push('teams');
    } catch (e) { errores.push('Teams: ' + e.message); }
  }
  if (usar('telegram') && cfg.telegramToken && cfg.telegramChatId) {
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
    (a.cupoBase ? ' · cupo debe volver a ' + money_(a.cupoBase) : '');
}

function escHtml_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function textoAlerta_(g, hoy, modo) {
  const html = modo === true;
  const b = function (s) { return html ? '<b>' + escHtml_(s) + '</b>' : modo === 'wsp' ? '*' + s + '*' : s.toUpperCase(); };
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
  return APP_URL;
}

// ════════════════════════════════════════════════════════════════
// WHATSAPP (opcional, vía CallMeBot — gratis para uso personal)
// ════════════════════════════════════════════════════════════════
// Activación: agrega +34 644 51 95 23 a tus contactos y envíale por WhatsApp
// "I allow callmebot to send me messages". Te responde con tu API key.

function enviarWhatsApp_(cfg, texto) {
  const MAX = 3500; // se parte en varios mensajes si es muy largo
  const partes = [];
  for (let i = 0; i < texto.length; i += MAX) partes.push(texto.slice(i, i + MAX));
  partes.forEach(function (parte, i) {
    if (i) Utilities.sleep(2500);
    const url = 'https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(cfg.wspTelefono) +
      '&text=' + encodeURIComponent(parte) + '&apikey=' + encodeURIComponent(cfg.wspApiKey);
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const body = res.getContentText() || '';
    if (res.getResponseCode() >= 300 || /error|invalid|not valid|APIKey/i.test(body) && !/queued|sent/i.test(body)) {
      throw new Error(body.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200) || 'HTTP ' + res.getResponseCode());
    }
  });
}

// ════════════════════════════════════════════════════════════════
// MICROSOFT TEAMS (opcional, webhook de Workflows / Power Automate)
// ════════════════════════════════════════════════════════════════
// Teams → canal o chat → ••• → Workflows → "Send webhook alerts to a channel"
// (o "…to a chat") → copiar la URL que entrega.

function enviarTeams_(cfg, g, hoy, esPrueba) {
  const body = [{ type: 'TextBlock', size: 'Large', weight: 'Bolder', wrap: true,
    text: (esPrueba ? '[PRUEBA] ' : '') + '💳 Control de cupos · ' + isoADMY_(hoy) }];
  const seccion = function (titulo, color, lista, extra) {
    if (!lista.length) return;
    body.push({ type: 'TextBlock', text: titulo, weight: 'Bolder', color: color, spacing: 'Large', wrap: true });
    lista.forEach(function (a) {
      const tarjeta = [a.banco, a.tarjeta ? '****' + a.tarjeta : ''].filter(String).join(' ');
      const facts = [
        { title: 'Aumento', value: '+' + money_(a.monto) },
        { title: 'Vigencia', value: isoADMY_(a.inicio) + ' → ' + isoADMY_(a.fin) }
      ];
      if (a.cupoBase) facts.push({ title: 'Vuelve a', value: money_(a.cupoBase) });
      if (extra(a)) facts.push({ title: 'Estado', value: extra(a) });
      body.push({ type: 'Container', separator: true, items: [
        { type: 'TextBlock', text: a.cliente + (tarjeta ? ' · ' + tarjeta : ''), weight: 'Bolder', wrap: true },
        { type: 'FactSet', facts: facts }
      ] });
    });
  };
  seccion('🔴 Vencidos — corte pendiente', 'Attention', g.vencidos, function (a) { return 'venció hace ' + (-a.diasRestantes) + ' día(s)'; });
  seccion('🟠 Vencen hoy — hacer trámite con el banco', 'Warning', g.venceHoy, function () { return 'hoy'; });
  seccion('🟡 Próximos a vencer', 'Accent', g.porVencer, function (a) { return 'faltan ' + a.diasRestantes + ' día(s)'; });
  seccion('🟢 Inician hoy', 'Good', g.inicianHoy, function () { return ''; });
  const payload = {
    type: 'message',
    attachments: [{
      contentType: 'application/vnd.microsoft.card.adaptive',
      content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4',
        body: body,
        actions: [{ type: 'Action.OpenUrl', title: 'Abrir la app', url: APP_URL }]
      }
    }]
  };
  const res = UrlFetchApp.fetch(cfg.teamsWebhook, {
    method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  if (code < 200 || code >= 300) throw new Error('HTTP ' + code + ' ' + (res.getContentText() || '').slice(0, 200));
}

// ════════════════════════════════════════════════════════════════
// INVITACIÓN DE CALENDARIO (.ics) — Outlook / Teams / Google / Apple
// ════════════════════════════════════════════════════════════════
// Al crear un aumento se envía al correo de alertas un archivo .ics con un
// evento de día completo en la fecha de corte y dos recordatorios
// (09:00 del día anterior y 09:00 del mismo día). Al abrirlo, Outlook/Teams
// lo agrega al calendario. Reeditar un aumento reenvía el mismo evento (UID)
// con los datos nuevos.

function icsEsc_(s) { return String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); }

function icsFecha_(iso) { return iso.replace(/-/g, ''); }

function icsDiaSiguiente_(iso) {
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + 1));
  return d.toISOString().slice(0, 10).replace(/-/g, '');
}

function enviarInvitacionIcs_(cfg, aumentos, esPrueba) {
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const lineas = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Control de Cupos//ES', 'CALSCALE:GREGORIAN', 'METHOD:PUBLISH'];
  aumentos.forEach(function (a) {
    const fin = aISO_(a.FECHA_FIN), ini = aISO_(a.FECHA_INICIO);
    lineas.push(
      'BEGIN:VEVENT',
      'UID:' + a.ID + '@control-cupos',
      'DTSTAMP:' + stamp,
      'SEQUENCE:' + Math.floor(Date.now() / 1000),
      'DTSTART;VALUE=DATE:' + icsFecha_(fin),
      'DTEND;VALUE=DATE:' + icsDiaSiguiente_(fin),
      'SUMMARY:' + icsEsc_('⚠️ Cortar cupo: ' + a.CLIENTE + ' +' + money_(num_(a.MONTO))),
      'DESCRIPTION:' + icsEsc_('Aumento temporal de cupo que vence hoy.\nCliente: ' + a.CLIENTE + '\nAumento: +' + money_(num_(a.MONTO)) +
        '\nVigencia: ' + isoADMY_(ini) + ' al ' + isoADMY_(fin) + (a.MOTIVO ? '\nMotivo: ' + a.MOTIVO : '') + '\n\nApp: ' + APP_URL),
      'TRANSP:TRANSPARENT',
      'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Cortar cupo mañana', 'TRIGGER:-PT15H', 'END:VALARM',
      'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Cortar cupo hoy', 'TRIGGER:PT9H', 'END:VALARM',
      'END:VEVENT'
    );
  });
  lineas.push('END:VCALENDAR');
  // Líneas de máx. 75 octetos (RFC 5545) — plegado simple por caracteres
  const ics = lineas.map(function (l) {
    const out = []; while (l.length > 70) { out.push(l.slice(0, 70)); l = ' ' + l.slice(70); } out.push(l); return out.join('\r\n');
  }).join('\r\n');

  const uno = aumentos.length === 1 ? aumentos[0] : null;
  const asunto = (esPrueba ? '[PRUEBA] ' : '') + '📅 ' + (uno
    ? 'Corte de cupo ' + isoADMY_(aISO_(uno.FECHA_FIN)) + ': ' + uno.CLIENTE + ' +' + money_(num_(uno.MONTO))
    : aumentos.length + ' cortes de cupo para agregar al calendario');
  MailApp.sendEmail({
    to: cfg.email, subject: asunto, name: APP_NAME,
    body: 'Abre el archivo adjunto (corte-cupo.ics) para agregar ' + (uno ? 'el corte' : 'los cortes') + ' a tu calendario de Outlook / Teams / Google.',
    htmlBody: '<p>Abre el archivo adjunto <b>corte-cupo.ics</b> para agregar ' + (uno ? 'el corte' : 'los ' + aumentos.length + ' cortes') +
      ' a tu calendario (Outlook / Teams / Google). Incluye recordatorios a las 09:00 del día anterior y del mismo día.</p>',
    attachments: [Utilities.newBlob(ics, 'text/calendar', 'corte-cupo.ics')]
  });
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
function detectarChatTelegram(token) {
  exigirSesion_(token);
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
