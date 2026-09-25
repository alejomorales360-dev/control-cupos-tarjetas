// ════════════════════════════════════════════════════════════════
// Control de Cupos Temporales de Tarjetas — Codigo.gs
// ════════════════════════════════════════════════════════════════
//
// Registra clientes con su cupo base y los aumentos temporales de cupo
// (monto + rango de fechas). Todos los días revisa qué aumentos están por
// vencer, vencen hoy o ya vencieron sin haberse finalizado, y avisa SOLO POR
// CORREO (sin APIs externas):
//   - Correo corporativo (Outlook): resumen diario.
//   - Calendario (Outlook / Teams): invitación de reunión al crear cada
//     aumento; se cancela sola al marcarlo finalizado o anulado.
//   - Teams (opcional): el mismo resumen al correo de un canal de Teams.
//
// Todo se guarda en la planilla (ver HEADERS y crearBaseDeDatos()).
// ════════════════════════════════════════════════════════════════

const APP_NAME     = 'Control de Cupos';
const APP_URL      = 'https://alejomorales360-dev.github.io/control-cupos-tarjetas/';
const SH_CLIENTES  = 'CLIENTES';
const SH_AUMENTOS  = 'AUMENTOS';
const SH_HISTORIAL = 'HISTORIAL';
const SH_USUARIOS  = 'USUARIOS';
const SH_CONFIG    = 'CONFIG';
const SH_SESIONES  = 'SESIONES';

// Todo vive en la planilla. Lo único que queda en las Script Properties es
// SPREADSHEET_ID (hace falta para saber dónde está la planilla).
const HEADERS = {
  CLIENTES:  ['ID','NOMBRE','DOCUMENTO','BANCO','TARJETA_ULT4','CUPO_BASE','EMAIL','TELEFONO','NOTAS','ACTIVO','CREADO'],
  AUMENTOS:  ['ID','ID_CLIENTE','CLIENTE','MONTO','FECHA_INICIO','FECHA_FIN','MOTIVO','ESTADO','FECHA_CIERRE','NOTA_CIERRE','ULTIMO_AVISO','EVENTO_CALENDAR','CREADO'],
  HISTORIAL: ['FECHA','ACCION','DETALLE','USUARIO'],
  USUARIOS:  ['USUARIO','NOMBRE','ROL','ACTIVO','HASH','SAL','CLAVE_NUEVA','CREADO','ULTIMO_ACCESO'],
  CONFIG:    ['CLAVE','VALOR','DESCRIPCION'],
  SESIONES:  ['TOKEN_HASH','USUARIO','VENCE','CREADO']
};

// Claves de la hoja CONFIG (valor por defecto y descripción que se ve en la hoja)
const CONFIG_DEF = [
  ['EMAIL_ALERTAS',  '',   'Correo(s) corporativo(s) que reciben alertas e invitaciones de calendario, separados por coma'],
  ['DIAS_AVISO',     '3',  'Días de anticipación para el aviso previo'],
  ['HORA_ALERTA',    '8',  'Hora (0-23) de la revisión diaria'],
  ['ICS_INVITACION', 'SI', 'SI = enviar invitación de calendario (Outlook/Teams) al crear cada aumento y cancelarla al finalizarlo'],
  ['TEAMS_EMAIL',    '',   'Opcional: correo de un canal de Teams (canal → ••• → Obtener dirección de correo electrónico)']
];

// Ajustes de versiones anteriores que ya no se usan (se quitan de la hoja CONFIG)
const CONFIG_OBSOLETA = ['USAR_CALENDAR', 'WSP_TELEFONO', 'WSP_APIKEY', 'TEAMS_WEBHOOK', 'TELEGRAM_TOKEN', 'TELEGRAM_CHAT_ID'];

const ESTADO_ACTIVO  = 'ACTIVO';
const ESTADO_FINALIZADO = 'FINALIZADO';
const ESTADO_LEGADO_CORTADO = 'CORTADO';   // nombre usado por versiones anteriores

// Columnas renombradas: la app acepta planillas con los nombres antiguos
const ALIAS_COLUMNAS = { FECHA_CORTE: 'FECHA_CIERRE', NOTA_CORTE: 'NOTA_CIERRE' };
const ESTADO_ANULADO = 'ANULADO';

const TRIGGER_FN = 'revisarVencimientos';

// ════════════════════════════════════════════════════════════════
// INSTALACIÓN
// ════════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════════
// CREAR LA BASE DE DATOS  ← ejecutar desde el editor
// ════════════════════════════════════════════════════════════════
//
// Si el script NO está creado desde la planilla (Extensiones → Apps Script),
// pega aquí la URL de tu planilla antes de ejecutar crearBaseDeDatos():
const PLANILLA_URL = '';

// Descripción de cada columna (aparece como nota al pasar el mouse por el encabezado)
const NOTAS_COLUMNAS = {
  CLIENTES: {
    ID: 'Identificador interno (lo genera la app, no editar)', NOMBRE: 'Nombre del titular de la tarjeta',
    DOCUMENTO: 'Cédula / RUT / documento (opcional)', BANCO: 'Banco emisor', TARJETA_ULT4: 'Solo los últimos 4 dígitos. Nunca el número completo',
    CUPO_BASE: 'Cupo fijo en USD (al que vuelve cuando vence un aumento)', EMAIL: 'Correo del cliente (opcional)', TELEFONO: 'Teléfono del cliente (opcional)',
    NOTAS: 'Observaciones', ACTIVO: 'SI / NO', CREADO: 'Fecha de alta'
  },
  AUMENTOS: {
    ID: 'Identificador interno (lo genera la app, no editar)', ID_CLIENTE: 'ID del cliente en la hoja CLIENTES', CLIENTE: 'Nombre del cliente (copia)',
    MONTO: 'Monto del aumento temporal en USD', FECHA_INICIO: 'Desde (dd-mm-aaaa)', FECHA_FIN: 'Hasta: fecha de término, en que hay que finalizar el aumento con el banco',
    MOTIVO: 'Motivo o N° de solicitud', ESTADO: 'ACTIVO = pendiente de finalizar · FINALIZADO = ya se hizo el trámite con el banco · ANULADO = no se aplicó',
    FECHA_CIERRE: 'Cuándo se marcó como finalizado/anulado', NOTA_CIERRE: 'N° de trámite del banco u observación del cierre',
    ULTIMO_AVISO: 'Última vez que se envió alerta por este aumento', EVENTO_CALENDAR: 'ID del evento de Google Calendar (interno)', CREADO: 'Fecha de registro'
  },
  HISTORIAL: { FECHA: 'Fecha y hora', ACCION: 'Qué se hizo', DETALLE: 'Detalle de la acción', USUARIO: 'Usuario de la app que lo hizo' },
  USUARIOS: {
    USUARIO: 'Nombre de usuario para entrar (minúsculas)', NOMBRE: 'Nombre visible', ROL: 'ADMIN = todo · OPERADOR = solo clientes y aumentos',
    ACTIVO: 'SI / NO (NO bloquea el acceso y cierra sus sesiones)', HASH: 'Clave cifrada. NO editar', SAL: 'Parte del cifrado. NO editar',
    CLAVE_NUEVA: 'Para resetear una clave: escribe aquí la clave nueva. Se usa en el próximo ingreso y se borra sola',
    CREADO: 'Fecha de alta', ULTIMO_ACCESO: 'Último ingreso'
  },
  CONFIG: { CLAVE: 'Nombre del ajuste (no editar)', VALOR: 'Valor del ajuste (editable; SI/NO en las opciones)', DESCRIPCION: 'Para qué sirve' },
  SESIONES: { TOKEN_HASH: 'Sesión cifrada. Borrar la fila cierra esa sesión', USUARIO: 'Usuario de la sesión', VENCE: 'Vencimiento (milisegundos)', CREADO: 'Inicio de sesión' }
};

const ANCHOS_COLUMNAS = {
  CLIENTES:  [110, 220, 120, 130, 110, 110, 180, 120, 220, 70, 130],
  AUMENTOS:  [110, 110, 200, 100, 105, 105, 200, 95, 105, 200, 105, 120, 130],
  HISTORIAL: [140, 160, 420, 140],
  USUARIOS:  [140, 180, 100, 70, 200, 120, 160, 130, 130],
  CONFIG:    [170, 320, 460],
  SESIONES:  [220, 140, 130, 130]
};

const COLORES_PESTANA = { CLIENTES: '#2563eb', AUMENTOS: '#16a34a', HISTORIAL: '#6b7280', USUARIOS: '#9333ea', CONFIG: '#ea580c', SESIONES: '#9ca3af' };

/**
 * Crea (o completa) toda la base de datos en la planilla:
 *   - Hojas CLIENTES, AUMENTOS, HISTORIAL, USUARIOS, CONFIG y SESIONES con
 *     encabezados, notas explicativas, anchos, formatos y listas desplegables.
 *   - CONFIG con todos los ajustes y valores por defecto (correo = tu cuenta).
 *   - Usuario "admin" con clave aleatoria (si aún no hay usuarios): la clave
 *     aparece en el Registro de ejecución.
 *   - Disparador diario de alertas.
 * Es seguro ejecutarla varias veces: nunca borra datos, solo agrega lo que falte.
 */
function crearBaseDeDatos() {
  exigirDueno_();
  const ss = planillaDestino_();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  CACHE_ = {};

  ensureSheets_(ss);
  darFormatoHojas_(ss);   // también renombra encabezados antiguos (FECHA_CORTE → FECHA_CIERRE…)
  migrarEstadosFinalizado_(ss);
  ['Hoja 1', 'Hoja1', 'Sheet1'].forEach(function (n) {
    const h = ss.getSheetByName(n);
    if (h && h.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(h);
  });
  // Orden de pestañas
  Object.keys(HEADERS).forEach(function (n, i) { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
  ss.setActiveSheet(ss.getSheetByName(SH_CLIENTES));

  migrarPropiedades_();
  if (!leerConfigHoja_().EMAIL_ALERTAS) setConfigHoja_({ EMAIL_ALERTAS: Session.getEffectiveUser().getEmail() });
  instalarTrigger_(getConfig_().hora);
  log_('BASE_CREADA', ss.getName() + ' · ' + ss.getUrl());

  Logger.log('✅ Base de datos lista en: ' + ss.getName());
  Logger.log('   ' + ss.getUrl());
  Logger.log('   Hojas: ' + Object.keys(HEADERS).join(', '));
  Logger.log('   Alertas a: ' + leerConfigHoja_().EMAIL_ALERTAS + ' · revisión diaria a las ' + getConfig_().hora + ':00');
  if (!leerUsuarios_().length) generarClaveAcceso();
  else Logger.log('   Usuarios existentes: ' + leerUsuarios_().map(function (u) { return u.USUARIO; }).join(', ') + ' (tus claves no se tocaron)');
}

/** Cambia en la hoja AUMENTOS el estado antiguo CORTADO por FINALIZADO. */
function migrarEstadosFinalizado_(ss) {
  const sh = ss.getSheetByName(SH_AUMENTOS);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const col = HEADERS.AUMENTOS.indexOf('ESTADO') + 1;
  const r = sh.getRange(2, col, n, 1);
  const v = r.getValues();
  let cambios = 0;
  v.forEach(function (fila) { if (fila[0] === ESTADO_LEGADO_CORTADO) { fila[0] = ESTADO_FINALIZADO; cambios++; } });
  if (cambios) { r.setValues(v); Logger.log('   ' + cambios + ' aumento(s) CORTADO pasaron a FINALIZADO'); }
}

/** Alias: versiones anteriores usaban setup(). */
function setup() {
  crearBaseDeDatos();
}

/** Planilla donde crear la base: la del script, PLANILLA_URL, o la ya configurada. */
function planillaDestino_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss && PLANILLA_URL) {
    const m = String(PLANILLA_URL).match(/\/d\/([a-zA-Z0-9_-]{20,})/) || String(PLANILLA_URL).match(/^([a-zA-Z0-9_-]{20,})$/);
    if (!m) throw new Error('PLANILLA_URL no es una URL válida de Google Sheets.');
    ss = SpreadsheetApp.openById(m[1]);
  }
  if (!ss) {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) {} }
  }
  if (!ss) throw new Error('No sé en qué planilla crear la base. Pega la URL de tu planilla en la constante PLANILLA_URL (arriba de crearBaseDeDatos) y vuelve a ejecutar.');
  return ss;
}

function darFormatoHojas_(ss) {
  const lista = function (valores) {
    return SpreadsheetApp.newDataValidation().requireValueInList(valores, true).setAllowInvalid(false).build();
  };
  const FILAS = 1000;
  Object.keys(HEADERS).forEach(function (name) {
    const sh = ss.getSheetByName(name);
    const h = HEADERS[name];
    const head = sh.getRange(1, 1, 1, h.length);
    head.setValues([h]).setFontWeight('bold').setBackground('#1f2937').setFontColor('#ffffff').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setRowHeight(1, 28);
    head.setNotes([h.map(function (k) { return (NOTAS_COLUMNAS[name] || {})[k] || ''; })]);
    (ANCHOS_COLUMNAS[name] || []).forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    if (COLORES_PESTANA[name]) sh.setTabColor(COLORES_PESTANA[name]);

    const col = function (k) { return sh.getRange(2, h.indexOf(k) + 1, FILAS, 1); };
    h.forEach(function (k) {
      if (/^FECHA_|^ULTIMO_AVISO$/.test(k)) col(k).setNumberFormat('dd-mm-yyyy');
      else if (k === 'CREADO' || k === 'ULTIMO_ACCESO' || (name === SH_HISTORIAL && k === 'FECHA')) col(k).setNumberFormat('dd-mm-yyyy hh:mm');
      else if (k === 'MONTO' || k === 'CUPO_BASE') col(k).setNumberFormat('#,##0.00');
      else if (['TARJETA_ULT4', 'DOCUMENTO', 'TELEFONO', 'HASH', 'SAL', 'CLAVE_NUEVA', 'TOKEN_HASH', 'USUARIO', 'CLAVE', 'VALOR'].indexOf(k) >= 0) col(k).setNumberFormat('@');
    });
    if (name === SH_CLIENTES || name === SH_USUARIOS) col('ACTIVO').setDataValidation(lista(['SI', 'NO']));
    if (name === SH_AUMENTOS) col('ESTADO').setDataValidation(lista([ESTADO_ACTIVO, ESTADO_FINALIZADO, ESTADO_ANULADO]));
    if (name === SH_USUARIOS) {
      col('ROL').setDataValidation(lista(['ADMIN', 'OPERADOR']));
      sh.getRange(2, h.indexOf('HASH') + 1, FILAS, 2).setFontColor('#9ca3af');
    }
    if (name === SH_CONFIG) {
      sh.getRange(2, 1, FILAS, 1).setFontWeight('bold');
      sh.getRange(2, 3, FILAS, 1).setFontColor('#6b7280').setWrap(true);
      leer_(SH_CONFIG).forEach(function (r) {
        if (String(r.CLAVE) === 'ICS_INVITACION') sh.getRange(r._row, 2).setDataValidation(lista(['SI', 'NO']));
      });
    }
    if (name === SH_AUMENTOS) {
      // Resaltar vencidos sin finalizar (rojo) y que vencen hoy (naranjo)
      const rango = sh.getRange(2, 1, FILAS, h.length);
      const cFin = String.fromCharCode(65 + h.indexOf('FECHA_FIN')), cEst = String.fromCharCode(65 + h.indexOf('ESTADO'));
      sh.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + cEst + '2="ACTIVO",$' + cFin + '2<>"",$' + cFin + '2<TODAY())')
          .setBackground('#fee2e2').setRanges([rango]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($' + cEst + '2="ACTIVO",$' + cFin + '2=TODAY())')
          .setBackground('#ffedd5').setRanges([rango]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($' + cEst + '2="FINALIZADO",$' + cEst + '2="ANULADO")')
          .setFontColor('#9ca3af').setRanges([rango]).build()
      ]);
    }
  });
}

function getSS_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* se recrea abajo */ }
  }
  let activa = null;
  try { activa = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (activa) {
    props.setProperty('SPREADSHEET_ID', activa.getId());
    ensureSheets_(activa);
    return activa;
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
  exigirAdmin_(token);
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
  // Llevar usuarios, sesiones y configuración a la planilla nueva si ahí no hay
  const usuarios = leerUsuarios_(), config = leerConfigHoja_(), sesiones = leer_(SH_SESIONES);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  CACHE_ = {};
  if (!leerUsuarios_().length) agregarFilas_(SH_USUARIOS, usuarios.map(filaUsuario_));
  if (!leer_(SH_SESIONES).length) agregarFilas_(SH_SESIONES, sesiones);
  const cfgNueva = leerConfigHoja_(), faltan = {};
  Object.keys(config).forEach(function (k) { if (config[k] !== '' && cfgNueva[k] === '') faltan[k] = config[k]; });
  setConfigHoja_(faltan);
  log_('PLANILLA_CAMBIADA', ss.getName() + ' · ' + ss.getUrl());
  return { nombre: ss.getName(), url: ss.getUrl() };
}

function ensureSheets_(ss) {
  ensureHojas_(ss);
  const cfg = ss.getSheetByName(SH_CONFIG);
  cfg.getDataRange().getValues().map(function (r, i) { return CONFIG_OBSOLETA.indexOf(String(r[0])) >= 0 ? i + 1 : 0; })
    .filter(Number).reverse().forEach(function (fila) { cfg.deleteRow(fila); });
  const existentes = cfg.getDataRange().getValues().map(function (r) { return String(r[0]); });
  const nuevas = CONFIG_DEF.filter(function (d) { return existentes.indexOf(d[0]) < 0; });
  if (nuevas.length) {
    const r = cfg.getRange(cfg.getLastRow() + 1, 1, nuevas.length, 3);
    r.setNumberFormat('@').setValues(nuevas);
  }
}

function ensureHojas_(ss) {
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
    marcarCortado: marcarFinalizado,   // nombre antiguo, se mantiene por compatibilidad
    marcarFinalizado: marcarFinalizado,
    anularAumento: anularAumento,
    reabrirAumento: reabrirAumento,
    obtenerConfig: obtenerConfig,
    guardarConfig: guardarConfig,
    revisarAhora: revisarAhora,
    enviarPrueba: enviarPrueba,
    importarDatos: importarDatos,
    cambiarPlanilla: cambiarPlanilla,
    listarUsuarios: listarUsuarios,
    guardarUsuario: guardarUsuario
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

// ── Usuarios (hoja USUARIOS) ──
// La clave nunca se guarda en texto: HASH + SAL. Para resetear la clave de
// alguien desde la planilla, escribe la clave nueva en CLAVE_NUEVA: en su
// próximo login se convierte en hash y esa celda se borra.

let CACHE_ = {};
let USUARIO_ACTUAL_ = '';

function normUsuario_(u) { return String(u || '').trim().toLowerCase(); }

// Los hash en base64 pueden empezar con "+" o "=" y Sheets los tomaría como
// número o fórmula: en la hoja se guardan con prefijo "h:".
function filaUsuario_(u) {
  const f = {};
  Object.keys(u).forEach(function (k) { f[k] = u[k]; });
  if (f.HASH && String(f.HASH).indexOf('h:') !== 0) f.HASH = 'h:' + f.HASH;
  return f;
}

/** Hash del token de sesión, en hexadecimal (seguro para guardar en una celda). */
function tokenHash_(token) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

function leerUsuarios_() {
  if (!CACHE_.usuarios) {
    migrarPropiedades_();
    CACHE_.usuarios = leer_(SH_USUARIOS).map(function (u) {
      u.USUARIO = normUsuario_(u.USUARIO);
      u.HASH = String(u.HASH || '').replace(/^h:/, '');
      return u;
    });
  }
  return CACHE_.usuarios;
}

function buscarUsuario_(usuario) {
  const n = normUsuario_(usuario);
  return leerUsuarios_().filter(function (u) { return u.USUARIO === n; })[0] || null;
}

function esActivo_(u) { return u && String(u.ACTIVO).toUpperCase() !== 'NO' && u.ACTIVO !== false; }
function esAdmin_(u) { return u && String(u.ROL).toUpperCase() === 'ADMIN'; }

function guardarUsuarioFila_(u) {
  u.USUARIO = normUsuario_(u.USUARIO);
  escribirFila_(SH_USUARIOS, filaUsuario_(u), u._row || null);
  CACHE_.usuarios = null;
}

function setClave_(u, clave) {
  u.SAL = Utilities.getUuid();
  u.HASH = hashClave_(clave, u.SAL);
  u.CLAVE_NUEVA = '';
}

function validarUsuario_(usuario) {
  if (!/^[a-z0-9._@-]{3,60}$/.test(usuario)) throw new Error('Usuario inválido (3 a 60 caracteres: letras, números, . _ - @).');
}

function validarClave_(clave) {
  if (String(clave || '').length < 8) throw new Error('La clave debe tener al menos 8 caracteres.');
}

/** Pasa a la planilla lo que versiones anteriores guardaban en Script Properties. */
function migrarPropiedades_() {
  if (CACHE_.migrado) return;
  CACHE_.migrado = true;
  const props = PropertiesService.getScriptProperties();
  const p = props.getProperties();
  const claves = Object.keys(p).filter(function (k) { return k !== 'SPREADSHEET_ID'; });
  if (!claves.length) return;

  if (p.LOGIN_HASH && !leer_(SH_USUARIOS).length) {
    agregarFilas_(SH_USUARIOS, [filaUsuario_({
      USUARIO: normUsuario_(p.LOGIN_USUARIO || 'admin'), NOMBRE: 'Administrador', ROL: 'ADMIN', ACTIVO: 'SI',
      HASH: p.LOGIN_HASH, SAL: p.LOGIN_SAL, CLAVE_NUEVA: '', CREADO: new Date(), ULTIMO_ACCESO: ''
    })]);
  }
  // Las sesiones antiguas no se migran: basta con volver a entrar.

  const cfg = {};
  CONFIG_DEF.forEach(function (d) {
    if (p[d[0]] === undefined) return;
    cfg[d[0]] = d[0] === 'ICS_INVITACION' ? (p[d[0]] === 'true' ? 'SI' : 'NO') : p[d[0]];
  });
  setConfigHoja_(cfg, true);

  claves.forEach(function (k) { props.deleteProperty(k); });
  log_('MIGRACION', 'Login, sesiones y configuración movidos a la planilla');
}

/**
 * Ejecutar desde el editor para crear el usuario "admin" (si no hay usuarios)
 * o resetear la clave del primer ADMIN. La clave aparece en el Registro de
 * ejecución. Cierra las sesiones de ese usuario.
 */
function generarClaveAcceso() {
  exigirDueno_();
  getSS_();
  let u = leerUsuarios_().filter(esAdmin_)[0];
  if (!u) u = { USUARIO: 'admin', NOMBRE: 'Administrador', ROL: 'ADMIN', ACTIVO: 'SI', CREADO: new Date(), ULTIMO_ACCESO: '' };
  const clave = claveAleatoria_(12);
  setClave_(u, clave);
  u.ACTIVO = 'SI';
  guardarUsuarioFila_(u);
  borrarSesiones_(u.USUARIO);
  CacheService.getScriptCache().remove('LOGIN_FALLOS');
  log_('CLAVE_GENERADA', 'Clave de ' + u.USUARIO + ' regenerada desde el editor');
  Logger.log('Usuario: ' + u.USUARIO);
  Logger.log('Clave:   ' + clave);
  Logger.log('Entra a la app y cámbiala en Configuración → Mi cuenta.');
}

function iniciarSesion(usuario, clave, recordar) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('LOGIN_FALLOS') || 0);
  if (fallos >= MAX_FALLOS) throw new Error('Demasiados intentos fallidos. Espera ' + BLOQUEO_MIN + ' minutos.');
  if (!leerUsuarios_().length) throw new Error('Aún no hay usuarios. Ejecuta generarClaveAcceso() en el editor de Apps Script.');

  return conLock_(function () {
    const u = buscarUsuario_(usuario);
    let ok = false;
    if (u && esActivo_(u)) {
      if (u.CLAVE_NUEVA !== '' && u.CLAVE_NUEVA != null) {
        // Clave puesta a mano en la planilla: se acepta y se convierte en hash
        ok = String(clave || '') === String(u.CLAVE_NUEVA);
        if (ok) setClave_(u, String(clave));
      } else {
        ok = !!u.HASH && hashClave_(String(clave || ''), u.SAL) === u.HASH;
      }
    }
    if (!ok) {
      cache.put('LOGIN_FALLOS', String(fallos + 1), BLOQUEO_MIN * 60);
      Utilities.sleep(1000);
      log_('LOGIN_FALLIDO', 'Usuario: ' + String(usuario || '').slice(0, 60));
      throw new Error('Usuario o clave incorrectos.');
    }
    cache.remove('LOGIN_FALLOS');

    limpiarSesionesVencidas_();
    const token = Utilities.getUuid() + Utilities.getUuid();
    const vence = Date.now() + (recordar ? SESION_DIAS_LARGA * 86400000 : SESION_HORAS_CORTA * 3600000);
    agregarFilas_(SH_SESIONES, [{ TOKEN_HASH: tokenHash_(token), USUARIO: u.USUARIO, VENCE: vence, CREADO: new Date() }]);
    u.ULTIMO_ACCESO = new Date();
    guardarUsuarioFila_(u);
    USUARIO_ACTUAL_ = u.USUARIO;
    log_('LOGIN', recordar ? 'Sesión de ' + SESION_DIAS_LARGA + ' días' : 'Sesión de ' + SESION_HORAS_CORTA + ' horas');
    return { token: token, vence: vence, usuario: u.USUARIO, nombre: u.NOMBRE || u.USUARIO, rol: String(u.ROL).toUpperCase() };
  });
}

function cerrarSesion(token) {
  if (!token) return true;
  const h = tokenHash_(token);
  const s = leer_(SH_SESIONES).filter(function (x) { return x.TOKEN_HASH === h; })[0];
  if (s) getSheet_(SH_SESIONES).deleteRow(s._row);
  return true;
}

/** Valida el token y devuelve el usuario. Lanza SESION_INVALIDA si no sirve. */
function exigirSesion_(token) {
  if (!token) throw new Error('SESION_INVALIDA');
  const h = tokenHash_(token);
  const s = leer_(SH_SESIONES).filter(function (x) { return x.TOKEN_HASH === h; })[0];
  if (!s) throw new Error('SESION_INVALIDA');
  const u = buscarUsuario_(s.USUARIO);
  if (Number(s.VENCE) < Date.now() || !esActivo_(u)) {
    getSheet_(SH_SESIONES).deleteRow(s._row);
    throw new Error('SESION_INVALIDA');
  }
  USUARIO_ACTUAL_ = u.USUARIO;
  return u;
}

function exigirAdmin_(token) {
  const u = exigirSesion_(token);
  if (!esAdmin_(u)) throw new Error('Solo un administrador puede hacer esto.');
  return u;
}

function borrarFilas_(name, filtro) {
  const sh = getSheet_(name);
  leer_(name).filter(filtro).map(function (x) { return x._row; }).sort(function (a, b) { return b - a; })
    .forEach(function (r) { sh.deleteRow(r); });
}

function limpiarSesionesVencidas_() {
  const ahora = Date.now();
  borrarFilas_(SH_SESIONES, function (x) { return Number(x.VENCE) < ahora; });
}

function borrarSesiones_(usuario) {
  const n = normUsuario_(usuario);
  borrarFilas_(SH_SESIONES, function (x) { return !usuario || normUsuario_(x.USUARIO) === n; });
}

/** Mi cuenta: cambiar mi usuario y/o clave. Cierra mis sesiones. */
function cambiarCredenciales(token, actual, nuevoUsuario, nuevaClave) {
  const u = exigirSesion_(token);
  if (hashClave_(String(actual || ''), u.SAL) !== u.HASH) throw new Error('La clave actual no es correcta.');
  const anterior = u.USUARIO;
  const usuario = normUsuario_(nuevoUsuario) || anterior;
  validarUsuario_(usuario);
  if (usuario !== anterior && buscarUsuario_(usuario)) throw new Error('Ya existe un usuario "' + usuario + '".');
  validarClave_(nuevaClave);
  u.USUARIO = usuario;
  setClave_(u, String(nuevaClave));
  guardarUsuarioFila_(u);
  borrarSesiones_(anterior);
  log_('CLAVE_CAMBIADA', 'Usuario: ' + usuario + (usuario !== anterior ? ' (antes ' + anterior + ')' : ''));
  return true;
}

// ── Administración de usuarios (solo ADMIN) ──

function listarUsuarios(token) {
  exigirAdmin_(token);
  return leerUsuarios_().map(function (u) {
    return {
      usuario: u.USUARIO, nombre: u.NOMBRE || '', rol: String(u.ROL || 'OPERADOR').toUpperCase(), activo: esActivo_(u),
      ultimoAcceso: u.ULTIMO_ACCESO instanceof Date ? Utilities.formatDate(u.ULTIMO_ACCESO, tz_(), 'dd-MM-yyyy HH:mm') : '',
      claveTemporal: u.CLAVE_NUEVA !== '' && u.CLAVE_NUEVA != null
    };
  });
}

/** Crea (nuevo=true) o edita un usuario. clave es obligatoria al crear; al editar, opcional (resetea). */
function guardarUsuario(token, d) {
  const yo = exigirAdmin_(token);
  return conLock_(function () {
    const usuario = normUsuario_(d.usuario);
    validarUsuario_(usuario);
    const rol = String(d.rol || '').toUpperCase() === 'ADMIN' ? 'ADMIN' : 'OPERADOR';
    let u = buscarUsuario_(usuario);
    if (d.nuevo) {
      if (u) throw new Error('Ya existe un usuario "' + usuario + '".');
      validarClave_(d.clave);
      u = { USUARIO: usuario, CREADO: new Date(), ULTIMO_ACCESO: '' };
    } else if (!u) throw new Error('Usuario no encontrado.');

    const activo = d.activo !== false;
    if (u.USUARIO === yo.USUARIO && (!activo || rol !== 'ADMIN')) throw new Error('No puedes quitarte a ti mismo el rol de administrador ni desactivarte.');

    u.NOMBRE = String(d.nombre || '').trim();
    u.ROL = rol;
    u.ACTIVO = activo ? 'SI' : 'NO';
    if (d.clave) { validarClave_(d.clave); setClave_(u, String(d.clave)); }
    guardarUsuarioFila_(u);
    if (d.clave || !activo) borrarSesiones_(usuario);
    log_(d.nuevo ? 'USUARIO_CREADO' : 'USUARIO_EDITADO', usuario + ' · ' + rol + (activo ? '' : ' · inactivo') + (d.clave && !d.nuevo ? ' · clave reseteada' : ''));
    return listarUsuarios(token);
  });
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
      h.forEach(function (k, j) { o[ALIAS_COLUMNAS[k] || k] = row[j]; });
      if (name === SH_AUMENTOS && o.ESTADO === ESTADO_LEGADO_CORTADO) o.ESTADO = ESTADO_FINALIZADO;
      return o;
    })
    .filter(function (o) { return name === SH_HISTORIAL || o[h[0]] !== ''; });
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
    else if (k === 'CREADO' || k === 'ULTIMO_ACCESO' || (name === SH_HISTORIAL && k === 'FECHA')) c.setNumberFormat('dd-mm-yyyy hh:mm');
    else if (k === 'HASH' || k === 'SAL' || k === 'CLAVE_NUEVA' || k === 'TOKEN_HASH' || k === 'USUARIO') c.setNumberFormat('@');
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
    let user = USUARIO_ACTUAL_;
    if (!user) { try { user = Session.getActiveUser().getEmail(); } catch (e) {} }
    getSheet_(SH_HISTORIAL).appendRow([new Date(), accion, detalle, user || 'sistema']);
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
 *   VENCE_HOY   hoy es la fecha de fin → hacer el trámite con el banco
 *   VENCIDO     ya pasó la fecha de fin y NO se ha marcado como finalizado
 *   FINALIZADO / ANULADO  cerrados
 */
function estadoCalculado_(a, hoy, diasAviso) {
  if (a.ESTADO === ESTADO_FINALIZADO) return 'FINALIZADO';
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
    fechaCorte: aISO_(a.FECHA_CIERRE),
    notaCorte: a.NOTA_CIERRE || '',
    ultimoAviso: aISO_(a.ULTIMO_AVISO),
    tieneEvento: !!a.EVENTO_CALENDAR
  };
}

// ════════════════════════════════════════════════════════════════
// API PARA LA INTERFAZ (google.script.run)
// ════════════════════════════════════════════════════════════════

function obtenerDatos(token) {
  const yo = exigirSesion_(token);
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
    planillaUrl: esAdmin_(yo) ? getSS_().getUrl() : '',
    sesion: { usuario: yo.USUARIO, nombre: yo.NOMBRE || yo.USUARIO, rol: String(yo.ROL).toUpperCase() }
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
              ESTADO: ESTADO_ACTIVO, FECHA_CIERRE: '', NOTA_CIERRE: '', ULTIMO_AVISO: '', EVENTO_CALENDAR: '', CREADO: new Date(),
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

    agregarFilas_(SH_CLIENTES, nuevosClientes);
    actualizaciones.forEach(function (c) { escribirFila_(SH_CLIENTES, c, c._row); });
    agregarFilas_(SH_AUMENTOS, nuevosAumentos);
    const cfgImp = getConfig_();
    if (nuevosAumentos.length && cfgImp.icsInvitacion && cfgImp.email) {
      try {
        // Hasta 20: una invitación por aumento (entra sola al calendario).
        // Más: un solo correo con todos los eventos, para cuidar la cuota diaria.
        if (nuevosAumentos.length <= 20) nuevosAumentos.forEach(function (a) { enviarInvitacionIcs_(cfgImp, [a], false, 'REQUEST'); });
        else enviarInvitacionIcs_(cfgImp, nuevosAumentos, false, 'PUBLISH');
      } catch (e) { Logger.log('No se pudo enviar la invitación: ' + e); }
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
    else if (k === 'CREADO' || k === 'ULTIMO_ACCESO') r.setNumberFormat('dd-mm-yyyy hh:mm');
    else if (k === 'HASH' || k === 'SAL' || k === 'CLAVE_NUEVA' || k === 'TOKEN_HASH' || k === 'USUARIO') r.setNumberFormat('@');
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
      FECHA_CIERRE: '',
      NOTA_CIERRE: '',
      ULTIMO_AVISO: existente ? existente.ULTIMO_AVISO : '',
      EVENTO_CALENDAR: existente ? existente.EVENTO_CALENDAR : '',
      CREADO: existente ? existente.CREADO : new Date()
    };


    escribirFila_(SH_AUMENTOS, obj, existente ? existente._row : null);
    const cfgAum = getConfig_();
    if (cfgAum.icsInvitacion && cfgAum.email) {
      try { enviarInvitacionIcs_(cfgAum, [obj], false, 'REQUEST'); } catch (e) { Logger.log('No se pudo enviar la invitación: ' + e); }
    }
    log_(existente ? 'AUMENTO_EDITADO' : 'AUMENTO_CREADO',
      cliente.NOMBRE + ' · +' + money_(monto) + ' del ' + isoADMY_(ini) + ' al ' + isoADMY_(fin));
    return obj.ID;
  });
}

function marcarFinalizado(token, id, nota) {
  exigirSesion_(token);
  return cerrarAumento_(id, ESTADO_FINALIZADO, nota);
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
    a.FECHA_CIERRE = '';
    a.NOTA_CIERRE = '';
    escribirFila_(SH_AUMENTOS, a, a._row);
    const cfg = getConfig_();
    if (cfg.icsInvitacion && cfg.email) {
      try { enviarInvitacionIcs_(cfg, [a], false, 'REQUEST'); } catch (e) { Logger.log('No se pudo enviar la invitación: ' + e); }
    }
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
    a.FECHA_CIERRE = new Date();
    a.NOTA_CIERRE = String(nota || '').trim();
    if (a.EVENTO_CALENDAR) {
      if (estado === ESTADO_ANULADO) { borrarEvento_(a.EVENTO_CALENDAR); a.EVENTO_CALENDAR = ''; }
      else marcarEventoFinalizado_(a.EVENTO_CALENDAR);
    }
    escribirFila_(SH_AUMENTOS, a, a._row);
    const cfg = getConfig_();
    if (cfg.icsInvitacion && cfg.email) {
      // Quita el evento del calendario de Outlook/Teams
      try { enviarInvitacionIcs_(cfg, [a], false, 'CANCEL', estado); } catch (e) { Logger.log('No se pudo enviar la cancelación: ' + e); }
    }
    log_(estado === ESTADO_FINALIZADO ? 'AUMENTO_FINALIZADO' : 'AUMENTO_ANULADO',
      a.CLIENTE + ' · +' + money_(num_(a.MONTO)) + (a.NOTA_CIERRE ? ' · ' + a.NOTA_CIERRE : ''));
    return true;
  });
}

// ════════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ════════════════════════════════════════════════════════════════

/** Lee la hoja CONFIG como { CLAVE: 'valor' } (todo texto). */
function leerConfigHoja_() {
  if (!CACHE_.config) {
    if (!CACHE_.migrado) migrarPropiedades_();
    const out = {};
    CONFIG_DEF.forEach(function (d) { out[d[0]] = ''; });
    leer_(SH_CONFIG).forEach(function (r) { out[String(r.CLAVE).trim()] = String(r.VALOR == null ? '' : r.VALOR).trim(); });
    CACHE_.config = out;
  }
  return CACHE_.config;
}

/** Escribe claves en la hoja CONFIG (crea la fila si no existe). */
function setConfigHoja_(valores, sinMigrar) {
  const keys = Object.keys(valores || {});
  if (!keys.length) return;
  if (!sinMigrar && !CACHE_.migrado) migrarPropiedades_();
  const sh = getSheet_(SH_CONFIG);
  const filas = leer_(SH_CONFIG);
  keys.forEach(function (k) {
    const v = valores[k] == null ? '' : String(valores[k]);
    const f = filas.filter(function (r) { return String(r.CLAVE).trim() === k; })[0];
    if (f) sh.getRange(f._row, 2).setNumberFormat('@').setValue(v);
    else {
      const def = CONFIG_DEF.filter(function (d) { return d[0] === k; })[0];
      const r = sh.getLastRow() + 1;
      sh.getRange(r, 1, 1, 3).setNumberFormat('@').setValues([[k, v, def ? def[2] : '']]);
    }
  });
  CACHE_.config = null;
}

function siNo_(v) { return /^(si|sí|true|1|x|yes)$/i.test(String(v || '').trim()); }

function getConfig_() {
  const p = leerConfigHoja_();
  return {
    email: p.EMAIL_ALERTAS || '',
    diasAviso: Math.max(0, parseInt(p.DIAS_AVISO || '3', 10) || 0),
    hora: Math.min(23, Math.max(0, parseInt(p.HORA_ALERTA || '8', 10) || 0)),
    icsInvitacion: siNo_(p.ICS_INVITACION),
    teamsEmail: p.TEAMS_EMAIL || ''
  };
}

function obtenerConfig(token) {
  exigirAdmin_(token);
  const c = getConfig_();
  const trigger = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === TRIGGER_FN; });
  return {
    email: c.email,
    diasAviso: c.diasAviso,
    hora: c.hora,
    icsInvitacion: c.icsInvitacion,
    teamsEmail: c.teamsEmail,
    triggerActivo: trigger,
    planilla: (function () { try { const ss = getSS_(); return { nombre: ss.getName(), url: ss.getUrl() }; } catch (e) { return null; } })(),
    cuotaCorreo: (function () { try { return MailApp.getRemainingDailyQuota(); } catch (e) { return null; } })(),
    zonaHoraria: tz_()
  };
}

function guardarConfig(token, cfg) {
  exigirAdmin_(token);
  const v = {};
  const correoOk = function (e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e); };
  const emails = String(cfg.email || '').split(/[,;\s]+/).filter(String);
  emails.forEach(function (e) { if (!correoOk(e)) throw new Error('Correo inválido: ' + e); });
  v.EMAIL_ALERTAS = emails.join(',');
  v.DIAS_AVISO = String(Math.max(0, parseInt(cfg.diasAviso, 10) || 0));
  const hora = Math.min(23, Math.max(0, parseInt(cfg.hora, 10) || 0));
  v.HORA_ALERTA = String(hora);
  v.ICS_INVITACION = cfg.icsInvitacion ? 'SI' : 'NO';
  const teams = String(cfg.teamsEmail || '').trim();
  if (teams && !correoOk(teams)) throw new Error('Correo del canal de Teams inválido: ' + teams);
  v.TEAMS_EMAIL = teams;

  setConfigHoja_(v);
  instalarTrigger_(hora);
  log_('CONFIG', 'Alertas: ' + emails.join(',') + (teams ? ' · Teams: ' + teams : '') + ' · aviso ' + v.DIAS_AVISO + ' días · hora ' + hora);
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
  exigirAdmin_(token);
  const cfg = getConfig_();
  const hoy = hoyISO_();
  const ejemplo = {
    id: 'PRUEBA', cliente: 'Cliente de prueba', banco: 'Banco', tarjeta: '1234', cupoBase: 5000,
    monto: 2000, inicio: hoy, fin: hoy, diasRestantes: 0, estadoCalc: 'VENCE_HOY', motivo: 'Mensaje de prueba'
  };
  let r;
  if (canal === 'ics') {
    if (!cfg.email) throw new Error('Configura primero el correo corporativo.');
    enviarInvitacionIcs_(cfg, [{ ID: 'PRUEBA-' + Date.now(), CLIENTE: 'Cliente de prueba', MONTO: 2000, FECHA_INICIO: hoy, FECHA_FIN: hoy, MOTIVO: 'Prueba' }], true, 'REQUEST');
    r = { canales: ['invitación de calendario'], errores: [] };
  } else {
    r = enviarAlertas_({ vencidos: [], venceHoy: [ejemplo], porVencer: [], inicianHoy: [] }, hoy, cfg, true, canal);
  }
  log_('PRUEBA_ALERTA', r.canales.join(', ') + (r.errores.length ? ' · errores: ' + r.errores.join(' | ') : ''));
  return r;
}

/** Envía el resumen al correo corporativo y (si está configurado) al canal de Teams. */
function enviarAlertas_(g, hoy, cfg, esPrueba, soloCanal) {
  const usar = function (c) { return !soloCanal || soloCanal === c; };
  const canales = [], errores = [];
  const urgentes = g.vencidos.length + g.venceHoy.length;
  const asunto = (esPrueba ? '[PRUEBA] ' : '') +
    'Control de Cupos: ' + (urgentes ? urgentes + ' aumento(s) por finalizar' : g.porVencer.length + ' aumento(s) proximos a vencer') +
    ' (' + isoADMY_(hoy) + ')';
  const correo = { subject: asunto, htmlBody: htmlAlerta_(g, hoy, cfg), body: textoAlerta_(g, hoy), name: APP_NAME };

  if (soloCanal === 'email' && !cfg.email) errores.push('Correo: no hay correo configurado.');
  if (soloCanal === 'teams' && !cfg.teamsEmail) errores.push('Teams: falta el correo del canal.');

  if (usar('email') && cfg.email) {
    try { MailApp.sendEmail(Object.assign({ to: cfg.email }, correo)); canales.push('correo'); }
    catch (e) { errores.push('Correo: ' + e.message); }
  }
  if (usar('teams') && cfg.teamsEmail) {
    try { MailApp.sendEmail(Object.assign({ to: cfg.teamsEmail }, correo)); canales.push('Teams'); }
    catch (e) { errores.push('Teams: ' + e.message); }
  }
  if (!canales.length && !errores.length) errores.push('No hay ningún correo de alerta configurado.');
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

// Correos deliberadamente simples (sin emojis, botones ni enlaces): los filtros
// anti-phishing corporativos suelen desviar correos externos con enlaces.
function textoAlerta_(g, hoy) {
  const out = ['CONTROL DE CUPOS - ' + isoADMY_(hoy), ''];
  const bloque = function (titulo, lista, extra) {
    if (!lista.length) return;
    out.push(titulo);
    lista.forEach(function (a) { out.push('- ' + lineaAumento_(a) + (extra(a) ? ' - ' + extra(a) : '')); });
    out.push('');
  };
  bloque('VENCIDOS - SIN FINALIZAR', g.vencidos, function (a) { return 'vencio hace ' + (-a.diasRestantes) + ' dia(s)'; });
  bloque('VENCEN HOY - hacer tramite con el banco', g.venceHoy, function () { return ''; });
  bloque('PROXIMOS A VENCER', g.porVencer, function (a) { return 'faltan ' + a.diasRestantes + ' dia(s)'; });
  bloque('INICIAN HOY', g.inicianHoy, function () { return ''; });
  out.push('Cuando hagas el tramite con el banco, marcalo como "Finalizado" en la app para dejar de recibir este aviso.');
  return out.join('\n');
}

function htmlAlerta_(g, hoy, cfg) {
  const td = 'padding:6px 8px;border:1px solid #ddd';
  const seccion = function (titulo, lista, extra) {
    if (!lista.length) return '';
    const filas = lista.map(function (a) {
      const tarjeta = [a.banco, a.tarjeta ? '****' + a.tarjeta : ''].filter(String).join(' ');
      return '<tr><td style="' + td + '"><b>' + escHtml_(a.cliente) + '</b>' + (tarjeta ? '<br>' + escHtml_(tarjeta) : '') + '</td>' +
        '<td style="' + td + ';text-align:right">+' + money_(a.monto) + '</td>' +
        '<td style="' + td + '">' + isoADMY_(a.inicio) + ' al <b>' + isoADMY_(a.fin) + '</b></td>' +
        '<td style="' + td + ';text-align:right">' + (a.cupoBase ? money_(a.cupoBase) : '-') + '</td>' +
        '<td style="' + td + '">' + escHtml_(extra(a)) + '</td></tr>';
    }).join('');
    return '<p style="margin:18px 0 6px"><b>' + titulo + '</b></p>' +
      '<table style="border-collapse:collapse;font-size:14px">' +
      '<tr style="background:#f2f2f2"><th style="' + td + '">Cliente</th><th style="' + td + '">Aumento</th><th style="' + td + '">Vigencia</th><th style="' + td + '">Cupo base</th><th style="' + td + '"></th></tr>' +
      filas + '</table>';
  };
  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#111">' +
    '<p><b>Control de Cupos - ' + isoADMY_(hoy) + '</b></p>' +
    seccion('Vencidos - sin finalizar', g.vencidos, function (a) { return 'vencio hace ' + (-a.diasRestantes) + ' dia(s)'; }) +
    seccion('Vencen hoy - hacer tramite con el banco', g.venceHoy, function () { return 'hoy'; }) +
    seccion('Proximos a vencer (' + cfg.diasAviso + ' dias)', g.porVencer, function (a) { return 'faltan ' + a.diasRestantes + ' dia(s)'; }) +
    seccion('Inician hoy', g.inicianHoy, function () { return ''; }) +
    '<p style="margin-top:18px">Cuando hagas el tramite con el banco, marcalo como <b>Finalizado</b> en la app para dejar de recibir este aviso.</p>' +
    '</div>';
}

// ════════════════════════════════════════════════════════════════
// INVITACIÓN DE CALENDARIO — Outlook / Teams (sin API, por correo)
// ════════════════════════════════════════════════════════════════
// REQUEST: al crear/editar/reabrir un aumento llega una invitación de reunión
//   (día completo, en la fecha de término) al correo corporativo; Outlook la pone
//   en el calendario, que es el mismo que muestra Teams. Recordatorios:
//   09:00 del día anterior y 09:00 del mismo día.
// CANCEL: al marcarlo finalizado o anulado llega la cancelación y el evento se
//   quita del calendario. Mismo UID para todo el ciclo de vida del aumento.
// PUBLISH: importación masiva (>20): un solo .ics con todos los eventos.

function icsEsc_(s) { return String(s || '').replace(/\\/g, '\\\\').replace(/;/g, '\\;').replace(/,/g, '\\,').replace(/\r?\n/g, '\\n'); }

function icsFecha_(iso) { return iso.replace(/-/g, ''); }

function icsDiaSiguiente_(iso) {
  const p = iso.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + 1));
  return d.toISOString().slice(0, 10).replace(/-/g, '');
}

function enviarInvitacionIcs_(cfg, aumentos, esPrueba, metodo, estadoCierre) {
  metodo = metodo || 'REQUEST';
  const stamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const organizador = Session.getEffectiveUser().getEmail();
  const asistentes = String(cfg.email || '').split(',').filter(String);
  const lineas = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Control de Cupos//ES', 'CALSCALE:GREGORIAN', 'METHOD:' + metodo];
  aumentos.forEach(function (a) {
    const fin = aISO_(a.FECHA_FIN), ini = aISO_(a.FECHA_INICIO);
    lineas.push(
      'BEGIN:VEVENT',
      'UID:' + a.ID + '@control-cupos',
      'DTSTAMP:' + stamp,
      'SEQUENCE:' + Math.floor(Date.now() / 1000),
      'DTSTART;VALUE=DATE:' + icsFecha_(fin),
      'DTEND;VALUE=DATE:' + icsDiaSiguiente_(fin),
      'SUMMARY:' + icsEsc_('Finalizar aumento: ' + a.CLIENTE + ' +' + money_(num_(a.MONTO))),
      'DESCRIPTION:' + icsEsc_('Aumento temporal de cupo que vence hoy: hacer el trámite con el banco para finalizarlo.\nCliente: ' + a.CLIENTE +
        '\nAumento: +' + money_(num_(a.MONTO)) + '\nVigencia: ' + isoADMY_(ini) + ' al ' + isoADMY_(fin) +
        (a.MOTIVO ? '\nMotivo: ' + a.MOTIVO : '') + '\n\nCuando lo finalices, marcalo en la app y este evento se quitara solo.'),
      'TRANSP:TRANSPARENT',
      'STATUS:' + (metodo === 'CANCEL' ? 'CANCELLED' : 'CONFIRMED')
    );
    if (metodo !== 'PUBLISH') {
      lineas.push('ORGANIZER;CN=' + icsEsc_(APP_NAME) + ':mailto:' + organizador);
      asistentes.forEach(function (e) { lineas.push('ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=FALSE:mailto:' + e); });
    }
    if (metodo !== 'CANCEL') {
      lineas.push(
        'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Finalizar aumento mañana', 'TRIGGER:-PT15H', 'END:VALARM',
        'BEGIN:VALARM', 'ACTION:DISPLAY', 'DESCRIPTION:Finalizar aumento hoy', 'TRIGGER:PT9H', 'END:VALARM'
      );
    }
    lineas.push('END:VEVENT');
  });
  lineas.push('END:VCALENDAR');
  // Líneas de máx. 75 octetos (RFC 5545) — plegado simple por caracteres
  const ics = lineas.map(function (l) {
    const out = []; while (l.length > 70) { out.push(l.slice(0, 70)); l = ' ' + l.slice(70); } out.push(l); return out.join('\r\n');
  }).join('\r\n');

  const uno = aumentos.length === 1 ? aumentos[0] : null;
  const detalle = uno ? uno.CLIENTE + ' +' + money_(num_(uno.MONTO)) + ' - termina ' + isoADMY_(aISO_(uno.FECHA_FIN)) : '';
  let asunto, html;
  if (metodo === 'CANCEL') {
    asunto = (estadoCierre === ESTADO_ANULADO ? 'Anulado: ' : 'Aumento finalizado: ') + detalle;
    html = '<p>' + (estadoCierre === ESTADO_ANULADO ? 'El aumento fue anulado' : 'El aumento ya fue finalizado') +
      '. Esta cancelación quita el evento de tu calendario.</p>';
  } else if (metodo === 'PUBLISH') {
    asunto = aumentos.length + ' aumentos para agregar al calendario';
    html = '<p>Abre el adjunto <b>aumentos-cupo.ics</b> para agregar las ' + aumentos.length + ' fechas de término a tu calendario de Outlook / Teams.</p>';
  } else {
    asunto = 'Termino de aumento: ' + detalle;
    html = '<p>Recordatorio en tu calendario para el <b>' + (uno ? isoADMY_(aISO_(uno.FECHA_FIN)) : '') + '</b>: finalizar el aumento de cupo de <b>' +
      escHtml_(uno ? uno.CLIENTE : '') + '</b> (+' + (uno ? money_(num_(uno.MONTO)) : '') + ').</p>' +
      '<p>Avisos a las 09:00 del día anterior y del mismo día. Al marcarlo como finalizado en la app, el evento se quita solo.</p>';
  }
  MailApp.sendEmail({
    to: cfg.email, subject: (esPrueba ? '[PRUEBA] ' : '') + asunto, name: APP_NAME,
    body: html.replace(/<[^>]+>/g, ''), htmlBody: html,
    attachments: [Utilities.newBlob(ics, 'text/calendar; charset=UTF-8; method=' + metodo, metodo === 'PUBLISH' ? 'aumentos-cupo.ics' : 'invite.ics')]
  });
}

// ════════════════════════════════════════════════════════════════
// GOOGLE CALENDAR (solo para eventos creados por versiones anteriores)
// ════════════════════════════════════════════════════════════════

function borrarEvento_(eventId) {
  try {
    const ev = CalendarApp.getDefaultCalendar().getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (e) { Logger.log('No se pudo borrar el evento: ' + e); }
}

function marcarEventoFinalizado_(eventId) {
  try {
    const ev = CalendarApp.getDefaultCalendar().getEventById(eventId);
    if (ev) {
      ev.setTitle(ev.getTitle().replace(/^⚠️ Cortar cupo|^Finalizar aumento/, '✅ Aumento finalizado'));
      ev.removeAllReminders();
    }
  } catch (e) { Logger.log('No se pudo actualizar el evento: ' + e); }
}
