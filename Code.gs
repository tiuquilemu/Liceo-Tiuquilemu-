// =====================================================================
// SISTEMA DE ASISTENCIA QR — LICEO TIUQUILEMU
// Backend en Google Apps Script. Esta es la ÚNICA base de datos real:
// todos los computadores que abran la app leen y escriben aquí mismo.
//
// ==== CONFIGURACIÓN (se hace una sola vez) ====
// 1. Ve a sheets.google.com y crea una hoja nueva llamada "Asistencia Liceo Tiuquilemu".
// 2. Ve a Extensiones > Apps Script.
// 3. Borra el código de ejemplo y pega TODO este archivo.
// 4. Haz clic en Implementar > Nueva implementación > tipo "Aplicación web".
//    - Ejecutar como: tu cuenta
//    - Quién tiene acceso: Cualquier usuario
// 5. Autoriza los permisos que pida Google (son de tu propia cuenta).
// 6. Copia la URL entregada ("URL de la aplicación web"): esa es la URL que
//    vas a pegar en la app la primera vez que la abras en cada computador.
// =====================================================================

const SHEET_CONFIG = 'Config';
const SHEET_ALUMNOS = 'Alumnos';
const SHEET_ASISTENCIA = 'Asistencia';
const TIMEZONE_FALLBACK = 'America/Santiago';

const ALUMNOS_COLS = ['rut','nombre','curso','apoderadoNombre','apoderadoTelefono','apoderadoEmail','callmebotApiKey'];
const ASISTENCIA_COLS = ['id','rut','nombre','curso','fecha','hora','timestamp','canal'];
const CONFIG_DEFAULTS = {
  schoolName: 'Liceo Tiuquilemu',
  adminPasswordHash: '',
  userPasswordHash: '',
  correosInforme: 'inspectoria@liceotiuquilemu.cl,informatica@liceotiuquilemu.cl',
  emailDireccion: 'informatica@liceotiuquilemu.cl',
  emailInspectoria: 'inspectoria@liceotiuquilemu.cl',
  horaEnvio: '12:00',
  horaRevisionInasistencia: '12:00',
  ultimaRevisionInasistencia: ''
};

function getSheet_(name, headers){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if(!sheet){
    sheet = ss.insertSheet(name);
    if(headers) sheet.appendRow(headers);
  }
  return sheet;
}

function jsonResponse_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function timezone_(){
  try{
    return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || TIMEZONE_FALLBACK;
  }catch(err){
    return Session.getScriptTimeZone() || TIMEZONE_FALLBACK;
  }
}

function today_(){
  return Utilities.formatDate(new Date(), timezone_(), 'dd/MM/yyyy');
}

function normalizeRut_(value){
  return String(value == null ? '' : value)
    .toUpperCase()
    .replace(/[^0-9K]/g, '');
}

function normalizeDate_(value){
  if(value instanceof Date && !isNaN(value.getTime())){
    return Utilities.formatDate(value, timezone_(), 'dd/MM/yyyy');
  }
  if(typeof value === 'number' && isFinite(value)){
    const millis = Date.UTC(1899, 11, 30) + Math.round(value * 86400000);
    return Utilities.formatDate(new Date(millis), 'UTC', 'dd/MM/yyyy');
  }

  const text = String(value == null ? '' : value).trim();
  if(!text) return '';
  let match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if(match){
    return ('0' + match[1]).slice(-2) + '/' +
      ('0' + match[2]).slice(-2) + '/' + match[3];
  }
  match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$/);
  if(match){
    return ('0' + match[3]).slice(-2) + '/' +
      ('0' + match[2]).slice(-2) + '/' + match[1];
  }
  const parsed = new Date(text);
  return isNaN(parsed.getTime())
    ? text
    : Utilities.formatDate(parsed, timezone_(), 'dd/MM/yyyy');
}

function normalizeTime_(value){
  if(value instanceof Date && !isNaN(value.getTime())){
    return Utilities.formatDate(value, timezone_(), 'HH:mm');
  }
  const original = String(value == null ? '' : value).trim();
  if(!original) return '';
  const text = original.toLowerCase().replace(/\./g, '').replace(/\s+/g, '');
  const match = text.match(/^(\d{1,2}):(\d{2})(?::\d{2})?([ap]m)?$/);
  if(!match) return original;
  let hour = parseInt(match[1], 10);
  const minute = parseInt(match[2], 10);
  if(match[3] === 'pm' && hour < 12) hour += 12;
  if(match[3] === 'am' && hour === 12) hour = 0;
  return ('0' + hour).slice(-2) + ':' + ('0' + minute).slice(-2);
}

function triggerHour_(value, fallback){
  if(value instanceof Date && !isNaN(value.getTime())){
    return parseInt(Utilities.formatDate(value, timezone_(), 'H'), 10);
  }
  if(typeof value === 'number' && isFinite(value)){
    return Math.floor((((value % 1) + 1) % 1) * 24);
  }
  const normalized = normalizeTime_(value);
  const match = normalized.match(/^(\d{2}):\d{2}$/);
  const hour = match ? parseInt(match[1], 10) : NaN;
  return isNaN(hour) || hour < 0 || hour > 23 ? fallback : hour;
}

function sharedLock_(){
  return LockService.getDocumentLock() || LockService.getScriptLock();
}

// ---------------- Config ----------------
function readConfig_(){
  const sheet = getSheet_(SHEET_CONFIG, ['clave','valor']);
  const data = sheet.getDataRange().getDisplayValues();
  const cfg = Object.assign({}, CONFIG_DEFAULTS);
  for(let i = 1; i < data.length; i++){
    if(data[i][0]) cfg[data[i][0]] = data[i][1];
  }
  return cfg;
}

function setConfigValue_(key, value){
  const sheet = getSheet_(SHEET_CONFIG, ['clave','valor']);
  const lastRow = Math.max(1, sheet.getLastRow());
  const keys = sheet.getRange(1, 1, lastRow, 1).getDisplayValues()
    .map(row => String(row[0] || '').trim());
  let row = keys.indexOf(key) + 1;
  if(row < 2) row = Math.max(2, sheet.getLastRow() + 1);
  sheet.getRange(row, 1).setValue(key);
  const valueCell = sheet.getRange(row, 2);
  valueCell.setNumberFormat('@');
  valueCell.setValue(String(value == null ? '' : value));
}

function writeConfig_(partial){
  partial = partial || {};
  Object.keys(partial).forEach(key => setConfigValue_(key, partial[key]));
  if(Object.prototype.hasOwnProperty.call(partial, 'horaEnvio')){
    asegurarTriggerDiario_(partial.horaEnvio);
  }
  if(Object.prototype.hasOwnProperty.call(partial, 'horaRevisionInasistencia')){
    asegurarTriggerInasistencia_(partial.horaRevisionInasistencia);
  }
  return readConfig_();
}

// ---------------- Alumnos ----------------
function readStudents_(){
  const sheet = getSheet_(SHEET_ALUMNOS, ALUMNOS_COLS);
  const data = sheet.getDataRange().getValues();
  const out = [];
  for(let i = 1; i < data.length; i++){
    if(!data[i][0]) continue;
    const row = {};
    ALUMNOS_COLS.forEach((c, idx) => row[c] = data[i][idx] || '');
    out.push(row);
  }
  return out;
}

function findStudentRow_(sheet, rut){
  const data = sheet.getDataRange().getValues();
  for(let i = 1; i < data.length; i++){
    if(normalizeRut_(data[i][0]) === normalizeRut_(rut)) return i + 1; // fila real (1-indexada)
  }
  return -1;
}

function addStudent_(student){
  const sheet = getSheet_(SHEET_ALUMNOS, ALUMNOS_COLS);
  const existingRow = findStudentRow_(sheet, student.rut);
  if(existingRow > 0) throw new Error('Ya existe un alumno con ese RUT');
  sheet.appendRow(ALUMNOS_COLS.map(c => student[c] || ''));
}

function updateStudent_(student){
  const sheet = getSheet_(SHEET_ALUMNOS, ALUMNOS_COLS);
  const row = findStudentRow_(sheet, student.rut);
  if(row < 0) throw new Error('Alumno no encontrado');
  sheet.getRange(row, 1, 1, ALUMNOS_COLS.length).setValues([ALUMNOS_COLS.map(c => student[c] || '')]);
}

function deleteStudent_(rut){
  const sheet = getSheet_(SHEET_ALUMNOS, ALUMNOS_COLS);
  const row = findStudentRow_(sheet, rut);
  if(row > 0) sheet.deleteRow(row);
}

// ---------------- Asistencia ----------------
function readAttendance_(){
  const sheet = getSheet_(SHEET_ASISTENCIA, ASISTENCIA_COLS);
  const data = sheet.getDataRange().getValues();
  const out = [];
  for(let i = 1; i < data.length; i++){
    if(!data[i][0]) continue;
    const row = {};
    ASISTENCIA_COLS.forEach((c, idx) => {
      let value = data[i][idx];
      if(c === 'fecha') value = normalizeDate_(value);
      if(c === 'hora') value = normalizeTime_(value);
      row[c] = value == null ? '' : value;
    });
    out.push(row);
  }
  return out.reverse(); // más reciente primero
}

function addAttendance_(record){
  const lock = sharedLock_();
  lock.waitLock(30000);
  try{
    const sheet = getSheet_(SHEET_ASISTENCIA, ASISTENCIA_COLS);
    const normalized = Object.assign({}, record, {
      fecha: normalizeDate_(record.fecha) || today_(),
      hora: normalizeTime_(record.hora)
    });
    const row = Math.max(2, sheet.getLastRow() + 1);
    // Formatear ANTES de escribir evita que Sheets convierta 02/09/2026
    // en un número de fecha y 07:58 en una fracción de día.
    sheet.getRange(row, ASISTENCIA_COLS.indexOf('fecha') + 1).setNumberFormat('@');
    sheet.getRange(row, ASISTENCIA_COLS.indexOf('hora') + 1).setNumberFormat('@');
    sheet.getRange(row, 1, 1, ASISTENCIA_COLS.length).setValues([
      ASISTENCIA_COLS.map(c => normalized[c] !== undefined ? normalized[c] : '')
    ]);
  }finally{
    lock.releaseLock();
  }
}

// ---------------- Endpoints ----------------
function doGet(e){
  const out = {
    config: readConfig_(),
    students: readStudents_(),
    attendance: readAttendance_()
  };
  return jsonResponse_(out);
}

function doPost(e){
  let body;
  try{
    body = JSON.parse(e.postData.contents);
  }catch(err){
    return jsonResponse_({ ok:false, error:'No se pudo leer la solicitud' });
  }

  try{
    switch(body.type){
      case 'save_config':
        return jsonResponse_({ ok:true, config: writeConfig_(body.config || {}) });

      case 'add_student':
        addStudent_(body.student);
        return jsonResponse_({ ok:true });

      case 'update_student':
        updateStudent_(body.student);
        return jsonResponse_({ ok:true });

      case 'delete_student':
        deleteStudent_(body.rut);
        return jsonResponse_({ ok:true });

      case 'add_attendance':
        addAttendance_(body.record);
        return jsonResponse_({ ok:true });

      case 'notify_guardian':
        if(body.email){
          MailApp.sendEmail({
            to: body.email,
            subject: 'Registro de asistencia - ' + (body.nombre || ''),
            body: body.mensaje || ''
          });
        }
        return jsonResponse_({ ok:true });

      case 'notify_whatsapp':
        return jsonResponse_(notifyWhatsApp_(body.phone, body.apikey, body.mensaje));

      case 'send_absence_check_now':
        return jsonResponse_(Object.assign(
          { ok:true },
          revisarInasistencias_({ origen:'manual' })
        ));

      case 'send_report_now':
        return jsonResponse_(Object.assign({ ok:true }, enviarInformeDiario_()));

      default:
        return jsonResponse_({ ok:false, error:'Tipo de solicitud no reconocido' });
    }
  }catch(err){
    return jsonResponse_({ ok:false, error: String(err) });
  }
}

// ---------------- Triggers automáticos ----------------
function ensureDailyTrigger_(handlerName, configuredTime, fallbackHour){
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if(trigger.getHandlerFunction() === handlerName) ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger(handlerName)
    .timeBased()
    .everyDays(1)
    .atHour(triggerHour_(configuredTime, fallbackHour))
    .nearMinute(0)
    .create();
}

function asegurarTriggerDiario_(horaEnvio){
  ensureDailyTrigger_('enviarInformeDiario_', horaEnvio, 18);
}

function asegurarTriggerInasistencia_(horaRevision){
  ensureDailyTrigger_('revisarInasistenciasProgramado_', horaRevision, 13);
}

function instalarTriggers(){
  const cfg = readConfig_();
  asegurarTriggerDiario_(cfg.horaEnvio);
  asegurarTriggerInasistencia_(cfg.horaRevisionInasistencia);
}

function revisarInasistenciasProgramado_(){
  return revisarInasistencias_({ origen:'automatico' });
}

// ---------------- Revisión consolidada de inasistencias ----------------
function markerIsToday_(marker, dateText){
  return String(marker || '').indexOf('|' + dateText + '|') >= 0;
}

function revisarInasistencias_(options){
  options = options || {};
  const lock = sharedLock_();
  lock.waitLock(30000);
  const today = today_();

  try{
    SpreadsheetApp.flush();
    const cfg = readConfig_();

    // Esta marca vive en la hoja Config compartida. Si tres equipos llaman la
    // revisión, solo el primero continúa y los otros dos no repiten avisos.
    if(markerIsToday_(cfg.ultimaRevisionInasistencia, today)){
      return {
        ausentes:0,
        notificados:0,
        yaEnviado:true,
        fecha:today
      };
    }
    setConfigValue_(
      'ultimaRevisionInasistencia',
      'PROCESANDO|' + today + '|' + String(Date.now())
    );

    const students = readStudents_();
    const presentRuts = new Set();
    readAttendance_().forEach(record => {
      if(normalizeDate_(record.fecha) !== today) return;
      const rut = normalizeRut_(record.rut);
      if(rut) presentRuts.add(rut);
    });

    // Si ninguna de las plataformas sincronizó registros, se bloquea el envío
    // masivo. Es más seguro no avisar que declarar ausentes a todos por una
    // falla de conexión o por una fecha mal guardada.
    if(students.length && presentRuts.size === 0){
      setConfigValue_(
        'ultimaRevisionInasistencia',
        'BLOQUEADO|' + today + '|' + String(Date.now())
      );
      return {
        ausentes:0,
        notificados:0,
        bloqueado:true,
        motivo:'No hay asistencias sincronizadas del día',
        fecha:today
      };
    }

    const absentStudents = students.filter(student => {
      const rut = normalizeRut_(student.rut);
      return rut && !presentRuts.has(rut);
    });

    let notified = 0;
    absentStudents.forEach(student => {
      if(notifyAbsence_(student, today, cfg.horaRevisionInasistencia).length) notified++;
    });

    setConfigValue_(
      'ultimaRevisionInasistencia',
      'ENVIADO|' + today + '|' + String(Date.now())
    );
    return {
      ausentes:absentStudents.length,
      notificados:notified,
      presentes:presentRuts.size,
      yaEnviado:false,
      fecha:today
    };
  }catch(err){
    setConfigValue_(
      'ultimaRevisionInasistencia',
      'ERROR|' + today + '|' + String(Date.now())
    );
    throw err;
  }finally{
    lock.releaseLock();
  }
}

function notifyAbsence_(student, dateText, reviewTime){
  const channels = [];
  const greeting = student.apoderadoNombre
    ? 'Estimado/a ' + student.apoderadoNombre + ':\n\n'
    : 'Estimado/a apoderado/a:\n\n';
  const configuredTime = normalizeTime_(reviewTime) || '12:00';
  const message = greeting +
    'A las ' + configuredTime + ' no se registra asistencia de ' + student.nombre +
    ' (' + student.curso + ') el día ' + dateText + '. ' +
    'Si el/la estudiante se encuentra en el establecimiento, por favor ignore ' +
    'este mensaje y comuníquese con Inspectoría para revisar el registro.';

  if(student.apoderadoEmail){
    try{
      MailApp.sendEmail({
        to:student.apoderadoEmail,
        subject:'Revisión de inasistencia - ' + student.nombre,
        body:message
      });
      channels.push('correo');
    }catch(err){
      console.error('No se pudo enviar correo de inasistencia: ' + err);
    }
  }

  if(student.apoderadoTelefono && student.callmebotApiKey){
    const result = notifyWhatsApp_(
      student.apoderadoTelefono,
      student.callmebotApiKey,
      message
    );
    if(result.ok) channels.push('whatsapp');
  }
  return channels;
}

function notifyWhatsApp_(phone, apiKey, message){
  if(!phone || !apiKey){
    return { ok:false, respuesta:'Falta teléfono o API Key' };
  }
  try{
    const url = 'https://api.callmebot.com/whatsapp.php?phone=' +
      encodeURIComponent(String(phone).replace(/\s+/g, '')) +
      '&text=' + encodeURIComponent(message || '') +
      '&apikey=' + encodeURIComponent(String(apiKey));
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions:true,
      followRedirects:true
    });
    const status = response.getResponseCode();
    const text = response.getContentText();
    return {
      ok:status >= 200 && status < 300 && !/error/i.test(text),
      respuesta:text,
      status:status
    };
  }catch(err){
    return { ok:false, respuesta:String(err && err.message || err) };
  }
}

// ---------------- Informe diario de asistencia ----------------
function reportRecipients_(cfg){
  const values = []
    .concat(String(cfg.correosInforme || '').split(/[;,\s]+/))
    .concat([cfg.emailDireccion, cfg.emailInspectoria])
    .map(value => String(value || '').trim())
    .filter(value => value.indexOf('@') > 0);
  return Array.from(new Set(values)).join(',');
}

function enviarInformeDiario_(){
  const cfg = readConfig_();
  const recipients = reportRecipients_(cfg);
  const today = today_();
  if(!recipients){
    return { enviado:false, motivo:'No hay destinatarios configurados', fecha:today };
  }

  const records = readAttendance_().filter(record =>
    normalizeDate_(record.fecha) === today
  );
  const fileName = 'Informe_Asistencia_' + today.replace(/\//g, '-');
  const tempSS = SpreadsheetApp.create(fileName);
  const tempSheet = tempSS.getActiveSheet();
  tempSheet.appendRow([
    'RUT',
    'Nombre',
    'Curso',
    'Fecha',
    'Hora',
    'Canal de notificación'
  ]);
  records.forEach(record => tempSheet.appendRow([
    record.rut,
    record.nombre,
    record.curso,
    normalizeDate_(record.fecha),
    normalizeTime_(record.hora),
    record.canal
  ]));
  tempSheet.setFrozenRows(1);
  tempSheet.autoResizeColumns(1, 6);

  const tempFile = DriveApp.getFileById(tempSS.getId());
  try{
    const blobPdf = tempFile.getAs(MimeType.PDF).setName(fileName + '.pdf');
    MailApp.sendEmail({
      to:recipients,
      subject:'Informe de asistencia - ' + today + ' - ' + (cfg.schoolName || ''),
      body:'Adjunto el informe de asistencia del día ' + today +
        '. Total de registros: ' + records.length + '.',
      attachments:[blobPdf]
    });
    return { enviado:true, registros:records.length, fecha:today };
  }finally{
    tempFile.setTrashed(true);
  }
}
