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

const ALUMNOS_COLS = ['rut','nombre','curso','apoderadoNombre','apoderadoTelefono','apoderadoEmail','callmebotApiKey'];
const ASISTENCIA_COLS = ['id','rut','nombre','curso','fecha','hora','timestamp','canal'];
const CONFIG_DEFAULTS = {
  schoolName: 'Liceo Tiuquilemu',
  adminPasswordHash: '',
  userPasswordHash: '',
  emailDireccion: '',
  emailInspectoria: '',
  horaEnvio: '18:00'
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

// ---------------- Config ----------------
function readConfig_(){
  const sheet = getSheet_(SHEET_CONFIG, ['clave','valor']);
  const data = sheet.getDataRange().getValues();
  const cfg = Object.assign({}, CONFIG_DEFAULTS);
  for(let i = 1; i < data.length; i++){
    if(data[i][0]) cfg[data[i][0]] = data[i][1];
  }
  return cfg;
}

function writeConfig_(partial){
  const sheet = getSheet_(SHEET_CONFIG, ['clave','valor']);
  const current = readConfig_();
  const merged = Object.assign({}, current, partial);
  sheet.clearContents();
  sheet.appendRow(['clave','valor']);
  Object.keys(merged).forEach(k => sheet.appendRow([k, merged[k]]));
  if(merged.horaEnvio) asegurarTriggerDiario_(merged.horaEnvio);
  return merged;
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
    if(String(data[i][0]).toUpperCase() === String(rut).toUpperCase()) return i + 1; // fila real (1-indexada)
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
    ASISTENCIA_COLS.forEach((c, idx) => row[c] = data[i][idx] || '');
    out.push(row);
  }
  return out.reverse(); // más reciente primero
}

function addAttendance_(record){
  const sheet = getSheet_(SHEET_ASISTENCIA, ASISTENCIA_COLS);
  sheet.appendRow(ASISTENCIA_COLS.map(c => record[c] !== undefined ? record[c] : ''));
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

      case 'send_report_now':
        enviarInformeDiario_();
        return jsonResponse_({ ok:true });

      default:
        return jsonResponse_({ ok:false, error:'Tipo de solicitud no reconocido' });
    }
  }catch(err){
    return jsonResponse_({ ok:false, error: String(err) });
  }
}

// ---------------- Informe diario automático ----------------
function asegurarTriggerDiario_(horaEnvio){
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if(t.getHandlerFunction() === 'enviarInformeDiario_') ScriptApp.deleteTrigger(t);
  });
  const hora = parseInt(String(horaEnvio || '18:00').split(':')[0], 10);
  ScriptApp.newTrigger('enviarInformeDiario_')
    .timeBased()
    .everyDays(1)
    .atHour(isNaN(hora) ? 18 : hora)
    .create();
}

function enviarInformeDiario_(){
  const cfg = readConfig_();
  const destinatarios = [cfg.emailDireccion, cfg.emailInspectoria].filter(String).join(',');
  if(!destinatarios) return;

  const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const registros = readAttendance_().filter(r => r.fecha === hoy);

  const nombreTemp = 'Informe_Asistencia_' + hoy.replace(/\//g, '-');
  const tempSS = SpreadsheetApp.create(nombreTemp);
  const tempHoja = tempSS.getActiveSheet();
  tempHoja.appendRow(['RUT','Nombre','Curso','Fecha','Hora','Canal de notificación']);
  registros.forEach(r => tempHoja.appendRow([r.rut, r.nombre, r.curso, r.fecha, r.hora, r.canal]));

  const blobExcel = DriveApp.getFileById(tempSS.getId()).getAs(MimeType.MICROSOFT_EXCEL);
  MailApp.sendEmail({
    to: destinatarios,
    subject: 'Informe de asistencia - ' + hoy + ' - ' + (cfg.schoolName || ''),
    body: 'Adjunto el informe de asistencia del día ' + hoy + '. Total de registros: ' + registros.length + '.',
    attachments: [blobExcel]
  });
  DriveApp.getFileById(tempSS.getId()).setTrashed(true);
}
