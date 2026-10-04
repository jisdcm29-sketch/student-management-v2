function reportEmailHeaderIndexV2_(headers, name) {
  for (var i = 0; i < headers.length; i++) if (String(headers[i] || '').trim() === name) return i;
  return -1;
}

function reportEmailEnsureHeadersV2_(sheet, requiredHeaders) {
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  requiredHeaders.forEach(function(name) {
    if (reportEmailHeaderIndexV2_(headers, name) >= 0) return;
    headers.push(name);
    sheet.getRange(1, headers.length).setValue(name);
  });
  return headers;
}

function ensureReportEmailDataStructureV2_(ss) {
  var students = ss.getSheetByName('Students');
  if (!students) {
    var error = new Error('Students 시트를 찾을 수 없습니다.');
    error.code = 'SHEET_NOT_FOUND';
    throw error;
  }
  reportEmailEnsureHeadersV2_(students, ['studentEmail', 'guardianName', 'guardianEmail']);

  var log = ss.getSheetByName('ReportSendLog');
  var created = false;
  if (!log) {
    log = ss.insertSheet('ReportSendLog');
    created = true;
  }
  var headers = ['logId','studentId','reportType','startDate','endDate','recipientType','recipientEmail','subject','eventType','eventAt','actorTeacherId','status','errorMessage'];
  reportEmailEnsureHeadersV2_(log, headers);
  if (created) log.setFrozenRows(1);
  return { studentsSheet: students, logSheet: log };
}

function reportEmailFindStudentRowV2_(sheet, studentId) {
  var id = String(studentId || '').trim();
  if (!id) return -1;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return -1;
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  var idx = reportEmailHeaderIndexV2_(headers, 'studentId');
  if (idx < 0) return -1;
  var finder = sheet.getRange(2, idx + 1, lastRow - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return finder ? finder.getRow() : -1;
}

function reportEmailReadStudentContactV2_(sheet, rowNumber) {
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  var values = sheet.getRange(rowNumber, 1, 1, lastColumn).getDisplayValues()[0];
  function v(name) { var i = reportEmailHeaderIndexV2_(headers, name); return i >= 0 ? String(values[i] || '').trim() : ''; }
  return { studentId:v('studentId'), name:v('name'), studentEmail:v('studentEmail'), guardianName:v('guardianName'), guardianEmail:v('guardianEmail') };
}

function reportEmailAddressValidV2_(value) {
  var v = String(value || '').trim();
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(v);
}

function getReportEmailContactV2_(auth, studentId) {
  var ss = getTeacherDataSpreadsheet_(auth.teacher);
  var setup = ensureReportEmailDataStructureV2_(ss);
  var row = reportEmailFindStudentRowV2_(setup.studentsSheet, studentId);
  if (row < 2) {
    var error = new Error('학생 정보를 찾을 수 없습니다.');
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }
  return reportEmailReadStudentContactV2_(setup.studentsSheet, row);
}

function prepareReportEmailV2_(auth, payload) {
  payload = payload || {};
  var studentId = String(payload.studentId || '').trim();
  var recipientEmail = String(payload.recipientEmail || '').trim();
  var recipientType = String(payload.recipientType || 'MANUAL').trim().toUpperCase();
  var reportType = String(payload.reportType || 'PARENT').trim().toUpperCase();
  var subject = String(payload.subject || '').trim();
  var startDate = String(payload.startDate || '').trim();
  var endDate = String(payload.endDate || '').trim();
  var eventType = String(payload.eventType || 'COMPOSE_OPENED').trim().toUpperCase();

  if (!studentId) { var e1 = new Error('학생 정보가 없습니다.'); e1.code='VALIDATION_ERROR'; throw e1; }
  if (!reportEmailAddressValidV2_(recipientEmail)) { var e2 = new Error('수신 이메일 주소 형식을 확인해 주세요.'); e2.code='VALIDATION_ERROR'; throw e2; }

  var ss = getTeacherDataSpreadsheet_(auth.teacher);
  var setup = ensureReportEmailDataStructureV2_(ss);
  var row = reportEmailFindStudentRowV2_(setup.studentsSheet, studentId);
  if (row < 2) { var e3 = new Error('학생 정보를 찾을 수 없습니다.'); e3.code='STUDENT_NOT_FOUND'; throw e3; }

  var logId = 'RSL-' + Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
  var eventAt = new Date();
  var obj = {
    logId: logId,
    studentId: studentId,
    reportType: reportType,
    startDate: startDate,
    endDate: endDate,
    recipientType: recipientType,
    recipientEmail: recipientEmail,
    subject: subject,
    eventType: eventType,
    eventAt: eventAt,
    actorTeacherId: String(auth.teacher && auth.teacher.teacherId || ''),
    status: 'PREPARED',
    errorMessage: ''
  };

  var logSheet = setup.logSheet;
  var headers = logSheet.getRange(1,1,1,logSheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  var rowValues = headers.map(function(h){ return Object.prototype.hasOwnProperty.call(obj,h) ? obj[h] : ''; });
  logSheet.appendRow(rowValues);
  SpreadsheetApp.flush();

  appendAuditLog_(auth.teacher.teacherId, 'REPORT_EMAIL_PREPARE', 'Student', studentId, 'SUCCESS', logId + ' / ' + recipientType);
  return { logId: logId, status: 'PREPARED', eventType: eventType };
}
