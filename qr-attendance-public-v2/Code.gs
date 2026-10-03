const V2_PUBLIC_QR_CONFIG = {
  registrySpreadsheetId: '14UljkUze6SSG8nIu8oC5LWTdI30fwNTOJQo0ERCytaE',
  sessionSheet: 'QrAttendanceSessions',
  teachersSheet: 'Teachers',
  auditSheet: 'AuditLog',
  studentsSheet: 'Students',
  attendanceSheet: 'Attendance',
  openBeforeMinutes: 30,
  lateFromMinutes: 20,
};

const V2_PUBLIC_QR_SESSION_HEADERS = [
  'sessionId','publicToken','sessionCode','teacherId','teacherName','dataSpreadsheetId',
  'classId','className','date','startTime','endTime','openTime','onTimeUntil',
  'lateFrom','lateUntil','status','createdAt','createdBy','closedAt','updatedAt'
];

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.qrToken = e && e.parameter ? String(e.parameter.qr || '').trim() : '';
  return template.evaluate()
    .setTitle('QR 출석 체크인')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupPublicQrAttendanceV2App() {
  const sheet = ensurePublicQrSessionSheetV2_();
  return {
    success: true,
    message: 'V2 QR 공개 체크인 앱 준비가 완료되었습니다.',
    sheetName: sheet.getName()
  };
}

function getPublicQrSessionInfoV2(publicToken) {
  try {
    const session = findPublicQrSessionV2_(publicToken);
    if (!session) throw new Error('유효하지 않거나 종료된 QR 출석 세션입니다.');
    assertPublicQrTeacherRouteV2_(session);
    if (publicQrSessionEndedV2_(session)) {
      throw new Error('수업이 종료되어 QR 출석이 마감되었습니다.');
    }
    return {
      success: true,
      session: {
        className: String(session.className || session.classId || ''),
        date: normalizePublicQrDateV2_(session.date),
        startTime: normalizePublicQrTimeV2_(session.startTime),
        endTime: normalizePublicQrTimeV2_(session.endTime)
      }
    };
  } catch (e) {
    return { success:false, error:e.message || String(e) };
  }
}

function submitPublicQrAttendanceV2(payload) {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    payload = payload || {};
    const publicToken = String(payload.publicToken || '').trim();
    const phone = normalizePublicQrPhoneV2_(payload.phone);
    if (!publicToken) throw new Error('QR 출석 세션 정보가 없습니다.');
    if (!phone) throw new Error('전화번호를 입력해 주세요.');

    const session = findPublicQrSessionV2_(publicToken);
    if (!session) throw new Error('유효하지 않거나 종료된 QR 출석 세션입니다.');
    const teacher = assertPublicQrTeacherRouteV2_(session);

    const tz = Session.getScriptTimeZone() || 'Asia/Ulaanbaatar';
    const now = new Date();
    const today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    const sessionDate = normalizePublicQrDateV2_(session.date);
    if (today !== sessionDate) throw new Error('오늘 수업의 QR 코드가 아닙니다.');

    const startMinutes = publicQrTimeToMinutesV2_(session.startTime);
    const endMinutes = publicQrTimeToMinutesV2_(session.endTime);
    if (startMinutes < 0 || endMinutes < 0 || endMinutes <= startMinutes) {
      throw new Error('수업 시간을 확인할 수 없습니다.');
    }
    const nowMinutes = Number(Utilities.formatDate(now, tz, 'H')) * 60 + Number(Utilities.formatDate(now, tz, 'm'));
    const elapsed = nowMinutes - startMinutes;

    if (elapsed < -V2_PUBLIC_QR_CONFIG.openBeforeMinutes) {
      throw new Error('아직 QR 출석 시간이 아닙니다. 수업 시작 30분 전부터 가능합니다.');
    }
    if (nowMinutes < startMinutes) {
      throw new Error('수업 시작 전입니다. 출석 체크는 수업 시작 시각부터 가능합니다.');
    }
    if (nowMinutes >= endMinutes) {
      throw new Error('수업이 종료되어 QR 출석이 마감되었습니다. 미체크 학생은 결석 처리됩니다.');
    }

    const dataSpreadsheetId = String(teacher.dataSpreadsheetId || '').trim();
    const ss = SpreadsheetApp.openById(dataSpreadsheetId);
    assertPublicQrClassV2_(ss, session.classId);

    const student = findPublicQrStudentByPhoneV2_(ss, phone, session.classId, sessionDate);
    if (!student) throw new Error('이 반의 재학생으로 등록된 전화번호를 찾을 수 없습니다.');

    if (!lock.tryLock(10000)) throw new Error('출석 저장이 처리 중입니다. 잠시 후 다시 시도해 주세요.');
    locked = true;

    const attendanceSheet = ss.getSheetByName(V2_PUBLIC_QR_CONFIG.attendanceSheet);
    if (!attendanceSheet) throw new Error('Attendance 시트를 찾을 수 없습니다.');
    const headers = publicQrHeadersV2_(attendanceSheet);

    if (hasPublicQrHolidayV2_(attendanceSheet, headers, sessionDate, session.classId)) {
      throw new Error('이 날짜는 휴무일로 설정되어 있습니다.');
    }

    const existing = findPublicQrExistingAttendanceV2_(
      attendanceSheet, headers, sessionDate, session.classId, student.studentId
    );
    const checkInTime = Utilities.formatDate(now, tz, 'HH:mm:ss');

    if (existing) {
      return {
        success: true,
        alreadyRecorded: true,
        studentName: String(student.name || ''),
        status: String(existing.status || '출석'),
        checkInTime: extractPublicQrMemoTimeV2_(existing.memo) || checkInTime,
        message: '이미 출석 기록이 있습니다. 중복 저장하지 않았습니다.'
      };
    }

    const status = elapsed >= V2_PUBLIC_QR_CONFIG.lateFromMinutes ? '지각' : '출석';
    const memo = 'QR 체크인 ' + checkInTime + ' · 세션 ' + String(session.sessionCode || '');
    const attendanceId = nextPublicQrAttendanceIdV2_(attendanceSheet, headers);

    const rowObj = {
      attendanceId: attendanceId,
      date: sessionDate,
      classId: String(session.classId || ''),
      studentId: String(student.studentId || ''),
      status: status,
      memo: memo,
      updatedAt: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss')
    };
    const row = headers.map(function(h){ return rowObj[h] == null ? '' : rowObj[h]; });
    attendanceSheet.getRange(attendanceSheet.getLastRow()+1,1,1,headers.length).setValues([row]);
    SpreadsheetApp.flush();

    appendPublicQrAuditV2_(
      String(session.teacherId || ''),
      'QR_ATTENDANCE_CHECKIN',
      'Attendance',
      attendanceId,
      'SUCCESS',
      String(student.studentId || '') + ' / ' + status
    );

    return {
      success: true,
      studentName: String(student.name || ''),
      status: status,
      checkInTime: checkInTime,
      message: status === '지각' ? '지각으로 출석 처리되었습니다.' : '출석 처리되었습니다.'
    };
  } catch (e) {
    return { success:false, error:e.message || String(e) };
  } finally {
    if (locked) {
      try { lock.releaseLock(); } catch (ignore) {}
    }
  }
}

function publicQrSessionEndedV2_(session) {
  const tz = Session.getScriptTimeZone() || 'Asia/Ulaanbaatar';
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const sessionDate = normalizePublicQrDateV2_(session.date);
  const endMinutes = publicQrTimeToMinutesV2_(session.endTime);
  if (!sessionDate || endMinutes < 0) return false;
  if (today > sessionDate) return true;
  if (today < sessionDate) return false;
  const nowMinutes = Number(Utilities.formatDate(new Date(), tz, 'H')) * 60 +
    Number(Utilities.formatDate(new Date(), tz, 'm'));
  return nowMinutes >= endMinutes;
}

function registrySpreadsheetV2_() {
  return SpreadsheetApp.openById(V2_PUBLIC_QR_CONFIG.registrySpreadsheetId);
}

function ensurePublicQrSessionSheetV2_() {
  const ss = registrySpreadsheetV2_();
  let sheet = ss.getSheetByName(V2_PUBLIC_QR_CONFIG.sessionSheet);
  if (!sheet) sheet = ss.insertSheet(V2_PUBLIC_QR_CONFIG.sessionSheet);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,V2_PUBLIC_QR_SESSION_HEADERS.length).setValues([V2_PUBLIC_QR_SESSION_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const current = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
  V2_PUBLIC_QR_SESSION_HEADERS.forEach(function(h){
    if (current.indexOf(h) < 0) {
      sheet.getRange(1,sheet.getLastColumn()+1).setValue(h);
      current.push(h);
    }
  });
  return sheet;
}

function findPublicQrSessionV2_(publicToken) {
  const token = String(publicToken || '').trim();
  if (!token) return null;

  const sheet = ensurePublicQrSessionSheetV2_();
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return null;

  const headers = values[0].map(function(v){ return String(v || '').trim(); });
  const tokenIdx = headers.indexOf('publicToken');
  const statusIdx = headers.indexOf('status');
  if (tokenIdx < 0) return null;

  for (let i=values.length-1;i>=1;i--) {
    if (String(values[i][tokenIdx] || '').trim() !== token) continue;
    if (statusIdx >= 0 && String(values[i][statusIdx] || '').trim() !== 'PREPARED') return null;
    const obj = {};
    headers.forEach(function(h,j){ if(h) obj[h]=values[i][j]; });
    return obj;
  }
  return null;
}

function assertPublicQrTeacherRouteV2_(session) {
  const teacherId = String(session.teacherId || '').trim();
  const dataSpreadsheetId = String(session.dataSpreadsheetId || '').trim();
  if (!teacherId || !dataSpreadsheetId) throw new Error('QR 출석 데이터 연결 정보가 없습니다.');

  const sheet = registrySpreadsheetV2_().getSheetByName(V2_PUBLIC_QR_CONFIG.teachersSheet);
  if (!sheet) throw new Error('Teachers 시트를 찾을 수 없습니다.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('교사 등록 정보를 찾을 수 없습니다.');
  const headers = values[0].map(function(v){ return String(v || '').trim(); });
  const teacherIdx = headers.indexOf('teacherId');
  const statusIdx = headers.indexOf('status');
  const dataIdx = headers.indexOf('dataSpreadsheetId');

  for (let i=1;i<values.length;i++) {
    if (String(values[i][teacherIdx] || '').trim() !== teacherId) continue;
    const status = String(values[i][statusIdx] || '').trim().toUpperCase();
    const registeredDataId = String(values[i][dataIdx] || '').trim();
    if (status !== 'ACTIVE') throw new Error('사용할 수 없는 교사 계정의 QR입니다.');
    if (registeredDataId !== dataSpreadsheetId) throw new Error('QR 출석 데이터 연결이 변경되었습니다. 새 QR을 사용해 주세요.');
    return { teacherId:teacherId, dataSpreadsheetId:registeredDataId };
  }
  throw new Error('QR 출석 교사 정보를 확인할 수 없습니다.');
}

function assertPublicQrClassV2_(ss, classId) {
  const sheet = ss.getSheetByName('Classes');
  if (!sheet) throw new Error('Classes 시트를 찾을 수 없습니다.');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) throw new Error('반 정보를 찾을 수 없습니다.');
  const headers = values[0].map(function(v){ return String(v || '').trim(); });
  const classIdx = headers.indexOf('classId');
  const statusIdx = headers.indexOf('status');
  const target = String(classId || '').trim();

  for (let i=1;i<values.length;i++) {
    if (String(values[i][classIdx] || '').trim() !== target) continue;
    if (statusIdx >= 0 && String(values[i][statusIdx] || '').trim() === '종료') {
      throw new Error('종료된 반의 QR 출석은 사용할 수 없습니다.');
    }
    return true;
  }
  throw new Error('QR 출석 반 정보를 찾을 수 없습니다.');
}

function findPublicQrStudentByPhoneV2_(ss, phone, classId, date) {
  const sheet = ss.getSheetByName(V2_PUBLIC_QR_CONFIG.studentsSheet);
  if (!sheet) throw new Error('Students 시트를 찾을 수 없습니다.');
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return null;

  const headers = values[0].map(function(v){ return String(v || '').trim(); });
  const idx = publicQrIndexMapV2_(headers);
  const targetClass = String(classId || '').trim();
  const targets = publicQrPhoneCandidatesV2_(phone);
  const targetDate = normalizePublicQrDateV2_(date);

  for (let i=1;i<values.length;i++) {
    const row = values[i];
    if (String(row[idx.classId] || '').trim() !== targetClass) continue;
    if (String(row[idx.status] || '').trim() !== '재학') continue;

    const enrollmentDate = idx.enrollmentDate >= 0 ? normalizePublicQrDateV2_(row[idx.enrollmentDate]) : '';
    if (enrollmentDate && targetDate < enrollmentDate) continue;

    const candidates = publicQrPhoneCandidatesV2_(row[idx.phone]);
    if (!candidates.some(function(v){ return targets.indexOf(v) >= 0; })) continue;

    return {
      studentId:String(row[idx.studentId] || '').trim(),
      name:String(row[idx.name] || '').trim()
    };
  }
  return null;
}

function publicQrHeadersV2_(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) throw new Error('Attendance 시트 헤더가 없습니다.');
  return sheet.getRange(1,1,1,lastCol).getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
}

function hasPublicQrHolidayV2_(sheet, headers, date, classId) {
  const idx = publicQrIndexMapV2_(headers);
  const values = sheet.getDataRange().getDisplayValues();
  for (let i=values.length-1;i>=1;i--) {
    const row=values[i];
    if (normalizePublicQrDateV2_(row[idx.date]) !== date) continue;
    if (String(row[idx.classId] || '').trim() !== String(classId || '').trim()) continue;
    if (String(row[idx.studentId] || '').trim() === '__HOLIDAY__' || String(row[idx.status] || '').trim() === '휴무') return true;
  }
  return false;
}

function findPublicQrExistingAttendanceV2_(sheet, headers, date, classId, studentId) {
  const idx = publicQrIndexMapV2_(headers);
  const values = sheet.getDataRange().getDisplayValues();
  for (let i=values.length-1;i>=1;i--) {
    const row=values[i];
    if (normalizePublicQrDateV2_(row[idx.date]) !== date) continue;
    if (String(row[idx.classId] || '').trim() !== String(classId || '').trim()) continue;
    if (String(row[idx.studentId] || '').trim() !== String(studentId || '').trim()) continue;
    const obj={};
    headers.forEach(function(h,j){ if(h) obj[h]=row[j]; });
    return obj;
  }
  return null;
}

function nextPublicQrAttendanceIdV2_(sheet, headers) {
  const idIdx = headers.indexOf('attendanceId');
  if (idIdx < 0) throw new Error('Attendance 시트에 attendanceId 열이 없습니다.');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 'ATT-0001';

  const ids = sheet.getRange(2,idIdx+1,lastRow-1,1).getDisplayValues();
  let maxNum=0;
  ids.forEach(function(r){
    const m=String(r[0] || '').match(/^ATT-(\d+)$/);
    if (m) maxNum=Math.max(maxNum,Number(m[1]));
  });
  return 'ATT-' + String(maxNum+1).padStart(4,'0');
}

function appendPublicQrAuditV2_(teacherId, action, targetType, targetId, result, detail) {
  try {
    const sheet = registrySpreadsheetV2_().getSheetByName(V2_PUBLIC_QR_CONFIG.auditSheet);
    if (!sheet) return;
    sheet.appendRow([
      'LOG-' + Utilities.getUuid(),
      new Date(),
      teacherId || '',
      action || '',
      targetType || '',
      targetId || '',
      result || '',
      detail || ''
    ]);
  } catch (e) {}
}

function publicQrIndexMapV2_(headers) {
  const map={};
  headers.forEach(function(h,i){ map[h]=i; });
  return new Proxy(map,{get:function(target,prop){ return Object.prototype.hasOwnProperty.call(target,prop) ? target[prop] : -1; }});
}

function normalizePublicQrDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  return String(value).trim().split('T')[0].split(' ')[0];
}

function normalizePublicQrTimeV2_(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'HH:mm');
  const text=String(value).trim();
  const match=text.match(/(?:^|[ T])(\d{1,2}):(\d{2})(?::\d{2})?(?:$|\s)/) || text.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return '';
  return String(Number(match[1])).padStart(2,'0') + ':' + match[2];
}

function publicQrTimeToMinutesV2_(value) {
  const t=normalizePublicQrTimeV2_(value);
  if (!t) return -1;
  const p=t.split(':').map(Number);
  return p[0]*60+p[1];
}

function normalizePublicQrPhoneV2_(value) {
  return String(value == null ? '' : value).replace(/[^0-9]/g,'');
}

function publicQrPhoneCandidatesV2_(value) {
  const phone=normalizePublicQrPhoneV2_(value);
  const out=[];
  if (phone) out.push(phone);
  if (phone.length>8) out.push(phone.slice(-8));
  return out.filter(function(v,i,a){ return v && a.indexOf(v)===i; });
}

function extractPublicQrMemoTimeV2_(memo) {
  const m=String(memo || '').match(/QR 체크인\s+(\d{2}:\d{2}:\d{2})/);
  return m ? m[1] : '';
}
