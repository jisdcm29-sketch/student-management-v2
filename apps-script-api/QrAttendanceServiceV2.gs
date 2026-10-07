const V2_QR_OPEN_BEFORE_MINUTES = 30;
const V2_QR_LATE_FROM_MINUTES = 20;
const V2_QR_CLOSE_BEFORE_END_MINUTES = 20;
const V2_QR_SESSION_SHEET = 'QrAttendanceSessions';
const V2_QR_PUBLIC_URL_PROPERTY = 'V2_QR_PUBLIC_APP_URL';
const V2_QR_SESSION_HEADERS = [
  'sessionId','publicToken','sessionCode','teacherId','teacherName','dataSpreadsheetId',
  'classId','className','date','startTime','endTime','openTime','onTimeUntil',
  'lateFrom','lateUntil','status','createdAt','createdBy','closedAt','updatedAt'
];

function normalizeQrTimeV2_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'HH:mm');
  }
  const text = String(value == null ? '' : value).trim();
  const match = text.match(/(?:^|[ T])(\d{1,2}):(\d{2})(?::\d{2})?(?:$|\s)/) ||
                text.match(/^(\d{1,2}):(\d{2})/);
  if (!match) return '';
  const h = Number(match[1]);
  const m = Number(match[2]);
  if (h < 0 || h > 23 || m < 0 || m > 59) return '';
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

function qrTimeToMinutesV2_(value) {
  const text = normalizeQrTimeV2_(value);
  if (!text) return -1;
  const parts = text.split(':').map(Number);
  return parts[0] * 60 + parts[1];
}

function qrMinutesToTimeV2_(minutes) {
  let value = Number(minutes);
  if (!isFinite(value)) return '';
  value = ((value % 1440) + 1440) % 1440;
  const h = Math.floor(value / 60);
  const m = value % 60;
  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

function buildQrRuleTimesV2_(startTime, endTime) {
  const startMin = qrTimeToMinutesV2_(startTime);
  const endMin = qrTimeToMinutesV2_(endTime);
  if (startMin < 0) {
    return { openTime:'', onTimeUntil:'', lateFrom:'', lateUntil:'' };
  }
  const lateFromMin = startMin + V2_QR_LATE_FROM_MINUTES;
  const closeMin = endMin > startMin
    ? Math.max(lateFromMin, endMin - V2_QR_CLOSE_BEFORE_END_MINUTES)
    : -1;
  return {
    openTime: qrMinutesToTimeV2_(startMin - V2_QR_OPEN_BEFORE_MINUTES),
    onTimeUntil: qrMinutesToTimeV2_(lateFromMin),
    lateFrom: qrMinutesToTimeV2_(lateFromMin),
    lateUntil: closeMin >= 0 ? qrMinutesToTimeV2_(closeMin) : ''
  };
}

function koreanDayFromDateV2_(date) {
  const parts = String(date || '').split('-').map(Number);
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) return '';
  const d = new Date(parts[0], parts[1]-1, parts[2], 12, 0, 0);
  return ['일','월','화','수','목','금','토'][d.getDay()] || '';
}

function findQrScheduleV2_(ss, classId, date) {
  const targetClassId = String(classId || '').trim();
  const day = koreanDayFromDateV2_(date);
  const schedules = readSheetObjects_(ss, 'ClassSchedules');

  const exact = schedules.find(function(row) {
    return String(row.classId || '').trim() === targetClassId &&
      String(row.dayOfWeek || '').trim() === day &&
      normalizeQrTimeV2_(row.startTime) &&
      normalizeQrTimeV2_(row.endTime);
  }) || null;

  if (!exact) return null;

  return {
    startTime: normalizeQrTimeV2_(exact.startTime),
    endTime: normalizeQrTimeV2_(exact.endTime),
    source: 'CLASS_SCHEDULE'
  };
}

function ensureQrSessionRegistryV2_() {
  const ss = getRegistrySpreadsheet_();
  let sheet = ss.getSheetByName(V2_QR_SESSION_SHEET);
  if (!sheet) sheet = ss.insertSheet(V2_QR_SESSION_SHEET);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1,1,1,V2_QR_SESSION_HEADERS.length).setValues([V2_QR_SESSION_HEADERS]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const current = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn())).getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
  V2_QR_SESSION_HEADERS.forEach(function(header) {
    if (current.indexOf(header) < 0) {
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(header);
      current.push(header);
    }
  });
  return sheet;
}

function qrSessionRowsV2_() {
  const sheet = ensureQrSessionRegistryV2_();
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  const values = sheet.getRange(1,1,lastRow,lastColumn).getDisplayValues();
  const headers = values[0].map(function(v){ return String(v || '').trim(); });
  return values.slice(1).map(function(row, index) {
    const obj = { __rowNumber:index+2 };
    headers.forEach(function(h,i){ if(h) obj[h]=row[i]; });
    return obj;
  });
}

function getActiveQrSessionV2_(auth, classId, date) {
  const teacherId = String(auth.teacher.teacherId || '').trim();
  const dataSpreadsheetId = String(auth.teacher.dataSpreadsheetId || '').trim();
  const classKey = String(classId || '').trim();
  const dateKey = normalizeAttendanceDateV2_(date);

  const rows = qrSessionRowsV2_();
  for (let i=rows.length-1;i>=0;i--) {
    const row = rows[i];
    if (String(row.teacherId || '').trim() !== teacherId) continue;
    if (String(row.dataSpreadsheetId || '').trim() !== dataSpreadsheetId) continue;
    if (String(row.classId || '').trim() !== classKey) continue;
    if (normalizeAttendanceDateV2_(row.date) !== dateKey) continue;
    if (String(row.status || '').trim() !== 'PREPARED') continue;
    return row;
  }
  return null;
}

function publicQrAppUrlV2_() {
  return String(PropertiesService.getScriptProperties().getProperty(V2_QR_PUBLIC_URL_PROPERTY) || '').trim();
}


function sanitizeQrSessionForClientV2_(session) {
  if (!session) return null;
  const safeSession = Object.assign({}, session);
  delete safeSession.dataSpreadsheetId;
  return safeSession;
}

function getQrAttendanceSetupV2_(auth, options) {
  options = options || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classId = String(options.classId || '').trim();
  const date = normalizeAttendanceDateV2_(options.date);

  const classInfo = getAttendanceClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const schedule = findQrScheduleV2_(ss, classId, date);
  finalizeEndedQrAttendanceForClassV2_(auth, classId, date);
  const activeSession = getActiveQrSessionV2_(auth, classId, date);
  const startTime = activeSession && activeSession.startTime
    ? normalizeQrTimeV2_(activeSession.startTime)
    : (schedule ? schedule.startTime : '');
  const endTime = activeSession && activeSession.endTime
    ? normalizeQrTimeV2_(activeSession.endTime)
    : (schedule ? schedule.endTime : '');
  const rules = buildQrRuleTimesV2_(startTime, endTime);

  if (activeSession) {
    activeSession.publicUrl = publicQrAppUrlV2_()
      ? publicQrAppUrlV2_() + '?qr=' + encodeURIComponent(String(activeSession.publicToken || ''))
      : '';
  }

  return {
    classId: classId,
    className: String(classInfo.className || classId),
    date: date,
    scheduleFound: !!schedule,
    scheduleSource: schedule ? schedule.source : '',
    startTime: startTime,
    endTime: endTime,
    openTime: rules.openTime,
    onTimeUntil: rules.onTimeUntil,
    lateFrom: rules.lateFrom,
    lateUntil: rules.lateUntil,
    publicAppConfigured: !!publicQrAppUrlV2_(),
    activeSession: sanitizeQrSessionForClientV2_(activeSession)
  };
}

function qrSessionHasEndedV2_(date, endTime) {
  const dateKey = normalizeAttendanceDateV2_(date);
  const timeKey = normalizeQrTimeV2_(endTime);
  if (!dateKey || !timeKey) return false;
  const tz = Session.getScriptTimeZone() || 'Asia/Ulaanbaatar';
  const nowKey = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm');
  return nowKey >= (dateKey + ' ' + timeKey);
}

function finalizeQrAbsencesForSessionV2_(auth, session) {
  if (!session || !qrSessionHasEndedV2_(session.date, session.endTime)) {
    return { finalized:false, absentCount:0 };
  }

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classId = String(session.classId || '').trim();
  const date = normalizeAttendanceDateV2_(session.date);
  const attendanceSheet = getRequiredDataSheetV2_(ss, 'Attendance');

  const existingRows = getAttendanceRowsV2_(ss, classId, date);
  const holiday = existingRows.some(function(row) {
    return String(row.studentId || '').trim() === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ ||
      String(row.status || '').trim() === ATTENDANCE_HOLIDAY_STATUS_V2_;
  });
  if (holiday) return { finalized:false, absentCount:0, holiday:true };

  const existingMap = {};
  existingRows.forEach(function(row) {
    const sid = String(row.studentId || '').trim();
    if (sid && sid !== ATTENDANCE_HOLIDAY_STUDENT_ID_V2_) existingMap[sid] = true;
  });

  const targets = getAttendanceTargetStudentsV2_(ss, classId, date);
  const now = new Date();
  let absentCount = 0;

  targets.forEach(function(student) {
    const studentId = String(student.studentId || '').trim();
    if (!studentId || existingMap[studentId]) return;

    appendObjectRowV2_(attendanceSheet, {
      attendanceId: makeNextPrefixedIdV2_(attendanceSheet, 'attendanceId', 'ATT-', 4),
      date: date,
      classId: classId,
      studentId: studentId,
      status: '결석',
      memo: 'QR 미체크 · 수업 종료 자동 결석',
      updatedAt: now
    });
    existingMap[studentId] = true;
    absentCount++;
  });

  if (absentCount > 0) SpreadsheetApp.flush();

  appendAuditLog_(
    auth.teacher.teacherId,
    'QR_ABSENCE_FINALIZE',
    'Attendance',
    classId + ':' + date,
    'SUCCESS',
    'auto absent ' + absentCount
  );

  return { finalized:true, absentCount:absentCount };
}

function finalizeEndedQrAttendanceForClassV2_(auth, classId, date) {
  const teacherId = String(auth.teacher.teacherId || '').trim();
  const dataSpreadsheetId = String(auth.teacher.dataSpreadsheetId || '').trim();
  const classKey = String(classId || '').trim();
  const dateKey = normalizeAttendanceDateV2_(date);
  const rows = qrSessionRowsV2_();

  for (let i=rows.length-1;i>=0;i--) {
    const row = rows[i];
    if (String(row.teacherId || '').trim() !== teacherId) continue;
    if (String(row.dataSpreadsheetId || '').trim() !== dataSpreadsheetId) continue;
    if (String(row.classId || '').trim() !== classKey) continue;
    if (normalizeAttendanceDateV2_(row.date) !== dateKey) continue;
    if (String(row.status || '').trim() !== 'PREPARED') continue;
    if (!qrSessionHasEndedV2_(row.date, row.endTime)) return { finalized:false, absentCount:0 };

    const result = finalizeQrAbsencesForSessionV2_(auth, row);
    closePreparedQrRowsV2_(auth, classKey, dateKey);
    SpreadsheetApp.flush();
    return result;
  }
  return { finalized:false, absentCount:0 };
}

function closePreparedQrRowsV2_(auth, classId, date) {
  const sheet = ensureQrSessionRegistryV2_();
  const rows = qrSessionRowsV2_();
  const headers = getHeaderMap_(sheet);
  const now = new Date();
  let count = 0;

  rows.forEach(function(row) {
    if (String(row.teacherId || '').trim() !== String(auth.teacher.teacherId || '').trim()) return;
    if (String(row.dataSpreadsheetId || '').trim() !== String(auth.teacher.dataSpreadsheetId || '').trim()) return;
    if (String(row.classId || '').trim() !== String(classId || '').trim()) return;
    if (normalizeAttendanceDateV2_(row.date) !== normalizeAttendanceDateV2_(date)) return;
    if (String(row.status || '').trim() !== 'PREPARED') return;

    if (headers.status != null) sheet.getRange(row.__rowNumber, headers.status+1).setValue('CLOSED');
    if (headers.closedAt != null) sheet.getRange(row.__rowNumber, headers.closedAt+1).setValue(now);
    if (headers.updatedAt != null) sheet.getRange(row.__rowNumber, headers.updatedAt+1).setValue(now);
    count++;
  });
  return count;
}

function startQrAttendanceSessionV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classId = String(payload.classId || '').trim();
  const date = normalizeAttendanceDateV2_(payload.date);
  const startTime = normalizeQrTimeV2_(payload.startTime);
  const endTime = normalizeQrTimeV2_(payload.endTime);

  const classInfo = getAttendanceClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!startTime || !endTime) {
    const error = new Error('수업 시작시간과 종료시간을 확인해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  const startMin = qrTimeToMinutesV2_(startTime);
  const endMin = qrTimeToMinutesV2_(endTime);
  if (startMin < 0 || endMin < 0 || endMin <= startMin) {
    const error = new Error('수업 종료시간은 시작시간보다 늦어야 합니다.');
    error.code = 'INVALID_TIME_RANGE';
    throw error;
  }
  if (startMin % 30 !== 0 || endMin % 30 !== 0) {
    const error = new Error('수업 시작시간과 종료시간은 30분 단위로 설정해 주세요.');
    error.code = 'INVALID_TIME_STEP';
    throw error;
  }

  const publicAppUrl = publicQrAppUrlV2_();
  if (!publicAppUrl) {
    const error = new Error('V2 학생용 QR 공개 웹앱 URL이 아직 등록되지 않았습니다.');
    error.code = 'QR_PUBLIC_APP_NOT_CONFIGURED';
    throw error;
  }

  closePreparedQrRowsV2_(auth, classId, date);

  const schedule = findQrScheduleV2_(ss, classId, date);
  const rules = buildQrRuleTimesV2_(startTime, endTime);
  const now = new Date();
  const uuid = Utilities.getUuid().replace(/-/g,'');
  const session = {
    sessionId: 'V2QR-' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyyMMdd-HHmmss') + '-' + uuid.slice(0,8),
    publicToken: uuid,
    sessionCode: uuid.slice(0,6).toUpperCase(),
    teacherId: String(auth.teacher.teacherId || ''),
    teacherName: String(auth.teacher.displayName || ''),
    dataSpreadsheetId: String(auth.teacher.dataSpreadsheetId || ''),
    classId: classId,
    className: String(classInfo.className || classId),
    date: date,
    startTime: startTime,
    endTime: endTime,
    openTime: rules.openTime,
    onTimeUntil: rules.onTimeUntil,
    lateFrom: rules.lateFrom,
    lateUntil: rules.lateUntil,
    status: 'PREPARED',
    createdAt: now,
    createdBy: String(auth.teacher.displayName || auth.teacher.teacherId || ''),
    closedAt: '',
    updatedAt: now
  };

  const sheet = ensureQrSessionRegistryV2_();
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
  const row = headers.map(function(h){ return session[h] == null ? '' : session[h]; });
  sheet.getRange(sheet.getLastRow()+1,1,1,headers.length).setValues([row]);
  SpreadsheetApp.flush();

  session.publicUrl = publicAppUrl + '?qr=' + encodeURIComponent(session.publicToken);
  session.scheduleFound = !!schedule;
  session.scheduleSource = schedule ? schedule.source : '';

  appendAuditLog_(
    auth.teacher.teacherId,
    'QR_SESSION_START',
    'QrAttendanceSession',
    session.sessionId,
    'SUCCESS',
    classId + ':' + date + ' ' + startTime + '-' + endTime
  );

  return {
    session: sanitizeQrSessionForClientV2_(session),
    scheduleFound: !!schedule,
    scheduleSource: schedule ? schedule.source : '',
    startTime: startTime,
    endTime: endTime,
    openTime: rules.openTime,
    onTimeUntil: rules.onTimeUntil,
    lateFrom: rules.lateFrom,
    lateUntil: rules.lateUntil
  };
}

function closeQrAttendanceSessionV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classId = String(payload.classId || '').trim();
  const date = normalizeAttendanceDateV2_(payload.date);

  getAttendanceClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const activeSession = getActiveQrSessionV2_(auth, classId, date);
  const finalization = activeSession
    ? finalizeQrAbsencesForSessionV2_(auth, activeSession)
    : { finalized:false, absentCount:0 };
  const closed = closePreparedQrRowsV2_(auth, classId, date);
  SpreadsheetApp.flush();

  appendAuditLog_(
    auth.teacher.teacherId,
    'QR_SESSION_CLOSE',
    'QrAttendanceSession',
    classId + ':' + date,
    'SUCCESS',
    'closed ' + closed
  );

  return {
    classId:classId,
    date:date,
    closedCount:closed,
    absenceFinalized:!!finalization.finalized,
    autoAbsentCount:Number(finalization.absentCount || 0)
  };
}


