const ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ = '__HOLIDAY__';
const ATTENDANCE_HOLIDAY_STATUS_V2_ = '휴무';

function normalizeAttendanceDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function normalizeAttendanceStatusV2_(value) {
  const s = String(value || '').trim();
  const allowed = ['출석', '지각', '결석', '조퇴'];
  return allowed.indexOf(s) >= 0 ? s : '출석';
}

function getAttendanceClassV2_(ss, classId) {
  const id = String(classId || '').trim();
  if (!id) {
    const error = new Error('반을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const found = readSheetObjects_(ss, 'Classes').find(function(row) {
    return String(row.classId || '').trim() === id;
  });

  if (!found) {
    const error = new Error('선택한 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }

  return found;
}

function isStudentActiveOnAttendanceDateV2_(student, date) {
  const targetDate = normalizeAttendanceDateV2_(date);
  if (!targetDate) return false;

  const enrollmentDate = normalizeAttendanceDateV2_(student.enrollmentDate);
  const stopDate = normalizeAttendanceDateV2_(student.stopDate);

  if (enrollmentDate && targetDate < enrollmentDate) return false;
  if (stopDate && targetDate >= stopDate) return false;

  return true;
}

function getAttendanceTargetStudentsV2_(ss, classId, date) {
  const id = String(classId || '').trim();
  return readSheetObjects_(ss, 'Students')
    .filter(function(row) {
      return String(row.classId || '').trim() === id &&
        isStudentActiveOnAttendanceDateV2_(row, date);
    })
    .sort(function(a, b) {
      return String(a.name || '').localeCompare(String(b.name || ''), 'ko');
    });
}

function getAttendanceRowsV2_(ss, classId, date) {
  const classKey = String(classId || '').trim();
  const dateKey = normalizeAttendanceDateV2_(date);

  return readSheetObjects_(ss, 'Attendance').filter(function(row) {
    return String(row.classId || '').trim() === classKey &&
      normalizeAttendanceDateV2_(row.date) === dateKey;
  });
}

function summarizeAttendanceV2_(rows) {
  const summary = { total: 0, present: 0, late: 0, absent: 0, early: 0, holiday: false };
  (rows || []).forEach(function(row) {
    if (String(row.studentId || '').trim() === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ ||
        String(row.status || '').trim() === ATTENDANCE_HOLIDAY_STATUS_V2_) {
      summary.holiday = true;
      return;
    }

    summary.total++;
    const status = String(row.status || '').trim();
    if (status === '지각') summary.late++;
    else if (status === '결석') summary.absent++;
    else if (status === '조퇴') summary.early++;
    else summary.present++;
  });
  return summary;
}

function loadAttendanceV2_(auth, options) {
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

  if (typeof finalizeEndedQrAttendanceForClassV2_ === 'function') {
    finalizeEndedQrAttendanceForClassV2_(auth, classId, date);
  }

  const students = getAttendanceTargetStudentsV2_(ss, classId, date);
  const savedRows = getAttendanceRowsV2_(ss, classId, date);
  const holidayRow = savedRows.find(function(row) {
    return String(row.studentId || '').trim() === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ ||
      String(row.status || '').trim() === ATTENDANCE_HOLIDAY_STATUS_V2_;
  }) || null;

  const savedMap = {};
  savedRows.forEach(function(row) {
    const sid = String(row.studentId || '').trim();
    if (!sid || sid === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_) return;
    savedMap[sid] = row;
  });

  const rows = students.map(function(student) {
    const sid = String(student.studentId || '').trim();
    const saved = savedMap[sid] || {};
    return {
      attendanceId: String(saved.attendanceId || ''),
      date: date,
      classId: classId,
      studentId: sid,
      studentName: String(student.name || ''),
      status: normalizeAttendanceStatusV2_(saved.status || '출석'),
      memo: String(saved.memo || '')
    };
  });

  return {
    classInfo: classInfo,
    date: date,
    holiday: holidayRow ? {
      attendanceId: String(holidayRow.attendanceId || ''),
      memo: String(holidayRow.memo || ''),
      status: ATTENDANCE_HOLIDAY_STATUS_V2_
    } : null,
    rows: rows,
    savedRows: savedRows.map(function(row) {
      const sid = String(row.studentId || '').trim();
      const student = students.find(function(s) { return String(s.studentId || '').trim() === sid; });
      return Object.assign({}, row, {
        date: normalizeAttendanceDateV2_(row.date),
        studentName: sid === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ ? '휴무일' : String(student && student.name || sid)
      });
    }),
    summary: summarizeAttendanceV2_(savedRows)
  };
}

function findAttendanceRowNumberV2_(sheet, classId, date, studentId) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return -1;

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim(); });
  const ci = headers.indexOf('classId');
  const di = headers.indexOf('date');
  const si = headers.indexOf('studentId');
  if (ci < 0 || di < 0 || si < 0) return -1;

  const classKey = String(classId || '').trim();
  const dateKey = normalizeAttendanceDateV2_(date);
  const studentKey = String(studentId || '').trim();

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][ci] || '').trim() === classKey &&
        normalizeAttendanceDateV2_(values[i][di]) === dateKey &&
        String(values[i][si] || '').trim() === studentKey) {
      return i + 1;
    }
  }
  return -1;
}

function saveAttendanceV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Attendance');
  const classId = String(payload.classId || '').trim();
  const date = normalizeAttendanceDateV2_(payload.date);
  const records = Array.isArray(payload.records) ? payload.records : [];

  getAttendanceClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!records.length) {
    const error = new Error('저장할 출석 학생이 없습니다.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const holidayRows = getAttendanceRowsV2_(ss, classId, date).filter(function(row) {
    return String(row.studentId || '').trim() === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ ||
      String(row.status || '').trim() === ATTENDANCE_HOLIDAY_STATUS_V2_;
  });
  if (holidayRows.length) {
    const error = new Error('현재 날짜는 휴무일입니다. 휴무 해제 후 출석을 저장해 주세요.');
    error.code = 'ATTENDANCE_HOLIDAY';
    throw error;
  }

  const targetStudents = getAttendanceTargetStudentsV2_(ss, classId, date);
  const targetMap = {};
  targetStudents.forEach(function(student) {
    targetMap[String(student.studentId || '').trim()] = true;
  });

  const now = new Date();
  const saved = [];

  records.forEach(function(item) {
    const studentId = String(item.studentId || '').trim();
    if (!targetMap[studentId]) {
      const error = new Error('이 날짜의 출석 대상이 아닌 학생이 포함되어 있습니다: ' + studentId);
      error.code = 'INVALID_ATTENDANCE_STUDENT';
      throw error;
    }

    const obj = {
      attendanceId: '',
      date: date,
      classId: classId,
      studentId: studentId,
      status: normalizeAttendanceStatusV2_(item.status),
      memo: String(item.memo || '').trim(),
      updatedAt: now
    };

    const rowNumber = findAttendanceRowNumberV2_(sheet, classId, date, studentId);
    if (rowNumber >= 2) {
      const headers = getSheetHeaderMapV2_(sheet);
      if (headers.attendanceId != null) {
        obj.attendanceId = String(sheet.getRange(rowNumber, headers.attendanceId + 1).getDisplayValue() || '').trim();
      }
      if (!obj.attendanceId) obj.attendanceId = makeNextPrefixedIdV2_(sheet, 'attendanceId', 'ATT-', 4);
      writeObjectToRowV2_(sheet, rowNumber, obj);
    } else {
      obj.attendanceId = makeNextPrefixedIdV2_(sheet, 'attendanceId', 'ATT-', 4);
      appendObjectRowV2_(sheet, obj);
    }

    saved.push(obj);
  });

  SpreadsheetApp.flush();
  appendAuditLog_(
    auth.teacher.teacherId,
    'ATTENDANCE_SAVE',
    'Attendance',
    classId + ':' + date,
    'SUCCESS',
    'saved ' + saved.length + ' students'
  );

  return {
    classId: classId,
    date: date,
    savedCount: saved.length,
    rows: saved,
    summary: summarizeAttendanceV2_(saved)
  };
}

function saveAttendanceHolidayV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Attendance');
  const classId = String(payload.classId || '').trim();
  const date = normalizeAttendanceDateV2_(payload.date);
  const memo = String(payload.memo || '').trim() || '학당 휴무';

  getAttendanceClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  deleteRowsMatchingV2_(sheet, function(row) {
    return String(row.classId || '').trim() === classId &&
      normalizeAttendanceDateV2_(row.date) === date;
  });

  const obj = {
    attendanceId: makeNextPrefixedIdV2_(sheet, 'attendanceId', 'ATT-', 4),
    date: date,
    classId: classId,
    studentId: ATTENDANCE_HOLIDAY_STUDENT_ID_V2_,
    status: ATTENDANCE_HOLIDAY_STATUS_V2_,
    memo: memo,
    updatedAt: new Date()
  };

  appendObjectRowV2_(sheet, obj);
  SpreadsheetApp.flush();

  appendAuditLog_(
    auth.teacher.teacherId,
    'ATTENDANCE_HOLIDAY_SAVE',
    'Attendance',
    classId + ':' + date,
    'SUCCESS',
    memo
  );

  return { classId: classId, date: date, holiday: obj };
}

function clearAttendanceHolidayV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Attendance');
  const classId = String(payload.classId || '').trim();
  const date = normalizeAttendanceDateV2_(payload.date);

  getAttendanceClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const deleted = deleteRowsMatchingV2_(sheet, function(row) {
    return String(row.classId || '').trim() === classId &&
      normalizeAttendanceDateV2_(row.date) === date &&
      (String(row.studentId || '').trim() === ATTENDANCE_HOLIDAY_STUDENT_ID_V2_ ||
       String(row.status || '').trim() === ATTENDANCE_HOLIDAY_STATUS_V2_);
  });

  SpreadsheetApp.flush();
  appendAuditLog_(
    auth.teacher.teacherId,
    'ATTENDANCE_HOLIDAY_CLEAR',
    'Attendance',
    classId + ':' + date,
    'SUCCESS',
    'deleted ' + deleted
  );

  return { classId: classId, date: date, cleared: deleted > 0 };
}
