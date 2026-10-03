function normalizeDailyRecordDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function getDailyRecordContextV2_(ss) {
  const classes = readSheetObjects_(ss, 'Classes');
  const students = readSheetObjects_(ss, 'Students');
  const classMap = {};
  const studentMap = {};

  classes.forEach(function(row) {
    classMap[String(row.classId || '').trim()] = row;
  });
  students.forEach(function(row) {
    studentMap[String(row.studentId || '').trim()] = row;
  });

  return { classMap: classMap, studentMap: studentMap };
}

function assertDailyRecordTargetV2_(ss, classId, studentId) {
  const classKey = String(classId || '').trim();
  const studentKey = String(studentId || '').trim();

  if (!classKey) {
    const error = new Error('반을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!studentKey) {
    const error = new Error('학생을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const ctx = getDailyRecordContextV2_(ss);
  const classInfo = ctx.classMap[classKey];
  const studentInfo = ctx.studentMap[studentKey];

  if (!classInfo) {
    const error = new Error('선택한 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }
  if (!studentInfo) {
    const error = new Error('선택한 학생을 찾을 수 없습니다.');
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }
  if (String(studentInfo.classId || '').trim() !== classKey) {
    const error = new Error('선택한 학생은 이 반에 속하지 않습니다.');
    error.code = 'STUDENT_CLASS_MISMATCH';
    throw error;
  }

  return { classInfo: classInfo, studentInfo: studentInfo };
}

function decorateDailyRecordV2_(row, ctx) {
  const classId = String(row.classId || '').trim();
  const studentId = String(row.studentId || '').trim();
  const classInfo = ctx.classMap[classId] || {};
  const studentInfo = ctx.studentMap[studentId] || {};

  return Object.assign({}, row, {
    date: normalizeDailyRecordDateV2_(row.date),
    className: String(classInfo.className || classId || ''),
    teacherName: String(classInfo.teacherName || ''),
    studentName: String(studentInfo.name || studentId || '')
  });
}

function listDailyRecordsV2_(auth, options) {
  options = options || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const ctx = getDailyRecordContextV2_(ss);
  let rows = readSheetObjects_(ss, 'StudentDailyRecords');

  const classId = String(options.classId || '').trim();
  const studentId = String(options.studentId || '').trim();
  const startDate = normalizeDailyRecordDateV2_(options.startDate);
  const endDate = normalizeDailyRecordDateV2_(options.endDate);

  if (classId) {
    rows = rows.filter(function(row) {
      return String(row.classId || '').trim() === classId;
    });
  }
  if (studentId) {
    rows = rows.filter(function(row) {
      return String(row.studentId || '').trim() === studentId;
    });
  }
  if (startDate) {
    rows = rows.filter(function(row) {
      const d = normalizeDailyRecordDateV2_(row.date);
      return d && d >= startDate;
    });
  }
  if (endDate) {
    rows = rows.filter(function(row) {
      const d = normalizeDailyRecordDateV2_(row.date);
      return d && d <= endDate;
    });
  }

  rows.sort(function(a, b) {
    const ad = normalizeDailyRecordDateV2_(a.date);
    const bd = normalizeDailyRecordDateV2_(b.date);
    if (ad !== bd) return bd.localeCompare(ad);
    return String(b.recordId || '').localeCompare(String(a.recordId || ''));
  });

  return {
    records: rows.map(function(row) {
      return decorateDailyRecordV2_(row, ctx);
    })
  };
}

function getDailyRecordV2_(auth, recordId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const id = String(recordId || '').trim();
  const rows = readSheetObjects_(ss, 'StudentDailyRecords');
  const found = rows.find(function(row) {
    return String(row.recordId || '').trim() === id;
  });

  if (!found) {
    const error = new Error('개별 지도 기록을 찾을 수 없습니다.');
    error.code = 'DAILY_RECORD_NOT_FOUND';
    throw error;
  }

  return {
    record: decorateDailyRecordV2_(found, getDailyRecordContextV2_(ss))
  };
}

function saveDailyRecordV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'StudentDailyRecords');

  let recordId = String(payload.recordId || '').trim();
  const classId = String(payload.classId || '').trim();
  const studentId = String(payload.studentId || '').trim();
  const date = normalizeDailyRecordDateV2_(payload.date);
  const assignment = String(payload.assignment || '').trim();
  const comment = String(payload.comment || '').trim();
  const nextGuide = String(payload.nextGuide || '').trim();

  assertDailyRecordTargetV2_(ss, classId, studentId);

  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!assignment && !comment && !nextGuide) {
    const error = new Error('과제, 코멘트, 다음 지도 참고 중 하나 이상 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  let rowNumber = recordId ? findDataRowByIdV2_(sheet, 'recordId', recordId) : -1;
  if (recordId && rowNumber < 2) {
    const error = new Error('수정할 개별 지도 기록을 찾을 수 없습니다.');
    error.code = 'DAILY_RECORD_NOT_FOUND';
    throw error;
  }

  if (!recordId) recordId = makeNextPrefixedIdV2_(sheet, 'recordId', 'DR-', 4);

  let createdAt = new Date();
  if (rowNumber >= 2) {
    const headers = getSheetHeaderMapV2_(sheet);
    if (headers.createdAt != null) {
      createdAt = sheet.getRange(rowNumber, headers.createdAt + 1).getValue() || createdAt;
    }
  }

  const now = new Date();
  const obj = {
    recordId: recordId,
    date: date,
    classId: classId,
    studentId: studentId,
    assignment: assignment,
    comment: comment,
    nextGuide: nextGuide,
    createdAt: createdAt,
    updatedAt: now
  };

  if (rowNumber >= 2) writeObjectToRowV2_(sheet, rowNumber, obj);
  else appendObjectRowV2_(sheet, obj);

  SpreadsheetApp.flush();
  appendAuditLog_(
    auth.teacher.teacherId,
    'DAILY_RECORD_SAVE',
    'StudentDailyRecord',
    recordId,
    'SUCCESS',
    rowNumber >= 2 ? 'updated' : 'created'
  );

  return {
    recordId: recordId,
    record: decorateDailyRecordV2_(obj, getDailyRecordContextV2_(ss))
  };
}

function deleteDailyRecordV2_(auth, recordId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'StudentDailyRecords');
  const id = String(recordId || '').trim();
  const rowNumber = findDataRowByIdV2_(sheet, 'recordId', id);

  if (rowNumber < 2) {
    const error = new Error('삭제할 개별 지도 기록을 찾을 수 없습니다.');
    error.code = 'DAILY_RECORD_NOT_FOUND';
    throw error;
  }

  sheet.deleteRow(rowNumber);
  SpreadsheetApp.flush();

  appendAuditLog_(
    auth.teacher.teacherId,
    'DAILY_RECORD_DELETE',
    'StudentDailyRecord',
    id,
    'SUCCESS',
    ''
  );

  return { recordId: id, deleted: true };
}
