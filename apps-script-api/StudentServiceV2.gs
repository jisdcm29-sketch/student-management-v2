function normalizeStudentDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function isArchivedStudentStatusV2_(status) {
  const s = String(status || '').trim();
  return s === '휴학' || s === '중단' || s === '중도포기';
}

function needsStudentStopDateV2_(status) {
  return isArchivedStudentStatusV2_(status);
}

function assertStudentClassExistsV2_(ss, classId) {
  const id = String(classId || '').trim();
  if (!id) {
    const error = new Error('반을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const classes = readSheetObjects_(ss, 'Classes');
  const found = classes.some(function(row) {
    return String(row.classId || '').trim() === id;
  });

  if (!found) {
    const error = new Error('선택한 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }
}

function listStudentsV2_(auth, options) {
  options = options || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  let rows = readSheetObjects_(ss, 'Students');

  if (options.classId) {
    const classId = String(options.classId || '').trim();
    rows = rows.filter(function(row) {
      return String(row.classId || '').trim() === classId;
    });
  }

  if (options.currentOnly !== false) {
    rows = rows.filter(function(row) {
      return !isArchivedStudentStatusV2_(row.status);
    });
  }

  rows = rows.map(function(row) {
    return Object.assign({}, row, {
      enrollmentDate: normalizeStudentDateV2_(row.enrollmentDate),
      scholarshipStartDate: normalizeStudentDateV2_(row.scholarshipStartDate),
      scholarshipEndDate: normalizeStudentDateV2_(row.scholarshipEndDate),
      paidUntilDate: normalizeStudentDateV2_(row.paidUntilDate),
      stopDate: normalizeStudentDateV2_(row.stopDate)
    });
  });

  rows.sort(function(a, b) {
    const byClass = String(a.classId || '').localeCompare(String(b.classId || ''));
    if (byClass !== 0) return byClass;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ko');
  });

  return { students: rows };
}

function getStudentV2_(auth, studentId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const students = readSheetObjects_(ss, 'Students');
  const id = String(studentId || '').trim();
  const student = students.find(function(row) {
    return String(row.studentId || '').trim() === id;
  });

  if (!student) {
    const error = new Error('학생 정보를 찾을 수 없습니다.');
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }

  return {
    student: Object.assign({}, student, {
      enrollmentDate: normalizeStudentDateV2_(student.enrollmentDate),
      scholarshipStartDate: normalizeStudentDateV2_(student.scholarshipStartDate),
      scholarshipEndDate: normalizeStudentDateV2_(student.scholarshipEndDate),
      paidUntilDate: normalizeStudentDateV2_(student.paidUntilDate),
      stopDate: normalizeStudentDateV2_(student.stopDate)
    })
  };
}

function saveStudentV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Students');

  let studentId = String(payload.studentId || '').trim();
  const classId = String(payload.classId || '').trim();
  const name = String(payload.name || '').trim();
  const phone = String(payload.phone || '').trim();
  const currentTopikLevel = String(payload.currentTopikLevel || '').trim();
  const targetTopikLevel = String(payload.targetTopikLevel || '').trim();
  const scholarshipType = String(payload.scholarshipType || '').trim();
  const scholarshipStartDate = normalizeStudentDateV2_(payload.scholarshipStartDate);
  const scholarshipEndDate = normalizeStudentDateV2_(payload.scholarshipEndDate);
  const paidUntilDate = normalizeStudentDateV2_(payload.paidUntilDate);
  const enrollmentDate = normalizeStudentDateV2_(payload.enrollmentDate);
  const status = String(payload.status || '재학').trim() || '재학';
  const stopDate = normalizeStudentDateV2_(payload.stopDate);
  const note = String(payload.note || '').trim();
  const photoUrl = String(payload.photoUrl || '').trim();

  assertStudentClassExistsV2_(ss, classId);

  if (!name) {
    const error = new Error('이름을 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!enrollmentDate) {
    const error = new Error('등록일을 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (needsStudentStopDateV2_(status) && !stopDate) {
    const error = new Error('중단 / 중도포기 / 휴학 날짜를 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  let rowNumber = studentId ? findDataRowByIdV2_(sheet, 'studentId', studentId) : -1;
  if (studentId && rowNumber < 2) {
    const error = new Error('수정할 학생을 찾을 수 없습니다.');
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }

  if (!studentId) studentId = makeNextPrefixedIdV2_(sheet, 'studentId', 'S-', 3);

  let createdAt = new Date();
  if (rowNumber >= 2) {
    const headers = getSheetHeaderMapV2_(sheet);
    if (headers.createdAt != null) {
      createdAt = sheet.getRange(rowNumber, headers.createdAt + 1).getValue() || createdAt;
    }
  }

  const obj = {
    studentId: studentId,
    classId: classId,
    photoUrl: photoUrl,
    name: name,
    phone: phone,
    currentTopikLevel: currentTopikLevel,
    targetTopikLevel: targetTopikLevel,
    scholarshipType: scholarshipType,
    scholarshipStartDate: scholarshipStartDate,
    scholarshipEndDate: scholarshipEndDate,
    paidUntilDate: paidUntilDate,
    enrollmentDate: enrollmentDate,
    status: status,
    stopDate: needsStudentStopDateV2_(status) ? stopDate : '',
    note: note,
    createdAt: createdAt
  };

  if (rowNumber >= 2) writeObjectToRowV2_(sheet, rowNumber, obj);
  else appendObjectRowV2_(sheet, obj);

  SpreadsheetApp.flush();
  appendAuditLog_(
    auth.teacher.teacherId,
    'STUDENT_SAVE',
    'Student',
    studentId,
    'SUCCESS',
    rowNumber >= 2 ? 'updated' : 'created'
  );

  return {
    studentId: studentId,
    student: obj
  };
}
