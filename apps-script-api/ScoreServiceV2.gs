const SCORE_HEADERS_V2_ = [
  'recordId','date','classId','studentId','examName',
  'readingScore','writingScore','listeningScore','speakingScore',
  'attitudeNote','teacherNote','createdAt'
];

function normalizeScoreDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function normalizeScoreNumberV2_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const num = Number(value);
  if (!Number.isFinite(num)) return '';
  if (num < 0) return 0;
  if (num > 100) return 100;
  return num;
}

function getScoresSheetV2_(ss) {
  const sheet = getRequiredDataSheetV2_(ss, 'Scores');
  const headers = sheet.getRange(1,1,1,Math.max(1,sheet.getLastColumn()))
    .getDisplayValues()[0]
    .map(function(v){ return String(v || '').trim(); });
  SCORE_HEADERS_V2_.forEach(function(header) {
    if (headers.indexOf(header) < 0) {
      const error = new Error('Scores 시트에 필요한 열이 없습니다: ' + header);
      error.code = 'SHEET_HEADER_ERROR';
      throw error;
    }
  });
  return sheet;
}

function getScoreClassV2_(ss, classId) {
  const id = String(classId || '').trim();
  if (!id) {
    const error = new Error('반을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  const found = readSheetObjects_(ss, 'Classes').find(function(row) {
    return String(row.classId || '').trim() === id;
  }) || null;
  if (!found) {
    const error = new Error('선택한 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }
  return found;
}

function getScoreStudentV2_(ss, studentId, classId, requireActive) {
  const sid = String(studentId || '').trim();
  const cid = String(classId || '').trim();
  if (!sid) {
    const error = new Error('학생을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  const student = readSheetObjects_(ss, 'Students').find(function(row) {
    return String(row.studentId || '').trim() === sid;
  }) || null;
  if (!student) {
    const error = new Error('학생 정보를 찾을 수 없습니다.');
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }
  if (String(student.classId || '').trim() !== cid) {
    const error = new Error('선택한 학생이 해당 반 소속이 아닙니다.');
    error.code = 'STUDENT_CLASS_MISMATCH';
    throw error;
  }
  if (requireActive && String(student.status || '').trim() !== '재학') {
    const error = new Error('현재 재학 중인 학생만 새 성적을 저장할 수 있습니다.');
    error.code = 'STUDENT_NOT_ACTIVE';
    throw error;
  }
  return student;
}

function listScoreStudentsV2_(auth, classId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  getScoreClassV2_(ss, classId);
  const cid = String(classId || '').trim();
  const students = readSheetObjects_(ss, 'Students')
    .filter(function(row) {
      return String(row.classId || '').trim() === cid &&
        String(row.status || '').trim() === '재학';
    })
    .sort(function(a,b) {
      return String(a.name || '').localeCompare(String(b.name || ''), 'ko');
    })
    .map(function(row) {
      return {
        studentId:String(row.studentId || ''),
        name:String(row.name || ''),
        classId:String(row.classId || ''),
        status:String(row.status || '')
      };
    });
  return { students:students };
}

function listScoresV2_(auth, options) {
  options = options || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  getScoresSheetV2_(ss);

  const classId = String(options.classId || '').trim();
  const studentId = String(options.studentId || '').trim();

  if (!classId && !studentId) return { scores:[] };
  if (classId) getScoreClassV2_(ss, classId);

  let rows = readSheetObjects_(ss, 'Scores');
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

  rows.sort(function(a,b) {
    const byDate = normalizeScoreDateV2_(b.date).localeCompare(normalizeScoreDateV2_(a.date));
    if (byDate !== 0) return byDate;
    return String(b.recordId || '').localeCompare(String(a.recordId || ''));
  });

  return {
    scores:rows.map(function(row) {
      return Object.assign({}, row, { date:normalizeScoreDateV2_(row.date) });
    })
  };
}

function findScoreRowV2_(sheet, recordId) {
  const target = String(recordId || '').trim();
  if (!target) return null;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return null;

  const values = sheet.getRange(1,1,lastRow,lastColumn).getDisplayValues();
  const headers = values[0].map(function(v){ return String(v || '').trim(); });
  const idIndex = headers.indexOf('recordId');
  if (idIndex < 0) return null;

  for (let i=1;i<values.length;i++) {
    if (String(values[i][idIndex] || '').trim() !== target) continue;
    const obj = {};
    headers.forEach(function(header,index) {
      if (header) obj[header] = values[i][index];
    });
    return { rowNumber:i+1, row:obj };
  }
  return null;
}

function getScoreV2_(auth, recordId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getScoresSheetV2_(ss);
  const found = findScoreRowV2_(sheet, recordId);
  if (!found) {
    const error = new Error('성적 정보를 찾을 수 없습니다.');
    error.code = 'SCORE_NOT_FOUND';
    throw error;
  }
  getScoreClassV2_(ss, found.row.classId);
  return {
    score:Object.assign({}, found.row, { date:normalizeScoreDateV2_(found.row.date) })
  };
}

function saveScoreV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getScoresSheetV2_(ss);

  const recordId = String(payload.recordId || '').trim();
  const date = normalizeScoreDateV2_(payload.date);
  const classId = String(payload.classId || '').trim();
  const studentId = String(payload.studentId || '').trim();
  const examName = String(payload.examName || '').trim();

  getScoreClassV2_(ss, classId);
  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!examName) {
    const error = new Error('시험명을 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  getScoreStudentV2_(ss, studentId, classId, !recordId);

  const found = recordId ? findScoreRowV2_(sheet, recordId) : null;
  if (recordId && !found) {
    const error = new Error('수정할 성적 정보를 찾을 수 없습니다.');
    error.code = 'SCORE_NOT_FOUND';
    throw error;
  }
  if (found) {
    getScoreClassV2_(ss, found.row.classId);
  }

  const now = new Date();
  const obj = {
    recordId: recordId || makeNextPrefixedIdV2_(sheet, 'recordId', 'SC-', 3),
    date: date,
    classId: classId,
    studentId: studentId,
    examName: examName,
    readingScore: normalizeScoreNumberV2_(payload.readingScore),
    writingScore: normalizeScoreNumberV2_(payload.writingScore),
    listeningScore: normalizeScoreNumberV2_(payload.listeningScore),
    speakingScore: normalizeScoreNumberV2_(payload.speakingScore),
    attitudeNote: String(payload.attitudeNote || '').trim(),
    teacherNote: String(payload.teacherNote || '').trim(),
    createdAt: now
  };

  if (found) {
    writeObjectToRowV2_(sheet, found.rowNumber, obj);
  } else {
    appendObjectRowV2_(sheet, obj);
  }
  SpreadsheetApp.flush();

  appendAuditLog_(
    auth.teacher.teacherId,
    found ? 'SCORE_UPDATE' : 'SCORE_CREATE',
    'Scores',
    obj.recordId,
    'SUCCESS',
    classId + ':' + studentId + ':' + examName
  );

  return {
    recordId:obj.recordId,
    score:Object.assign({}, obj, { date:date })
  };
}

function deleteScoreV2_(auth, recordId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getScoresSheetV2_(ss);
  const found = findScoreRowV2_(sheet, recordId);
  if (!found) {
    const error = new Error('삭제할 성적 정보를 찾을 수 없습니다.');
    error.code = 'SCORE_NOT_FOUND';
    throw error;
  }

  getScoreClassV2_(ss, found.row.classId);
  sheet.deleteRow(found.rowNumber);
  SpreadsheetApp.flush();

  appendAuditLog_(
    auth.teacher.teacherId,
    'SCORE_DELETE',
    'Scores',
    String(recordId || ''),
    'SUCCESS',
    String(found.row.classId || '') + ':' + String(found.row.studentId || '')
  );

  return { recordId:String(recordId || ''), deleted:true };
}
