function manualScoreHeaderIndexV2_(headers, name) {
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === name) return i;
  }
  return -1;
}

function manualScoreEnsureHeadersV2_(sheet, requiredHeaders) {
  var lastColumn = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) {
    return String(v || '').trim();
  });
  requiredHeaders.forEach(function(name) {
    if (manualScoreHeaderIndexV2_(headers, name) >= 0) return;
    headers.push(name);
    sheet.getRange(1, headers.length).setValue(name);
  });
  return headers;
}

function ensureManualScoreStructureV2_(ss) {
  var sheet = ss.getSheetByName('Scores');
  if (!sheet) {
    var error = new Error('Scores 시트를 찾을 수 없습니다.');
    error.code = 'SHEET_NOT_FOUND';
    throw error;
  }
  manualScoreEnsureHeadersV2_(sheet, [
    'scoreType', 'book', 'lesson', 'vocabScore', 'grammarScore', 'mixedScore',
    'maxScore', 'sourceType', 'updatedAt'
  ]);
  return sheet;
}

function manualScoreNormalizeDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function manualScoreNumberOrBlankV2_(value, fieldLabel, maxScore) {
  if (value === '' || value === null || value === undefined) return '';
  var n = Number(value);
  if (!isFinite(n)) {
    var e1 = new Error(fieldLabel + ' 점수를 숫자로 입력해 주세요.');
    e1.code = 'VALIDATION_ERROR';
    throw e1;
  }
  if (n < 0 || n > maxScore) {
    var e2 = new Error(fieldLabel + ' 점수는 0~' + maxScore + ' 사이로 입력해 주세요.');
    e2.code = 'VALIDATION_ERROR';
    throw e2;
  }
  return n;
}

function manualScoreAllowedLessonV2_(book, lesson) {
  var ranges = {
    'SNU-1A':[1,8], 'SNU-1B':[9,16], 'SNU-2A':[1,9], 'SNU-2B':[10,18],
    'SNU-3A':[1,9], 'SNU-3B':[10,18], 'SNU-4A':[1,9], 'SNU-4B':[10,18]
  };
  var r = ranges[String(book || '').trim()];
  var n = Number(lesson);
  return !!r && isFinite(n) && n >= r[0] && n <= r[1] && Math.floor(n) === n;
}

function manualScoreFindStudentV2_(ss, studentId) {
  var sheet = ss.getSheetByName('Students');
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return null;
  var headers = sheet.getRange(1,1,1,lastColumn).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  var idx = manualScoreHeaderIndexV2_(headers, 'studentId');
  if (idx < 0) return null;
  var id = String(studentId || '').trim();
  var found = sheet.getRange(2, idx + 1, lastRow - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  if (!found) return null;
  var values = sheet.getRange(found.getRow(),1,1,lastColumn).getDisplayValues()[0];
  var obj = {};
  headers.forEach(function(h,i){ if(h) obj[h] = values[i]; });
  return obj;
}

function manualScoreFindRowV2_(sheet, recordId) {
  var id = String(recordId || '').trim();
  if (!id) return -1;
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  var idx = manualScoreHeaderIndexV2_(headers, 'recordId');
  if (idx < 0 || sheet.getLastRow() < 2) return -1;
  var found = sheet.getRange(2, idx + 1, sheet.getLastRow() - 1, 1).createTextFinder(id).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}

function manualScoreReadRowV2_(sheet, rowNumber) {
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  var values = sheet.getRange(rowNumber,1,1,sheet.getLastColumn()).getDisplayValues()[0];
  var obj = {};
  headers.forEach(function(h,i){ if(h) obj[h] = values[i]; });
  obj.date = manualScoreNormalizeDateV2_(obj.date);
  return obj;
}

function listTeacherManualScoresV2_(auth, studentId) {
  var id = String(studentId || '').trim();
  if (!id) {
    var e0 = new Error('학생을 선택해 주세요.'); e0.code = 'VALIDATION_ERROR'; throw e0;
  }
  var ss = getTeacherDataSpreadsheet_(auth.teacher);
  if (!manualScoreFindStudentV2_(ss, id)) {
    var e1 = new Error('학생 정보를 찾을 수 없습니다.'); e1.code = 'STUDENT_NOT_FOUND'; throw e1;
  }
  var sheet = ensureManualScoreStructureV2_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [] };
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getDisplayValues()[0].map(function(v){return String(v||'').trim();});
  var studentIdx = manualScoreHeaderIndexV2_(headers, 'studentId');
  if (studentIdx < 0) return { rows: [] };
  var matches = sheet.getRange(2, studentIdx + 1, lastRow - 1, 1).createTextFinder(id).matchEntireCell(true).findAll();
  var rows = (matches || []).map(function(cell){ return manualScoreReadRowV2_(sheet, cell.getRow()); }).filter(function(row){
    return String(row.sourceType || '').trim().toUpperCase() === 'TEACHER_MANUAL';
  });
  rows.sort(function(a,b){
    var d = String(b.date || '').localeCompare(String(a.date || ''));
    if (d !== 0) return d;
    return String(b.updatedAt || b.createdAt || '').localeCompare(String(a.updatedAt || a.createdAt || ''));
  });
  return { rows: rows };
}

function saveTeacherManualScoreV2_(auth, payload) {
  payload = payload || {};
  var ss = getTeacherDataSpreadsheet_(auth.teacher);
  var sheet = ensureManualScoreStructureV2_(ss);
  var studentId = String(payload.studentId || '').trim();
  var student = manualScoreFindStudentV2_(ss, studentId);
  if (!student) { var e0 = new Error('학생 정보를 찾을 수 없습니다.'); e0.code='STUDENT_NOT_FOUND'; throw e0; }

  var scoreType = String(payload.scoreType || 'SNU').trim().toUpperCase();
  if (['SNU','TOPIK_I','TOPIK_II','GENERAL'].indexOf(scoreType) < 0) {
    var e1 = new Error('지원하지 않는 평가 유형입니다.'); e1.code='VALIDATION_ERROR'; throw e1;
  }
  var date = manualScoreNormalizeDateV2_(payload.date);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) { var e2 = new Error('평가일을 입력해 주세요.'); e2.code='VALIDATION_ERROR'; throw e2; }
  var maxScore = Number(payload.maxScore || 100);
  if (!isFinite(maxScore) || maxScore <= 0 || maxScore > 1000) { var e3 = new Error('만점 기준을 확인해 주세요.'); e3.code='VALIDATION_ERROR'; throw e3; }

  var book = String(payload.book || '').trim();
  var lesson = String(payload.lesson || '').trim();
  var examName = String(payload.examName || '').trim();
  var vocabScore = manualScoreNumberOrBlankV2_(payload.vocabScore, '어휘', maxScore);
  var grammarScore = manualScoreNumberOrBlankV2_(payload.grammarScore, '문법', maxScore);
  var mixedScore = manualScoreNumberOrBlankV2_(payload.mixedScore, '종합', maxScore);
  var readingScore = manualScoreNumberOrBlankV2_(payload.readingScore, '읽기', maxScore);
  var writingScore = manualScoreNumberOrBlankV2_(payload.writingScore, '쓰기', maxScore);
  var listeningScore = manualScoreNumberOrBlankV2_(payload.listeningScore, '듣기', maxScore);
  var speakingScore = manualScoreNumberOrBlankV2_(payload.speakingScore, '말하기', maxScore);

  if (scoreType === 'SNU') {
    if (!manualScoreAllowedLessonV2_(book, lesson)) { var e4 = new Error('서울대 교재와 과를 정확히 선택해 주세요.'); e4.code='VALIDATION_ERROR'; throw e4; }
    if (readingScore === '' && listeningScore === '' && writingScore === '' && mixedScore === '') { var e5 = new Error('읽기·듣기·쓰기·종합 중 하나 이상의 점수를 입력해 주세요.'); e5.code='VALIDATION_ERROR'; throw e5; }
    examName = examName || (book + ' ' + lesson + '과 단원평가');
    vocabScore = ''; grammarScore = ''; speakingScore = '';
  } else {
    book = ''; lesson = ''; vocabScore = ''; grammarScore = ''; mixedScore = '';
    if (scoreType === 'TOPIK_I') {
      writingScore = ''; speakingScore = '';
      if (readingScore === '' && listeningScore === '') { var e6 = new Error('TOPIK I 읽기 또는 듣기 점수를 입력해 주세요.'); e6.code='VALIDATION_ERROR'; throw e6; }
      examName = examName || 'TOPIK I 평가';
    } else if (scoreType === 'TOPIK_II') {
      speakingScore = '';
      if (readingScore === '' && writingScore === '' && listeningScore === '') { var e7 = new Error('TOPIK II 읽기·쓰기·듣기 중 하나 이상의 점수를 입력해 주세요.'); e7.code='VALIDATION_ERROR'; throw e7; }
      examName = examName || 'TOPIK II 평가';
    } else {
      if (readingScore === '' && writingScore === '' && listeningScore === '' && speakingScore === '') { var e8 = new Error('읽기·쓰기·듣기·말하기 중 하나 이상의 점수를 입력해 주세요.'); e8.code='VALIDATION_ERROR'; throw e8; }
      if (!examName) { var e9 = new Error('평가명을 입력해 주세요.'); e9.code='VALIDATION_ERROR'; throw e9; }
    }
  }

  var recordId = String(payload.recordId || '').trim();
  var rowNumber = recordId ? manualScoreFindRowV2_(sheet, recordId) : -1;
  var createdAt = new Date();
  if (recordId && rowNumber < 2) { var e10 = new Error('수정할 성적 기록을 찾을 수 없습니다.'); e10.code='SCORE_NOT_FOUND'; throw e10; }
  if (rowNumber >= 2) {
    var old = manualScoreReadRowV2_(sheet, rowNumber);
    if (String(old.sourceType || '').trim().toUpperCase() !== 'TEACHER_MANUAL') { var e11 = new Error('자동 성적 기록은 이 화면에서 수정할 수 없습니다.'); e11.code='READ_ONLY_SCORE'; throw e11; }
    if (String(old.studentId || '').trim() !== studentId) { var e12 = new Error('다른 학생의 성적 기록은 수정할 수 없습니다.'); e12.code='VALIDATION_ERROR'; throw e12; }
    createdAt = old.createdAt || createdAt;
  } else {
    recordId = 'SCR-' + Utilities.getUuid().replace(/-/g,'').slice(0,16).toUpperCase();
  }

  var now = new Date();
  var obj = {
    recordId: recordId,
    date: date,
    classId: String(student.classId || '').trim(),
    studentId: studentId,
    examName: examName,
    readingScore: readingScore,
    writingScore: writingScore,
    listeningScore: listeningScore,
    speakingScore: speakingScore,
    attitudeNote: String(payload.attitudeNote || '').trim(),
    teacherNote: String(payload.teacherNote || '').trim(),
    createdAt: createdAt,
    scoreType: scoreType,
    book: book,
    lesson: lesson,
    vocabScore: vocabScore,
    grammarScore: grammarScore,
    mixedScore: mixedScore,
    maxScore: maxScore,
    sourceType: 'TEACHER_MANUAL',
    updatedAt: now
  };

  if (rowNumber >= 2) writeObjectToRowV2_(sheet, rowNumber, obj);
  else appendObjectRowV2_(sheet, obj);
  SpreadsheetApp.flush();
  appendAuditLog_(auth.teacher.teacherId, 'MANUAL_SCORE_SAVE', 'Score', recordId, 'SUCCESS', scoreType + ' / ' + studentId);
  return { recordId: recordId, score: obj, mode: rowNumber >= 2 ? 'updated' : 'created' };
}

function deleteTeacherManualScoreV2_(auth, recordId) {
  var ss = getTeacherDataSpreadsheet_(auth.teacher);
  var sheet = ensureManualScoreStructureV2_(ss);
  var rowNumber = manualScoreFindRowV2_(sheet, recordId);
  if (rowNumber < 2) { var e1 = new Error('삭제할 성적 기록을 찾을 수 없습니다.'); e1.code='SCORE_NOT_FOUND'; throw e1; }
  var row = manualScoreReadRowV2_(sheet, rowNumber);
  if (String(row.sourceType || '').trim().toUpperCase() !== 'TEACHER_MANUAL') { var e2 = new Error('자동 성적 기록은 삭제할 수 없습니다.'); e2.code='READ_ONLY_SCORE'; throw e2; }
  sheet.deleteRow(rowNumber);
  SpreadsheetApp.flush();
  appendAuditLog_(auth.teacher.teacherId, 'MANUAL_SCORE_DELETE', 'Score', String(recordId || ''), 'SUCCESS', String(row.scoreType || '') + ' / ' + String(row.studentId || ''));
  return { recordId: String(recordId || ''), deleted: true };
}
