const SCORE_HUB_SOURCE_V2_ = {
  MOBILE_ID: '1y4xaZD8SQUVLztDhBSytvi-_naVCGYX5gTOyZqmhMUE',
  TOPIK1_READING_ID: '18HXty992Riii2-csrB2aHpQ7vVt8qD2NOWFMp1yG63M',
  TOPIK1_LISTENING_ID: '1F4Bpcb4tIwuRMRy7aLb9j1YLxzxCg_wWyYEiLxkg798'
};

function scoreHubPhoneV2_(value) {
  return String(value || '').replace(/\D/g, '');
}

function scoreHubNumberV2_(value) {
  if (value === '' || value === null || value === undefined) return null;
  const n = Number(value);
  return Number.isFinite(n) ? n : null;
}

function scoreHubRowsByPhoneV2_(spreadsheetId, sheetName, phoneHeaders, phone) {
  const targetPhone = scoreHubPhoneV2_(phone);
  if (!targetPhone) return [];

  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(v) { return String(v || '').trim(); });

  let phoneIndex = -1;
  for (let i = 0; i < phoneHeaders.length; i++) {
    phoneIndex = headers.indexOf(phoneHeaders[i]);
    if (phoneIndex >= 0) break;
  }
  if (phoneIndex < 0) {
    throw new Error(sheetName + ' 전화번호 열을 찾을 수 없습니다.');
  }

  const cells = sheet.getRange(2, phoneIndex + 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(targetPhone)
    .matchEntireCell(true)
    .findAll();

  return cells.map(function(cell) {
    const values = sheet.getRange(cell.getRow(), 1, 1, lastColumn).getDisplayValues()[0];
    const row = {};
    headers.forEach(function(header, index) {
      if (header) row[header] = values[index];
    });
    return row;
  });
}

function scoreHubStudentV2_(auth, studentId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sid = String(studentId || '').trim();
  const row = readSheetObjects_(ss, 'Students').find(function(item) {
    return String(item.studentId || '').trim() === sid;
  });
  if (!row) {
    const error = new Error('학생 정보를 찾을 수 없습니다.');
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }
  return {
    studentId: String(row.studentId || ''),
    classId: String(row.classId || ''),
    name: String(row.name || ''),
    phone: scoreHubPhoneV2_(row.phone),
    status: String(row.status || '')
  };
}

function scoreHubLatestScoreV2_(rows, scoreField) {
  const scores = rows.map(function(row) { return scoreHubNumberV2_(row[scoreField]); })
    .filter(function(v) { return v !== null; });
  return {
    latestScore: scores.length ? scores[scores.length - 1] : null,
    bestScore: scores.length ? Math.max.apply(null, scores) : null,
    attempts: rows.length
  };
}

function getScoreHubSourcesV2_(auth) {
  return {
    version: 'phase1-20261003',
    sources: [
      {key:'SNU_MOBILE', label:'서울대 모바일 단원 성취', enabled:true},
      {key:'WORKBOOK_READING', label:'워크북 읽기 평가', enabled:true},
      {key:'WORKBOOK_LISTENING', label:'워크북 듣기 평가', enabled:true},
      {key:'TOPIK1_READING', label:'TOPIK I 읽기', enabled:true},
      {key:'TOPIK1_LISTENING', label:'TOPIK I 듣기', enabled:true},
      {key:'TOPIK2', label:'TOPIK II', enabled:false, deferred:true}
    ]
  };
}

function getScoreHubStudentSummaryV2_(auth, studentId) {
  const student = scoreHubStudentV2_(auth, studentId);
  if (!student.phone) {
    const error = new Error('학생 전화번호가 없어 모바일 성적을 연결할 수 없습니다.');
    error.code = 'STUDENT_PHONE_REQUIRED';
    throw error;
  }

  const tests = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.MOBILE_ID, 'TestResults', ['phone'], student.phone
  );
  const snuRows = tests.filter(function(row) {
    return /^SNU-/.test(String(row.book || '')) &&
      ['vocab','grammar','mixed'].indexOf(String(row.testType || '')) >= 0;
  });

  const wbReading = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.MOBILE_ID, '워크북읽기평가', ['전화번호'], student.phone
  );
  const wbListeningAll = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.MOBILE_ID, '워크북듣기평가', ['전화번호'], student.phone
  );
  const wbListening = wbListeningAll.filter(function(row) {
    return String(row['상태'] || '') === 'SUBMITTED';
  });

  const topik1Reading = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.TOPIK1_READING_ID, 'All_Results', ['student_phone'], student.phone
  );
  const topik1Listening = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.TOPIK1_LISTENING_ID, 'All_Results', ['student_phone'], student.phone
  );

  const progressRows = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.MOBILE_ID, '진도현황', ['전화번호'], student.phone
  );
  const progress = progressRows.length ? progressRows[0] : null;

  function splitTopik(rows) {
    return {
      exam: rows.filter(function(r) {
        return ['QUESTION_PRACTICE','WRONG_REVIEW'].indexOf(String(r.result_type || '').toUpperCase()) < 0;
      }),
      practice: rows.filter(function(r) {
        return String(r.result_type || '').toUpperCase() === 'QUESTION_PRACTICE';
      }),
      review: rows.filter(function(r) {
        return String(r.result_type || '').toUpperCase() === 'WRONG_REVIEW';
      })
    };
  }

  return {
    version:'phase1-20261003',
    student:student,
    progress:progress ? {
      firstLoginAt:String(progress['최초접속'] || ''),
      recentLoginAt:String(progress['최근접속'] || ''),
      startBook:String(progress['시작교재'] || ''),
      startLesson:String(progress['시작과'] || ''),
      currentBook:String(progress['현재교재'] || '')
    } : null,
    snu:{
      summary:scoreHubLatestScoreV2_(snuRows, 'bestScore'),
      rows:snuRows.slice(-30)
    },
    workbookReading:{
      summary:scoreHubLatestScoreV2_(wbReading, '점수(100)'),
      rows:wbReading.slice(-30)
    },
    workbookListening:{
      summary:scoreHubLatestScoreV2_(wbListening, '점수(100)'),
      inProgressCount:wbListeningAll.length - wbListening.length,
      rows:wbListening.slice(-30)
    },
    topik1Reading:splitTopik(topik1Reading),
    topik1Listening:splitTopik(topik1Listening),
    activity:{
      totalEvents:snuRows.length + wbReading.length + wbListeningAll.length +
        topik1Reading.length + topik1Listening.length,
      retryEvents:snuRows.filter(function(r){ return String(r.status || '').toUpperCase() === 'RETRY'; }).length,
      reviewEvents:topik1Reading.filter(function(r){ return String(r.result_type || '').toUpperCase() === 'WRONG_REVIEW'; }).length +
        topik1Listening.filter(function(r){ return String(r.result_type || '').toUpperCase() === 'WRONG_REVIEW'; }).length
    },
    deferred:{topik2:true}
  };
}
