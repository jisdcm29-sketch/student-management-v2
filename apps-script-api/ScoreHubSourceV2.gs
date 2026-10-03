const SCORE_HUB_SOURCE_V2_ = {
  MOBILE_ID: '1y4xaZD8SQUVLztDhBSytvi-_naVCGYX5gTOyZqmhMUE',
  TOPIK1_READING_ID: '18HXty992Riii2-csrB2aHpQ7vVt8qD2NOWFMp1yG63M'
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

function scoreHubTestResultSummaryV2_(rows) {
  const base = scoreHubLatestScoreV2_(rows, 'bestScore');
  base.attempts = rows.reduce(function(sum, row) {
    return sum + (scoreHubNumberV2_(row.attemptsToday) || 0);
  }, 0);
  base.passRows = rows.filter(function(row) {
    return String(row.status || '').trim().toUpperCase() === 'PASS';
  }).length;
  base.retryRows = rows.filter(function(row) {
    return String(row.status || '').trim().toUpperCase() === 'RETRY';
  }).length;
  return base;
}

function scoreHubBuildSnuMasteryV2_(rows) {
  const byLesson = {};
  rows.forEach(function(row) {
    const book = String(row.book || '').trim();
    const lesson = String(row.lesson || '').trim();
    const type = String(row.testType || '').trim();
    if (!book || !lesson || ['vocab','grammar','mixed'].indexOf(type) < 0) return;
    const key = book + '|' + lesson;
    if (!byLesson[key]) {
      byLesson[key] = {
        book: book,
        lesson: lesson,
        vocab: { bestScore:null, attempts:0 },
        grammar: { bestScore:null, attempts:0 },
        mixed: { bestScore:null, attempts:0 }
      };
    }
    const item = byLesson[key][type];
    const score = scoreHubNumberV2_(row.bestScore);
    const attempts = scoreHubNumberV2_(row.attemptsToday) || 0;
    item.attempts += attempts;
    if (score !== null && (item.bestScore === null || score > item.bestScore)) item.bestScore = score;
  });

  return Object.keys(byLesson).map(function(key) {
    const item = byLesson[key];
    item.passed = ['vocab','grammar','mixed'].every(function(type) {
      return item[type].bestScore !== null && Number(item[type].bestScore) >= 90;
    });
    item.status = item.passed ? 'PASS' : 'RETRY';
    item.totalAttempts = item.vocab.attempts + item.grammar.attempts + item.mixed.attempts;
    return item;
  }).sort(function(a,b) {
    const bookCmp = String(a.book).localeCompare(String(b.book));
    if (bookCmp !== 0) return bookCmp;
    return Number(a.lesson || 0) - Number(b.lesson || 0);
  });
}

function getScoreHubSourcesV2_(auth) {
  return {
    version: 'phase1-20261003',
    sources: [
      {key:'SNU_VOCAB', label:'서울대 각 과 어휘 테스트', enabled:true},
      {key:'SNU_GRAMMAR', label:'서울대 각 과 문법 테스트', enabled:true},
      {key:'SNU_MIXED', label:'서울대 각 과 종합 테스트', enabled:true},
      {key:'REVIEW_READING', label:'복습 읽기 평가', enabled:true},
      {key:'REVIEW_LISTENING', label:'복습 듣기 평가', enabled:true},
      {key:'TOPIK1_COLLOCATION', label:'TOPIK I 연어 시험', enabled:true},
      {key:'TOPIK1_GRAMMAR', label:'TOPIK I 문법 시험', enabled:true},
      {key:'TOPIK1_READING', label:'TOPIK I 읽기평가', enabled:true},
      {key:'TOPIK1_LISTENING', label:'TOPIK I 듣기', enabled:false, deferred:true},
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
  const snuVocabRows = snuRows.filter(function(row) {
    return String(row.testType || '') === 'vocab';
  });
  const snuGrammarRows = snuRows.filter(function(row) {
    return String(row.testType || '') === 'grammar';
  });
  const snuMixedRows = snuRows.filter(function(row) {
    return String(row.testType || '') === 'mixed';
  });

  const topik1CollocationRows = tests.filter(function(row) {
    return String(row.book || '') === 'TOPIK1' &&
      String(row.testType || '') === 'collocation';
  });
  const topik1GrammarRows = tests.filter(function(row) {
    return String(row.book || '') === 'TOPIK1' &&
      String(row.testType || '') === 'grammar';
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

  const progressRows = scoreHubRowsByPhoneV2_(
    SCORE_HUB_SOURCE_V2_.MOBILE_ID, '진도현황', ['전화번호'], student.phone
  );
  const progress = progressRows.length ? progressRows[0] : null;

  function splitTopik(rows) {
    function pack(items) {
      return {
        summary: scoreHubLatestScoreV2_(items, 'section_score_100'),
        rows: items.slice(-30)
      };
    }
    return {
      exam: pack(rows.filter(function(r) {
        return ['QUESTION_PRACTICE','WRONG_REVIEW'].indexOf(String(r.result_type || '').toUpperCase()) < 0;
      })),
      practice: pack(rows.filter(function(r) {
        return String(r.result_type || '').toUpperCase() === 'QUESTION_PRACTICE';
      })),
      review: pack(rows.filter(function(r) {
        return String(r.result_type || '').toUpperCase() === 'WRONG_REVIEW';
      }))
    };
  }

  return {
    version:'phase1-20261003',
    student:student,
    progress:progress ? {
      firstLoginAt:String(progress['최초접속'] || ''),
      recentLoginAt:String(progress['최근접속'] || ''),
      elapsedText:String(progress['접속경과'] || ''),
      startBook:String(progress['시작교재'] || ''),
      startLesson:String(progress['시작과'] || ''),
      currentBook:String(progress['현재교재'] || ''),
      currentLesson:String(progress['현재과'] || ''),
      passProgress:String(progress['통과현황'] || ''),
      vocabBest:scoreHubNumberV2_(progress['어휘최고']),
      grammarBest:scoreHubNumberV2_(progress['문법최고']),
      mixedBest:scoreHubNumberV2_(progress['종합최고']),
      currentState:String(progress['현재상태'] || ''),
      nextStep:String(progress['다음단계'] || ''),
      topikCollocationState:String(progress['TOPIK연어'] || ''),
      topikGrammarState:String(progress['TOPIK문법'] || ''),
      recentActivityAt:String(progress['최근활동'] || ''),
      recentTest:String(progress['최근시험'] || ''),
      recentScore:scoreHubNumberV2_(progress['최근점수']),
      totalAttempts:scoreHubNumberV2_(progress['총응시']) || 0
    } : null,
    snu:{
      vocab:{
        summary:scoreHubTestResultSummaryV2_(snuVocabRows),
        rows:snuVocabRows.slice(-30)
      },
      grammar:{
        summary:scoreHubTestResultSummaryV2_(snuGrammarRows),
        rows:snuGrammarRows.slice(-30)
      },
      mixed:{
        summary:scoreHubTestResultSummaryV2_(snuMixedRows),
        rows:snuMixedRows.slice(-30)
      },
      rows:snuRows.slice(-60),
      mastery:scoreHubBuildSnuMasteryV2_(snuRows)
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
    topik1Collocation:{
      summary:scoreHubTestResultSummaryV2_(topik1CollocationRows),
      rows:topik1CollocationRows.slice(-30)
    },
    topik1Grammar:{
      summary:scoreHubTestResultSummaryV2_(topik1GrammarRows),
      rows:topik1GrammarRows.slice(-30)
    },
    topik1Reading:splitTopik(topik1Reading),
    activity:{
      totalEvents:snuRows.length + wbReading.length + wbListeningAll.length +
        topik1CollocationRows.length + topik1GrammarRows.length + topik1Reading.length,
      retryEvents:snuRows.filter(function(r){ return String(r.status || '').toUpperCase() === 'RETRY'; }).length +
        topik1CollocationRows.filter(function(r){ return String(r.status || '').toUpperCase() === 'RETRY'; }).length +
        topik1GrammarRows.filter(function(r){ return String(r.status || '').toUpperCase() === 'RETRY'; }).length,
      reviewEvents:topik1Reading.filter(function(r){ return String(r.result_type || '').toUpperCase() === 'WRONG_REVIEW'; }).length
    },
    deferred:{topik1Listening:true, topik2:true}
  };
}
