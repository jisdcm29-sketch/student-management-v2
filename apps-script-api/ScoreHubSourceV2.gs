const SCORE_HUB_SOURCE_V2_ = {
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

function scoreHubBooleanV2_(value) {
  if (value === true || value === false) return value;
  return String(value || '').trim().toUpperCase() === 'TRUE';
}

function scoreHubRowsByPhoneFromSpreadsheetV2_(ss, sheetName, phoneHeaders, phone) {
  const targetPhone = scoreHubPhoneV2_(phone);
  if (!targetPhone) return [];

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];

  const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim(); });

  let phoneIndex = -1;
  for (let i = 0; i < phoneHeaders.length; i++) {
    phoneIndex = headers.indexOf(phoneHeaders[i]);
    if (phoneIndex >= 0) break;
  }
  if (phoneIndex < 0) {
    const error = new Error(sheetName + ' 전화번호 열을 찾을 수 없습니다.');
    error.code = 'SCORE_SOURCE_PHONE_HEADER_NOT_FOUND';
    throw error;
  }

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    if (scoreHubPhoneV2_(values[r][phoneIndex]) !== targetPhone) continue;
    const row = {};
    headers.forEach(function(header, index) {
      if (header) row[header] = values[r][index];
    });
    rows.push(row);
  }
  return rows;
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

function scoreHubReadLocalStudentRowsV2_(ss, sheetName, student) {
  const rows = readSheetObjects_(ss, sheetName);
  const sid = String(student.studentId || '').trim();
  const phone = scoreHubPhoneV2_(student.phone);
  return rows.filter(function(row) {
    const rowSid = String(row.studentId || '').trim();
    const rowPhone = scoreHubPhoneV2_(row.normalizedPhone);
    return (sid && rowSid === sid) || (phone && rowPhone === phone);
  });
}

function scoreHubMapWorkbookReadingRowV2_(row) {
  return {
    '제출시각': String(row.submittedAt || ''),
    '응시ID': String(row.attemptId || ''),
    '전화번호': String(row.normalizedPhone || ''),
    '학생이름': String(row.sourceName || ''),
    '반': String(row.sourceClass || ''),
    '교재': String(row.book || ''),
    '복습': String(row.review || ''),
    '평가영역': String(row.area || ''),
    '점수(100)': String(row.score || ''),
    '정답수': String(row.correct || ''),
    '전체문항': String(row.total || ''),
    '미응답': String(row.unanswered || ''),
    '응시시간(초)': String(row.durationSec || ''),
    '시간초과': String(row.timeout || ''),
    '답안JSON': String(row.answersJson || ''),
    'userAgent': String(row.userAgent || '')
  };
}

function scoreHubMapWorkbookListeningRowV2_(row) {
  return {
    '시작시각': String(row.startedAt || ''),
    '제출시각': String(row.submittedAt || ''),
    '응시ID': String(row.attemptId || ''),
    '전화번호': String(row.normalizedPhone || ''),
    '학생이름': String(row.sourceName || ''),
    '반': String(row.sourceClass || ''),
    '교재': String(row.book || ''),
    '복습': String(row.review || ''),
    '상태': String(row.status || ''),
    '점수(100)': String(row.score || ''),
    '정답수': String(row.correct || ''),
    '전체문항': String(row.total || ''),
    '답안JSON': String(row.answersJson || ''),
    '응시시간(초)': String(row.durationSec || ''),
    '기기ID': String(row.deviceId || ''),
    '정답버전': String(row.answerVersion || ''),
    'userAgent': String(row.userAgent || '')
  };
}

function scoreHubMasteryFromLocalV2_(rows) {
  return rows.map(function(row) {
    const passed = scoreHubBooleanV2_(row.passed);
    return {
      book: String(row.book || ''),
      lesson: String(row.lesson || ''),
      vocab: {
        bestScore: scoreHubNumberV2_(row.vocabBest),
        attempts: scoreHubNumberV2_(row.vocabAttempts) || 0
      },
      grammar: {
        bestScore: scoreHubNumberV2_(row.grammarBest),
        attempts: scoreHubNumberV2_(row.grammarAttempts) || 0
      },
      mixed: {
        bestScore: scoreHubNumberV2_(row.mixedBest),
        attempts: scoreHubNumberV2_(row.mixedAttempts) || 0
      },
      totalAttempts: scoreHubNumberV2_(row.totalAttempts) || 0,
      passed: passed,
      status: passed ? 'PASS' : 'RETRY'
    };
  }).sort(function(a, b) {
    const bookCmp = String(a.book).localeCompare(String(b.book));
    if (bookCmp !== 0) return bookCmp;
    return Number(a.lesson || 0) - Number(b.lesson || 0);
  });
}

function scoreHubSyncInfoV2_(ss) {
  const rows = readSheetObjects_(ss, 'SyncState');
  const out = {};
  rows.forEach(function(row) {
    const key = String(row.sourceType || '').trim();
    if (!key) return;
    out[key] = {
      status: String(row.status || ''),
      lastSuccessAt: String(row.lastSuccessAt || ''),
      lastRunAt: String(row.lastRunAt || '')
    };
  });
  return out;
}

function getScoreHubSourcesV2_(auth) {
  return {
    version: 'phase2-local-sync-20261003',
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

  // 핵심 변경: 모바일 원본이 아니라 V2 전용 교사 데이터 Spreadsheet를 우선 조회한다.
  const localSs = getTeacherDataSpreadsheet_(auth.teacher);

  const tests = scoreHubReadLocalStudentRowsV2_(localSs, 'LearningTestResults', student);
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
    return String(row.book || '') === 'TOPIK1' && String(row.testType || '') === 'collocation';
  });
  const topik1GrammarRows = tests.filter(function(row) {
    return String(row.book || '') === 'TOPIK1' && String(row.testType || '') === 'grammar';
  });

  const wbReading = scoreHubReadLocalStudentRowsV2_(localSs, 'WorkbookReadingResults', student)
    .map(scoreHubMapWorkbookReadingRowV2_);
  const wbListeningAll = scoreHubReadLocalStudentRowsV2_(localSs, 'WorkbookListeningResults', student)
    .map(scoreHubMapWorkbookListeningRowV2_);
  const wbListening = wbListeningAll.filter(function(row) {
    return String(row['상태'] || '') === 'SUBMITTED';
  });

  const progressRows = scoreHubReadLocalStudentRowsV2_(localSs, 'StudentLearningProgress', student);
  const progress = progressRows.length ? progressRows[0] : null;

  const masteryRows = scoreHubMasteryFromLocalV2_(
    scoreHubReadLocalStudentRowsV2_(localSs, 'ScoreHubMastery', student)
  );

  // TOPIK I 읽기 결과는 아직 V2 동기화 대상이 아니므로 기존 원본을 유지한다.
  const topik1ReadingSs = SpreadsheetApp.openById(SCORE_HUB_SOURCE_V2_.TOPIK1_READING_ID);
  const topik1Reading = scoreHubRowsByPhoneFromSpreadsheetV2_(
    topik1ReadingSs, 'All_Results', ['student_phone'], student.phone
  );

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
    version:'phase2-local-sync-20261003',
    sourceMode:'V2_LOCAL_SYNC',
    syncInfo:scoreHubSyncInfoV2_(localSs),
    student:student,
    progress:progress ? {
      firstLoginAt:String(progress.firstLoginAt || ''),
      recentLoginAt:String(progress.recentLoginAt || ''),
      elapsedText:String(progress.elapsedText || ''),
      startBook:String(progress.startBook || ''),
      startLesson:String(progress.startLesson || ''),
      currentBook:String(progress.currentBook || ''),
      currentLesson:String(progress.currentLesson || ''),
      passProgress:String(progress.passProgress || ''),
      vocabBest:scoreHubNumberV2_(progress.currentVocabBest),
      grammarBest:scoreHubNumberV2_(progress.currentGrammarBest),
      mixedBest:scoreHubNumberV2_(progress.currentMixedBest),
      currentState:String(progress.currentState || ''),
      nextStep:String(progress.nextStep || ''),
      topikCollocationState:String(progress.topikCollocationState || ''),
      topikGrammarState:String(progress.topikGrammarState || ''),
      recentActivityAt:String(progress.recentActivityAt || ''),
      recentTest:String(progress.recentTest || ''),
      recentScore:scoreHubNumberV2_(progress.recentScore),
      totalAttempts:scoreHubNumberV2_(progress.totalAttempts) || 0
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
      mastery:masteryRows
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
