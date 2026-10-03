const LEARNING_SYNC_V2_ = Object.freeze({
  LEGACY_STUDENT_MANAGEMENT_ID: '1Y5qoA0mQp-7EAQoXM7GsOF6y3MkdT5XhVEmjGjQc_TY',
  MOBILE_ID: '1y4xaZD8SQUVLztDhBSytvi-_naVCGYX5gTOyZqmhMUE',
  ACTIVE_CLASS_STATUS: '운영중',
  ACTIVE_STUDENT_STATUS: '재학',
  SHEETS: Object.freeze({
    TEST_RESULTS: 'LearningTestResults',
    WORKBOOK_READING: 'WorkbookReadingResults',
    WORKBOOK_LISTENING: 'WorkbookListeningResults',
    PROGRESS: 'StudentLearningProgress',
    MASTERY: 'ScoreHubMastery',
    SYNC_STATE: 'SyncState',
    SYNC_LOG: 'SyncLog'
  })
});

const LEARNING_SYNC_HEADERS_V2_ = Object.freeze({
  LearningTestResults: [
    'sourceKey','studentId','normalizedPhone','date','sourceName','sourceClass',
    'book','lesson','testType','bestScore','attemptsToday','bestCorrect','total',
    'bestTimeout','firstAt','bestAt','lastAt','status','sourceRow','checksum','syncedAt'
  ],
  WorkbookReadingResults: [
    'attemptId','studentId','normalizedPhone','submittedAt','sourceName','sourceClass',
    'book','review','area','score','correct','total','unanswered','durationSec',
    'timeout','answersJson','userAgent','sourceRow','checksum','syncedAt'
  ],
  WorkbookListeningResults: [
    'attemptId','studentId','normalizedPhone','startedAt','submittedAt','sourceName',
    'sourceClass','book','review','status','score','correct','total','answersJson',
    'durationSec','deviceId','answerVersion','userAgent','sourceRow','checksum','syncedAt'
  ],
  StudentLearningProgress: [
    'studentId','normalizedPhone','sourceName','active','sourceClass','firstLoginAt',
    'recentLoginAt','elapsedText','startBook','startLesson','currentBook','currentLesson',
    'passProgress','currentVocabBest','currentGrammarBest','currentMixedBest','currentState',
    'nextStep','topikCollocationState','topikGrammarState','recentActivityAt','recentTest',
    'recentScore','totalAttempts','sourceRow','checksum','syncedAt'
  ],
  ScoreHubMastery: [
    'studentId','normalizedPhone','book','lesson','vocabBest','grammarBest','mixedBest',
    'vocabAttempts','grammarAttempts','mixedAttempts','totalAttempts','passed',
    'lastActivityAt','syncedAt'
  ],
  SyncState: [
    'sourceType','lastSuccessAt','lastRunAt','status','scannedRows','matchedRows',
    'insertedRows','updatedRows','skippedRows','errorMessage'
  ],
  SyncLog: [
    'syncId','runType','sourceType','startedAt','finishedAt','status','scannedRows',
    'matchedRows','insertedRows','updatedRows','skippedRows','errorMessage'
  ]
});

function learningSyncPhoneV2_(value) {
  return String(value || '').replace(/\D/g, '');
}

function learningSyncTextV2_(value) {
  return String(value == null ? '' : value).trim();
}

function learningSyncNumberV2_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(value);
  return Number.isFinite(n) ? n : value;
}

function learningSyncBoolTextV2_(value) {
  const t = learningSyncTextV2_(value).toUpperCase();
  if (t === 'TRUE') return true;
  if (t === 'FALSE') return false;
  return value;
}

function learningSyncHexV2_(bytes) {
  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function learningSyncChecksumV2_(values) {
  const json = JSON.stringify(values || []);
  return learningSyncHexV2_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    json,
    Utilities.Charset.UTF_8
  ));
}

function learningSyncReadSheetV2_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    const error = new Error('필수 원본 시트를 찾을 수 없습니다: ' + sheetName);
    error.code = 'LEARNING_SYNC_SOURCE_SHEET_NOT_FOUND';
    throw error;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    return { sheet: sheet, headers: [], rows: [] };
  }

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(v) { return learningSyncTextV2_(v); });
  const rows = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row.some(function(v) { return learningSyncTextV2_(v) !== ''; })) continue;
    const obj = { __sourceRow: r + 1 };
    headers.forEach(function(header, index) {
      if (header) obj[header] = row[index];
    });
    rows.push(obj);
  }

  return { sheet: sheet, headers: headers, rows: rows };
}

function learningSyncEnsureSheetV2_(targetSs, sheetName) {
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];
  if (!headers) throw new Error('정의되지 않은 동기화 시트: ' + sheetName);

  let sheet = targetSs.getSheetByName(sheetName);
  if (!sheet) {
    sheet = targetSs.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .slice(0, headers.length)
    .map(function(v) { return learningSyncTextV2_(v); });

  const same = headers.every(function(h, i) { return current[i] === h; });
  if (!same) {
    const error = new Error(sheetName + ' 헤더가 예상 구조와 다릅니다. 자동 수정하지 않습니다.');
    error.code = 'LEARNING_SYNC_TARGET_HEADER_MISMATCH';
    throw error;
  }
  return sheet;
}

function learningSyncRowForHeadersV2_(headers, record) {
  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  });
}

function learningSyncGetTargetStudentMapV2_(targetSs) {
  const data = learningSyncReadSheetV2_(targetSs, 'Students').rows;
  const byPhone = {};
  data.forEach(function(row) {
    const phone = learningSyncPhoneV2_(row.phone);
    if (!phone) return;
    if (!byPhone[phone]) byPhone[phone] = [];
    byPhone[phone].push({
      studentId: learningSyncTextV2_(row.studentId),
      classId: learningSyncTextV2_(row.classId),
      name: learningSyncTextV2_(row.name),
      status: learningSyncTextV2_(row.status),
      phone: phone
    });
  });
  return byPhone;
}

function learningSyncMobilePhoneSetV2_(mobileSs) {
  const set = {};
  const specs = [
    ['TestResults', 'phone'],
    ['워크북읽기평가', '전화번호'],
    ['워크북듣기평가', '전화번호'],
    ['진도현황', '전화번호']
  ];

  specs.forEach(function(spec) {
    const rows = learningSyncReadSheetV2_(mobileSs, spec[0]).rows;
    rows.forEach(function(row) {
      const phone = learningSyncPhoneV2_(row[spec[1]]);
      if (phone) set[phone] = true;
    });
  });
  return set;
}

function learningSyncEligiblePhonesV2_(targetSs, legacySs, mobileSs) {
  const legacyClasses = learningSyncReadSheetV2_(legacySs, 'Classes').rows;
  const legacyStudents = learningSyncReadSheetV2_(legacySs, 'Students').rows;
  const mobilePhones = learningSyncMobilePhoneSetV2_(mobileSs);
  const targetByPhone = learningSyncGetTargetStudentMapV2_(targetSs);

  const activeClassIds = {};
  legacyClasses.forEach(function(row) {
    if (learningSyncTextV2_(row.status) === LEARNING_SYNC_V2_.ACTIVE_CLASS_STATUS) {
      activeClassIds[learningSyncTextV2_(row.classId)] = true;
    }
  });

  const eligible = {};
  const unresolved = [];

  legacyStudents.forEach(function(row) {
    const classId = learningSyncTextV2_(row.classId);
    const status = learningSyncTextV2_(row.status);
    const phone = learningSyncPhoneV2_(row.phone);

    if (!activeClassIds[classId]) return;
    if (status !== LEARNING_SYNC_V2_.ACTIVE_STUDENT_STATUS) return;
    if (!phone || !mobilePhones[phone]) return;

    const targetRows = targetByPhone[phone] || [];
    if (targetRows.length !== 1) {
      unresolved.push({
        phone: phone,
        legacyStudentId: learningSyncTextV2_(row.studentId),
        legacyName: learningSyncTextV2_(row.name),
        targetMatchCount: targetRows.length
      });
      return;
    }

    eligible[phone] = {
      phone: phone,
      studentId: targetRows[0].studentId,
      targetClassId: targetRows[0].classId,
      targetName: targetRows[0].name,
      legacyStudentId: learningSyncTextV2_(row.studentId),
      legacyClassId: classId,
      legacyName: learningSyncTextV2_(row.name)
    };
  });

  return { eligibleByPhone: eligible, unresolved: unresolved };
}

function learningSyncCollectSourceV2_(targetSs, legacySs, mobileSs) {
  const eligibility = learningSyncEligiblePhonesV2_(targetSs, legacySs, mobileSs);
  if (eligibility.unresolved.length) {
    const error = new Error('연결 대상 중 V2 전화번호 매핑이 1개가 아닌 학생이 있습니다. 먼저 확인해 주세요.');
    error.code = 'LEARNING_SYNC_TARGET_PHONE_MAPPING_AMBIGUOUS';
    error.details = eligibility.unresolved;
    throw error;
  }

  const eligible = eligibility.eligibleByPhone;
  const eligiblePhones = Object.keys(eligible);
  const now = new Date();

  const testsRaw = learningSyncReadSheetV2_(mobileSs, 'TestResults').rows;
  const readingRaw = learningSyncReadSheetV2_(mobileSs, '워크북읽기평가').rows;
  const listeningRaw = learningSyncReadSheetV2_(mobileSs, '워크북듣기평가').rows;
  const progressRaw = learningSyncReadSheetV2_(mobileSs, '진도현황').rows;

  const tests = [];
  testsRaw.forEach(function(row) {
    const phone = learningSyncPhoneV2_(row.phone);
    const link = eligible[phone];
    if (!link) return;

    const sourceKey = [
      'TEST', phone, learningSyncTextV2_(row.date), learningSyncTextV2_(row.book),
      learningSyncTextV2_(row.lesson), learningSyncTextV2_(row.testType)
    ].join('|');
    const checksumValues = [
      row.date,row.phone,row.name,row.klass,row.book,row.lesson,row.testType,row.bestScore,
      row.attemptsToday,row.bestCorrect,row.total,row.bestTimeout,row.firstAt,row.bestAt,
      row.lastAt,row.status
    ];

    tests.push({
      sourceKey: sourceKey,
      studentId: link.studentId,
      normalizedPhone: phone,
      date: row.date || '',
      sourceName: row.name || '',
      sourceClass: row.klass || '',
      book: row.book || '',
      lesson: row.lesson || '',
      testType: row.testType || '',
      bestScore: learningSyncNumberV2_(row.bestScore),
      attemptsToday: learningSyncNumberV2_(row.attemptsToday),
      bestCorrect: learningSyncNumberV2_(row.bestCorrect),
      total: learningSyncNumberV2_(row.total),
      bestTimeout: learningSyncNumberV2_(row.bestTimeout),
      firstAt: row.firstAt || '',
      bestAt: row.bestAt || '',
      lastAt: row.lastAt || '',
      status: row.status || '',
      sourceRow: row.__sourceRow,
      checksum: learningSyncChecksumV2_(checksumValues),
      syncedAt: now
    });
  });

  const reading = [];
  readingRaw.forEach(function(row) {
    const phone = learningSyncPhoneV2_(row['전화번호']);
    const link = eligible[phone];
    if (!link) return;
    const attemptId = learningSyncTextV2_(row['응시ID']);
    if (!attemptId) return;

    const checksumValues = [
      row['제출시각'],row['응시ID'],row['전화번호'],row['학생이름'],row['반'],row['교재'],
      row['복습'],row['평가영역'],row['점수(100)'],row['정답수'],row['전체문항'],row['미응답'],
      row['응시시간(초)'],row['시간초과'],row['답안JSON'],row['userAgent']
    ];

    reading.push({
      attemptId: attemptId,
      studentId: link.studentId,
      normalizedPhone: phone,
      submittedAt: row['제출시각'] || '',
      sourceName: row['학생이름'] || '',
      sourceClass: row['반'] || '',
      book: row['교재'] || '',
      review: row['복습'] || '',
      area: row['평가영역'] || '',
      score: learningSyncNumberV2_(row['점수(100)']),
      correct: learningSyncNumberV2_(row['정답수']),
      total: learningSyncNumberV2_(row['전체문항']),
      unanswered: learningSyncNumberV2_(row['미응답']),
      durationSec: learningSyncNumberV2_(row['응시시간(초)']),
      timeout: learningSyncBoolTextV2_(row['시간초과']),
      answersJson: row['답안JSON'] || '',
      userAgent: row['userAgent'] || '',
      sourceRow: row.__sourceRow,
      checksum: learningSyncChecksumV2_(checksumValues),
      syncedAt: now
    });
  });

  const listening = [];
  listeningRaw.forEach(function(row) {
    const phone = learningSyncPhoneV2_(row['전화번호']);
    const link = eligible[phone];
    if (!link) return;
    const attemptId = learningSyncTextV2_(row['응시ID']);
    if (!attemptId) return;

    const checksumValues = [
      row['시작시각'],row['제출시각'],row['응시ID'],row['전화번호'],row['학생이름'],row['반'],
      row['교재'],row['복습'],row['상태'],row['점수(100)'],row['정답수'],row['전체문항'],
      row['답안JSON'],row['응시시간(초)'],row['기기ID'],row['정답버전'],row['userAgent']
    ];

    listening.push({
      attemptId: attemptId,
      studentId: link.studentId,
      normalizedPhone: phone,
      startedAt: row['시작시각'] || '',
      submittedAt: row['제출시각'] || '',
      sourceName: row['학생이름'] || '',
      sourceClass: row['반'] || '',
      book: row['교재'] || '',
      review: row['복습'] || '',
      status: row['상태'] || '',
      score: learningSyncNumberV2_(row['점수(100)']),
      correct: learningSyncNumberV2_(row['정답수']),
      total: learningSyncNumberV2_(row['전체문항']),
      answersJson: row['답안JSON'] || '',
      durationSec: learningSyncNumberV2_(row['응시시간(초)']),
      deviceId: row['기기ID'] || '',
      answerVersion: row['정답버전'] || '',
      userAgent: row['userAgent'] || '',
      sourceRow: row.__sourceRow,
      checksum: learningSyncChecksumV2_(checksumValues),
      syncedAt: now
    });
  });

  const progress = [];
  progressRaw.forEach(function(row) {
    const phone = learningSyncPhoneV2_(row['전화번호']);
    const link = eligible[phone];
    if (!link) return;

    const checksumValues = [
      row['전화번호'],row['학생이름'],row['사용여부'],row['반'],row['최초접속'],row['최근접속'],
      row['접속경과'],row['시작교재'],row['시작과'],row['현재교재'],row['현재과'],row['통과현황'],
      row['어휘최고'],row['문법최고'],row['종합최고'],row['현재상태'],row['다음단계'],row['TOPIK연어'],
      row['TOPIK문법'],row['최근활동'],row['최근시험'],row['최근점수'],row['총응시']
    ];

    progress.push({
      studentId: link.studentId,
      normalizedPhone: phone,
      sourceName: row['학생이름'] || '',
      active: learningSyncBoolTextV2_(row['사용여부']),
      sourceClass: row['반'] || '',
      firstLoginAt: row['최초접속'] || '',
      recentLoginAt: row['최근접속'] || '',
      elapsedText: row['접속경과'] || '',
      startBook: row['시작교재'] || '',
      startLesson: row['시작과'] || '',
      currentBook: row['현재교재'] || '',
      currentLesson: row['현재과'] || '',
      passProgress: row['통과현황'] || '',
      currentVocabBest: learningSyncNumberV2_(row['어휘최고']),
      currentGrammarBest: learningSyncNumberV2_(row['문법최고']),
      currentMixedBest: learningSyncNumberV2_(row['종합최고']),
      currentState: row['현재상태'] || '',
      nextStep: row['다음단계'] || '',
      topikCollocationState: row['TOPIK연어'] || '',
      topikGrammarState: row['TOPIK문법'] || '',
      recentActivityAt: row['최근활동'] || '',
      recentTest: row['최근시험'] || '',
      recentScore: learningSyncNumberV2_(row['최근점수']),
      totalAttempts: learningSyncNumberV2_(row['총응시']),
      sourceRow: row.__sourceRow,
      checksum: learningSyncChecksumV2_(checksumValues),
      syncedAt: now
    });
  });

  const mastery = learningSyncBuildMasteryV2_(tests, now);

  return {
    eligiblePhones: eligiblePhones,
    tests: tests,
    reading: reading,
    listening: listening,
    progress: progress,
    mastery: mastery,
    scanned: {
      TestResults: testsRaw.length,
      WorkbookReading: readingRaw.length,
      WorkbookListening: listeningRaw.length,
      Progress: progressRaw.length
    }
  };
}

function learningSyncBuildMasteryV2_(tests, now) {
  const map = {};

  tests.forEach(function(row) {
    const book = learningSyncTextV2_(row.book);
    const lesson = learningSyncTextV2_(row.lesson);
    const testType = learningSyncTextV2_(row.testType);
    if (!/^SNU-/.test(book)) return;
    if (['vocab','grammar','mixed'].indexOf(testType) < 0) return;

    const key = [row.studentId, row.normalizedPhone, book, lesson].join('|');
    if (!map[key]) {
      map[key] = {
        studentId: row.studentId,
        normalizedPhone: row.normalizedPhone,
        book: book,
        lesson: lesson,
        vocabBest: '', grammarBest: '', mixedBest: '',
        vocabAttempts: 0, grammarAttempts: 0, mixedAttempts: 0,
        totalAttempts: 0,
        passed: false,
        lastActivityAt: '',
        syncedAt: now
      };
    }

    const item = map[key];
    const score = Number(row.bestScore);
    const attempts = Number(row.attemptsToday) || 0;
    if (testType === 'vocab') {
      if (Number.isFinite(score)) item.vocabBest = item.vocabBest === '' ? score : Math.max(Number(item.vocabBest), score);
      item.vocabAttempts += attempts;
    } else if (testType === 'grammar') {
      if (Number.isFinite(score)) item.grammarBest = item.grammarBest === '' ? score : Math.max(Number(item.grammarBest), score);
      item.grammarAttempts += attempts;
    } else if (testType === 'mixed') {
      if (Number.isFinite(score)) item.mixedBest = item.mixedBest === '' ? score : Math.max(Number(item.mixedBest), score);
      item.mixedAttempts += attempts;
    }

    item.totalAttempts = item.vocabAttempts + item.grammarAttempts + item.mixedAttempts;
    item.lastActivityAt = row.lastAt || row.bestAt || row.date || item.lastActivityAt;
  });

  return Object.keys(map).map(function(key) {
    const item = map[key];
    item.passed = Number(item.vocabBest) >= 90 && Number(item.grammarBest) >= 90 && Number(item.mixedBest) >= 90;
    return item;
  });
}

function learningSyncUpsertV2_(targetSs, sheetName, keyHeader, records) {
  const sheet = learningSyncEnsureSheetV2_(targetSs, sheetName);
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];
  const keyIndex = headers.indexOf(keyHeader);
  if (keyIndex < 0) throw new Error(sheetName + ' 키 헤더를 찾을 수 없습니다: ' + keyHeader);

  const checksumIndex = headers.indexOf('checksum');
  const lastRow = sheet.getLastRow();
  const existing = {};

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    values.forEach(function(row, index) {
      const key = learningSyncTextV2_(row[keyIndex]);
      if (!key) return;
      existing[key] = {
        rowNumber: index + 2,
        checksum: checksumIndex >= 0 ? learningSyncTextV2_(row[checksumIndex]) : ''
      };
    });
  }

  const appendRows = [];
  let updated = 0;
  let skipped = 0;

  records.forEach(function(record) {
    const key = learningSyncTextV2_(record[keyHeader]);
    if (!key) return;
    const found = existing[key];

    if (!found) {
      appendRows.push(learningSyncRowForHeadersV2_(headers, record));
      return;
    }

    const incomingChecksum = checksumIndex >= 0 ? learningSyncTextV2_(record.checksum) : '';
    if (checksumIndex >= 0 && found.checksum === incomingChecksum) {
      skipped++;
      return;
    }

    sheet.getRange(found.rowNumber, 1, 1, headers.length)
      .setValues([learningSyncRowForHeadersV2_(headers, record)]);
    updated++;
  });

  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length)
      .setValues(appendRows);
  }

  return { inserted: appendRows.length, updated: updated, skipped: skipped };
}

function learningSyncReplaceMasteryV2_(targetSs, records) {
  const sheetName = LEARNING_SYNC_V2_.SHEETS.MASTERY;
  const sheet = learningSyncEnsureSheetV2_(targetSs, sheetName);
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
  }
  if (records.length) {
    const values = records.map(function(record) {
      return learningSyncRowForHeadersV2_(headers, record);
    });
    sheet.getRange(2, 1, values.length, headers.length).setValues(values);
  }
  return { written: records.length };
}

function learningSyncUpsertProgressV2_(targetSs, records) {
  return learningSyncUpsertV2_(
    targetSs,
    LEARNING_SYNC_V2_.SHEETS.PROGRESS,
    'normalizedPhone',
    records
  );
}

function learningSyncWriteStateV2_(targetSs, sourceType, stats, status, errorMessage) {
  const sheetName = LEARNING_SYNC_V2_.SHEETS.SYNC_STATE;
  const sheet = learningSyncEnsureSheetV2_(targetSs, sheetName);
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];
  const now = new Date();
  const lastRow = sheet.getLastRow();
  let rowNumber = -1;

  if (lastRow >= 2) {
    const sourceValues = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    for (let i = 0; i < sourceValues.length; i++) {
      if (learningSyncTextV2_(sourceValues[i][0]) === sourceType) {
        rowNumber = i + 2;
        break;
      }
    }
  }

  const record = {
    sourceType: sourceType,
    lastSuccessAt: status === 'SUCCESS' ? now : '',
    lastRunAt: now,
    status: status,
    scannedRows: stats.scannedRows || 0,
    matchedRows: stats.matchedRows || 0,
    insertedRows: stats.insertedRows || 0,
    updatedRows: stats.updatedRows || 0,
    skippedRows: stats.skippedRows || 0,
    errorMessage: errorMessage || ''
  };

  if (rowNumber > 0 && status !== 'SUCCESS') {
    const current = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const lastSuccessIndex = headers.indexOf('lastSuccessAt');
    if (lastSuccessIndex >= 0) record.lastSuccessAt = current[lastSuccessIndex] || '';
  }

  const values = learningSyncRowForHeadersV2_(headers, record);
  if (rowNumber > 0) sheet.getRange(rowNumber, 1, 1, headers.length).setValues([values]);
  else sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length).setValues([values]);
}

function learningSyncAppendLogV2_(targetSs, runType, sourceType, startedAt, finishedAt, status, stats, errorMessage) {
  const sheetName = LEARNING_SYNC_V2_.SHEETS.SYNC_LOG;
  const sheet = learningSyncEnsureSheetV2_(targetSs, sheetName);
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];
  const record = {
    syncId: 'SYNC-' + Utilities.getUuid(),
    runType: runType,
    sourceType: sourceType,
    startedAt: startedAt,
    finishedAt: finishedAt,
    status: status,
    scannedRows: stats.scannedRows || 0,
    matchedRows: stats.matchedRows || 0,
    insertedRows: stats.insertedRows || 0,
    updatedRows: stats.updatedRows || 0,
    skippedRows: stats.skippedRows || 0,
    errorMessage: errorMessage || ''
  };
  sheet.getRange(sheet.getLastRow() + 1, 1, 1, headers.length)
    .setValues([learningSyncRowForHeadersV2_(headers, record)]);
}

function previewInitialMobileLearningSyncV2() {
  const targetSs = SpreadsheetApp.openById('1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q');
  const legacySs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.LEGACY_STUDENT_MANAGEMENT_ID);
  const mobileSs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.MOBILE_ID);
  const data = learningSyncCollectSourceV2_(targetSs, legacySs, mobileSs);

  const result = {
    mode: 'PREVIEW_ONLY',
    eligibleStudentCount: data.eligiblePhones.length,
    scanned: data.scanned,
    matched: {
      LearningTestResults: data.tests.length,
      WorkbookReadingResults: data.reading.length,
      WorkbookListeningResults: data.listening.length,
      StudentLearningProgress: data.progress.length,
      ScoreHubMastery: data.mastery.length
    },
    note: '이 함수는 시트를 생성하거나 수정하지 않습니다.'
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function applyInitialMobileLearningSyncV2() {
  const startedAt = new Date();
  const targetSs = SpreadsheetApp.openById('1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q');
  const legacySs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.LEGACY_STUDENT_MANAGEMENT_ID);
  const mobileSs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.MOBILE_ID);

  let data;
  try {
    data = learningSyncCollectSourceV2_(targetSs, legacySs, mobileSs);

    const testStats = learningSyncUpsertV2_(targetSs, LEARNING_SYNC_V2_.SHEETS.TEST_RESULTS, 'sourceKey', data.tests);
    const readingStats = learningSyncUpsertV2_(targetSs, LEARNING_SYNC_V2_.SHEETS.WORKBOOK_READING, 'attemptId', data.reading);
    const listeningStats = learningSyncUpsertV2_(targetSs, LEARNING_SYNC_V2_.SHEETS.WORKBOOK_LISTENING, 'attemptId', data.listening);
    const progressStats = learningSyncUpsertProgressV2_(targetSs, data.progress);
    const masteryStats = learningSyncReplaceMasteryV2_(targetSs, data.mastery);

    const finishedAt = new Date();
    const sourceStats = {
      TestResults: { scannedRows:data.scanned.TestResults, matchedRows:data.tests.length, insertedRows:testStats.inserted, updatedRows:testStats.updated, skippedRows:testStats.skipped },
      WorkbookReading: { scannedRows:data.scanned.WorkbookReading, matchedRows:data.reading.length, insertedRows:readingStats.inserted, updatedRows:readingStats.updated, skippedRows:readingStats.skipped },
      WorkbookListening: { scannedRows:data.scanned.WorkbookListening, matchedRows:data.listening.length, insertedRows:listeningStats.inserted, updatedRows:listeningStats.updated, skippedRows:listeningStats.skipped },
      Progress: { scannedRows:data.scanned.Progress, matchedRows:data.progress.length, insertedRows:progressStats.inserted, updatedRows:progressStats.updated, skippedRows:progressStats.skipped }
    };

    Object.keys(sourceStats).forEach(function(sourceType) {
      learningSyncWriteStateV2_(targetSs, sourceType, sourceStats[sourceType], 'SUCCESS', '');
      learningSyncAppendLogV2_(targetSs, 'INITIAL_MANUAL', sourceType, startedAt, finishedAt, 'SUCCESS', sourceStats[sourceType], '');
    });

    const result = {
      mode: 'APPLY_INITIAL',
      eligibleStudentCount: data.eligiblePhones.length,
      results: {
        LearningTestResults: testStats,
        WorkbookReadingResults: readingStats,
        WorkbookListeningResults: listeningStats,
        StudentLearningProgress: progressStats,
        ScoreHubMastery: masteryStats
      },
      sheetsCreatedOrVerified: [
        LEARNING_SYNC_V2_.SHEETS.TEST_RESULTS,
        LEARNING_SYNC_V2_.SHEETS.WORKBOOK_READING,
        LEARNING_SYNC_V2_.SHEETS.WORKBOOK_LISTENING,
        LEARNING_SYNC_V2_.SHEETS.PROGRESS,
        LEARNING_SYNC_V2_.SHEETS.MASTERY,
        LEARNING_SYNC_V2_.SHEETS.SYNC_STATE,
        LEARNING_SYNC_V2_.SHEETS.SYNC_LOG
      ]
    };

    console.log(JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    const finishedAt = new Date();
    const message = err && err.message ? err.message : String(err);
    try {
      learningSyncAppendLogV2_(targetSs, 'INITIAL_MANUAL', 'ALL', startedAt, finishedAt, 'FAILED', {}, message);
    } catch (ignore) {}
    throw err;
  }
}

function learningSyncPreviewUpsertStatsV2_(targetSs, sheetName, keyHeader, records) {
  const sheet = learningSyncEnsureSheetV2_(targetSs, sheetName);
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];
  const keyIndex = headers.indexOf(keyHeader);
  if (keyIndex < 0) throw new Error(sheetName + ' 키 헤더를 찾을 수 없습니다: ' + keyHeader);

  const checksumIndex = headers.indexOf('checksum');
  const existing = {};
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    values.forEach(function(row) {
      const key = learningSyncTextV2_(row[keyIndex]);
      if (!key) return;
      existing[key] = checksumIndex >= 0 ? learningSyncTextV2_(row[checksumIndex]) : '';
    });
  }

  let insert = 0;
  let update = 0;
  let skip = 0;

  records.forEach(function(record) {
    const key = learningSyncTextV2_(record[keyHeader]);
    if (!key) return;
    if (!Object.prototype.hasOwnProperty.call(existing, key)) {
      insert++;
      return;
    }
    const incomingChecksum = checksumIndex >= 0 ? learningSyncTextV2_(record.checksum) : '';
    if (checksumIndex >= 0 && existing[key] === incomingChecksum) skip++;
    else update++;
  });

  return { wouldInsert: insert, wouldUpdate: update, wouldSkip: skip };
}

function learningSyncMasteryComparableV2_(records) {
  return records.map(function(record) {
    return [
      learningSyncTextV2_(record.studentId),
      learningSyncPhoneV2_(record.normalizedPhone),
      learningSyncTextV2_(record.book),
      learningSyncTextV2_(record.lesson),
      learningSyncTextV2_(record.vocabBest),
      learningSyncTextV2_(record.grammarBest),
      learningSyncTextV2_(record.mixedBest),
      learningSyncTextV2_(record.vocabAttempts),
      learningSyncTextV2_(record.grammarAttempts),
      learningSyncTextV2_(record.mixedAttempts),
      learningSyncTextV2_(record.totalAttempts),
      String(learningSyncBoolTextV2_(record.passed)).toUpperCase(),
      learningSyncTextV2_(record.lastActivityAt)
    ].join('|');
  }).sort();
}

function learningSyncCurrentMasteryComparableV2_(targetSs) {
  const sheetName = LEARNING_SYNC_V2_.SHEETS.MASTERY;
  const sheet = learningSyncEnsureSheetV2_(targetSs, sheetName);
  const headers = LEARNING_SYNC_HEADERS_V2_[sheetName];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
  const records = values.filter(function(row) {
    return row.some(function(v) { return learningSyncTextV2_(v) !== ''; });
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(header, index) {
      if (header) obj[header] = row[index];
    });
    return obj;
  });

  return learningSyncMasteryComparableV2_(records);
}

function learningSyncMasteryChangedV2_(targetSs, incomingRecords) {
  const current = learningSyncCurrentMasteryComparableV2_(targetSs);
  const incoming = learningSyncMasteryComparableV2_(incomingRecords);
  if (current.length !== incoming.length) return true;
  for (let i = 0; i < current.length; i++) {
    if (current[i] !== incoming[i]) return true;
  }
  return false;
}

function learningSyncApplyMasteryIfChangedV2_(targetSs, incomingRecords) {
  if (!learningSyncMasteryChangedV2_(targetSs, incomingRecords)) {
    return { written: 0, skippedUnchanged: incomingRecords.length };
  }
  const result = learningSyncReplaceMasteryV2_(targetSs, incomingRecords);
  result.skippedUnchanged = 0;
  return result;
}

function previewIncrementalMobileLearningSyncV2() {
  const targetSs = SpreadsheetApp.openById('1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q');
  const legacySs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.LEGACY_STUDENT_MANAGEMENT_ID);
  const mobileSs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.MOBILE_ID);
  const data = learningSyncCollectSourceV2_(targetSs, legacySs, mobileSs);

  const testStats = learningSyncPreviewUpsertStatsV2_(
    targetSs, LEARNING_SYNC_V2_.SHEETS.TEST_RESULTS, 'sourceKey', data.tests
  );
  const readingStats = learningSyncPreviewUpsertStatsV2_(
    targetSs, LEARNING_SYNC_V2_.SHEETS.WORKBOOK_READING, 'attemptId', data.reading
  );
  const listeningStats = learningSyncPreviewUpsertStatsV2_(
    targetSs, LEARNING_SYNC_V2_.SHEETS.WORKBOOK_LISTENING, 'attemptId', data.listening
  );
  const progressStats = learningSyncPreviewUpsertStatsV2_(
    targetSs, LEARNING_SYNC_V2_.SHEETS.PROGRESS, 'normalizedPhone', data.progress
  );
  const masteryChanged = learningSyncMasteryChangedV2_(targetSs, data.mastery);

  const result = {
    mode: 'PREVIEW_INCREMENTAL',
    eligibleStudentCount: data.eligiblePhones.length,
    scanned: data.scanned,
    changes: {
      LearningTestResults: testStats,
      WorkbookReadingResults: readingStats,
      WorkbookListeningResults: listeningStats,
      StudentLearningProgress: progressStats,
      ScoreHubMastery: {
        wouldRewrite: masteryChanged,
        incomingRows: data.mastery.length
      }
    },
    note: '이 함수는 데이터를 수정하지 않습니다. 초기 동기화 직후라면 insert/update는 0이고 기존 행은 skip이어야 정상입니다.'
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}

function learningSyncRunIncrementalV2_(runType) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    const error = new Error('다른 학습자료 동기화가 실행 중입니다. 잠시 후 다시 실행해 주세요.');
    error.code = 'LEARNING_SYNC_ALREADY_RUNNING';
    throw error;
  }

  const normalizedRunType = String(runType || 'INCREMENTAL_MANUAL').trim() || 'INCREMENTAL_MANUAL';
  const startedAt = new Date();
  const targetSs = SpreadsheetApp.openById('1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q');

  try {
    const legacySs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.LEGACY_STUDENT_MANAGEMENT_ID);
    const mobileSs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.MOBILE_ID);
    const data = learningSyncCollectSourceV2_(targetSs, legacySs, mobileSs);

    const testStats = learningSyncUpsertV2_(
      targetSs, LEARNING_SYNC_V2_.SHEETS.TEST_RESULTS, 'sourceKey', data.tests
    );
    const readingStats = learningSyncUpsertV2_(
      targetSs, LEARNING_SYNC_V2_.SHEETS.WORKBOOK_READING, 'attemptId', data.reading
    );
    const listeningStats = learningSyncUpsertV2_(
      targetSs, LEARNING_SYNC_V2_.SHEETS.WORKBOOK_LISTENING, 'attemptId', data.listening
    );
    const progressStats = learningSyncUpsertProgressV2_(targetSs, data.progress);
    const masteryStats = learningSyncApplyMasteryIfChangedV2_(targetSs, data.mastery);

    const finishedAt = new Date();
    const sourceStats = {
      TestResults: {
        scannedRows: data.scanned.TestResults,
        matchedRows: data.tests.length,
        insertedRows: testStats.inserted,
        updatedRows: testStats.updated,
        skippedRows: testStats.skipped
      },
      WorkbookReading: {
        scannedRows: data.scanned.WorkbookReading,
        matchedRows: data.reading.length,
        insertedRows: readingStats.inserted,
        updatedRows: readingStats.updated,
        skippedRows: readingStats.skipped
      },
      WorkbookListening: {
        scannedRows: data.scanned.WorkbookListening,
        matchedRows: data.listening.length,
        insertedRows: listeningStats.inserted,
        updatedRows: listeningStats.updated,
        skippedRows: listeningStats.skipped
      },
      Progress: {
        scannedRows: data.scanned.Progress,
        matchedRows: data.progress.length,
        insertedRows: progressStats.inserted,
        updatedRows: progressStats.updated,
        skippedRows: progressStats.skipped
      }
    };

    Object.keys(sourceStats).forEach(function(sourceType) {
      learningSyncWriteStateV2_(targetSs, sourceType, sourceStats[sourceType], 'SUCCESS', '');
      learningSyncAppendLogV2_(
        targetSs, normalizedRunType, sourceType, startedAt, finishedAt,
        'SUCCESS', sourceStats[sourceType], ''
      );
    });

    const result = {
      mode: normalizedRunType === 'INCREMENTAL_AUTO' ? 'APPLY_INCREMENTAL_AUTO' : 'APPLY_INCREMENTAL',
      runType: normalizedRunType,
      eligibleStudentCount: data.eligiblePhones.length,
      results: {
        LearningTestResults: testStats,
        WorkbookReadingResults: readingStats,
        WorkbookListeningResults: listeningStats,
        StudentLearningProgress: progressStats,
        ScoreHubMastery: masteryStats
      }
    };

    console.log(JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    const finishedAt = new Date();
    const message = err && err.message ? err.message : String(err);
    try {
      learningSyncAppendLogV2_(
        targetSs, normalizedRunType, 'ALL', startedAt, finishedAt,
        'FAILED', {}, message
      );
    } catch (ignore) {}
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function applyIncrementalMobileLearningSyncV2() {
  return learningSyncRunIncrementalV2_('INCREMENTAL_MANUAL');
}

function runAutomaticMobileLearningSyncV2() {
  return learningSyncRunIncrementalV2_('INCREMENTAL_AUTO');
}

function learningSyncDeleteAutoTriggersV2_() {
  const handler = 'runAutomaticMobileLearningSyncV2';
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
      deleted += 1;
    }
  });
  return deleted;
}

function installMobileLearningAutoSync5MinV2() {
  const deletedExisting = learningSyncDeleteAutoTriggersV2_();
  const trigger = ScriptApp.newTrigger('runAutomaticMobileLearningSyncV2')
    .timeBased()
    .everyMinutes(5)
    .create();

  const result = {
    installed: true,
    handler: 'runAutomaticMobileLearningSyncV2',
    everyMinutes: 5,
    deletedExisting: deletedExisting,
    triggerId: trigger.getUniqueId()
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function removeMobileLearningAutoSyncV2() {
  const deleted = learningSyncDeleteAutoTriggersV2_();
  const result = {
    removed: true,
    handler: 'runAutomaticMobileLearningSyncV2',
    deleted: deleted
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function getMobileLearningAutoSyncStatusV2() {
  const handler = 'runAutomaticMobileLearningSyncV2';
  const matches = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return trigger.getHandlerFunction() === handler;
  });
  const result = {
    enabled: matches.length > 0,
    handler: handler,
    triggerCount: matches.length,
    triggerIds: matches.map(function(trigger) { return trigger.getUniqueId(); })
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}
