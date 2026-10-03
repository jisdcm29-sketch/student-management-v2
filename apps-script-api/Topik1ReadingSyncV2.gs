const TOPIK1_READING_SYNC_V2_ = Object.freeze({
  SOURCE_ID: '18HXty992Riii2-csrB2aHpQ7vVt8qD2NOWFMp1yG63M',
  SOURCE_SHEET: 'All_Results',
  TARGET_SHEET: 'Topik1ReadingResults',
  SOURCE_TYPE: 'Topik1Reading',
  AUTO_HANDLER: 'runAutomaticTopik1ReadingSyncV2'
});

const TOPIK1_READING_SYNC_HEADERS_V2_ = Object.freeze([
  'attempt_id','studentId','normalizedPhone','recorded_at','student_phone','student_name',
  'result_type','generated_exam_mode','generated_exam_round','generated_exam_label',
  'test_name','test_scope','started_at','submitted_at','duration_seconds','total_questions',
  'answered_count','correct_count','wrong_count','unanswered_count','earned_points',
  'total_possible_points','section_score_100','source_rounds','correct_question_numbers',
  'wrong_question_numbers','unanswered_question_numbers','client_version','sourceRow',
  'checksum','syncedAt'
]);

function topik1ReadingSyncEnsureTargetSheetV2_(targetSs) {
  let sheet = targetSs.getSheetByName(TOPIK1_READING_SYNC_V2_.TARGET_SHEET);
  if (!sheet) {
    sheet = targetSs.insertSheet(TOPIK1_READING_SYNC_V2_.TARGET_SHEET);
    sheet.getRange(1, 1, 1, TOPIK1_READING_SYNC_HEADERS_V2_.length)
      .setValues([TOPIK1_READING_SYNC_HEADERS_V2_]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  const lastColumn = Math.max(sheet.getLastColumn(), TOPIK1_READING_SYNC_HEADERS_V2_.length);
  const current = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .slice(0, TOPIK1_READING_SYNC_HEADERS_V2_.length)
    .map(function(v) { return learningSyncTextV2_(v); });
  const same = TOPIK1_READING_SYNC_HEADERS_V2_.every(function(h, i) { return current[i] === h; });
  if (!same) {
    const error = new Error(TOPIK1_READING_SYNC_V2_.TARGET_SHEET + ' 헤더가 예상 구조와 다릅니다. 자동 수정하지 않습니다.');
    error.code = 'TOPIK1_READING_SYNC_TARGET_HEADER_MISMATCH';
    throw error;
  }
  return sheet;
}

function topik1ReadingSyncRowForHeadersV2_(record) {
  return TOPIK1_READING_SYNC_HEADERS_V2_.map(function(header) {
    return Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '';
  });
}

function topik1ReadingSyncCollectV2_() {
  const targetSs = SpreadsheetApp.openById('1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q');
  const legacySs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.LEGACY_STUDENT_MANAGEMENT_ID);
  const mobileSs = SpreadsheetApp.openById(LEARNING_SYNC_V2_.MOBILE_ID);
  const sourceSs = SpreadsheetApp.openById(TOPIK1_READING_SYNC_V2_.SOURCE_ID);

  const eligibility = learningSyncEligiblePhonesV2_(targetSs, legacySs, mobileSs);
  if (eligibility.unresolved.length) {
    const error = new Error('TOPIK I 읽기 연결 대상 중 V2 전화번호 매핑이 1개가 아닌 학생이 있습니다.');
    error.code = 'TOPIK1_READING_SYNC_TARGET_PHONE_MAPPING_AMBIGUOUS';
    error.details = eligibility.unresolved;
    throw error;
  }

  const eligible = eligibility.eligibleByPhone;
  const source = learningSyncReadSheetV2_(sourceSs, TOPIK1_READING_SYNC_V2_.SOURCE_SHEET);
  const now = new Date();
  const records = [];

  source.rows.forEach(function(row) {
    const phone = learningSyncPhoneV2_(row.student_phone);
    const link = eligible[phone];
    if (!link) return;

    const attemptId = learningSyncTextV2_(row.attempt_id);
    if (!attemptId) return;

    const checksumValues = [
      row.recorded_at,row.attempt_id,row.student_phone,row.student_name,row.result_type,
      row.generated_exam_mode,row.generated_exam_round,row.generated_exam_label,row.test_name,
      row.test_scope,row.started_at,row.submitted_at,row.duration_seconds,row.total_questions,
      row.answered_count,row.correct_count,row.wrong_count,row.unanswered_count,row.earned_points,
      row.total_possible_points,row.section_score_100,row.source_rounds,row.correct_question_numbers,
      row.wrong_question_numbers,row.unanswered_question_numbers,row.client_version
    ];

    records.push({
      attempt_id: attemptId,
      studentId: link.studentId,
      normalizedPhone: phone,
      recorded_at: row.recorded_at || '',
      student_phone: row.student_phone || '',
      student_name: row.student_name || '',
      result_type: row.result_type || '',
      generated_exam_mode: row.generated_exam_mode || '',
      generated_exam_round: row.generated_exam_round || '',
      generated_exam_label: row.generated_exam_label || '',
      test_name: row.test_name || '',
      test_scope: row.test_scope || '',
      started_at: row.started_at || '',
      submitted_at: row.submitted_at || '',
      duration_seconds: learningSyncNumberV2_(row.duration_seconds),
      total_questions: learningSyncNumberV2_(row.total_questions),
      answered_count: learningSyncNumberV2_(row.answered_count),
      correct_count: learningSyncNumberV2_(row.correct_count),
      wrong_count: learningSyncNumberV2_(row.wrong_count),
      unanswered_count: learningSyncNumberV2_(row.unanswered_count),
      earned_points: learningSyncNumberV2_(row.earned_points),
      total_possible_points: learningSyncNumberV2_(row.total_possible_points),
      section_score_100: learningSyncNumberV2_(row.section_score_100),
      source_rounds: row.source_rounds || '',
      correct_question_numbers: row.correct_question_numbers || '',
      wrong_question_numbers: row.wrong_question_numbers || '',
      unanswered_question_numbers: row.unanswered_question_numbers || '',
      client_version: row.client_version || '',
      sourceRow: row.__sourceRow,
      checksum: learningSyncChecksumV2_(checksumValues),
      syncedAt: now
    });
  });

  return {
    targetSs: targetSs,
    eligibleStudentCount: Object.keys(eligible).length,
    scannedRows: source.rows.length,
    records: records
  };
}

function topik1ReadingSyncPreviewStatsV2_(targetSs, records) {
  const sheet = targetSs.getSheetByName(TOPIK1_READING_SYNC_V2_.TARGET_SHEET);
  if (!sheet || sheet.getLastRow() < 2) {
    return { wouldInsert: records.length, wouldUpdate: 0, wouldSkip: 0, targetSheetExists: !!sheet };
  }

  const headers = TOPIK1_READING_SYNC_HEADERS_V2_;
  const keyIndex = headers.indexOf('attempt_id');
  const checksumIndex = headers.indexOf('checksum');
  const existing = {};
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getDisplayValues();

  values.forEach(function(row) {
    const key = learningSyncTextV2_(row[keyIndex]);
    if (!key) return;
    existing[key] = learningSyncTextV2_(row[checksumIndex]);
  });

  let insert = 0;
  let update = 0;
  let skip = 0;
  records.forEach(function(record) {
    const key = learningSyncTextV2_(record.attempt_id);
    if (!Object.prototype.hasOwnProperty.call(existing, key)) {
      insert++;
    } else if (existing[key] === learningSyncTextV2_(record.checksum)) {
      skip++;
    } else {
      update++;
    }
  });

  return { wouldInsert: insert, wouldUpdate: update, wouldSkip: skip, targetSheetExists: true };
}

function topik1ReadingSyncUpsertV2_(targetSs, records) {
  const sheet = topik1ReadingSyncEnsureTargetSheetV2_(targetSs);
  const headers = TOPIK1_READING_SYNC_HEADERS_V2_;
  const keyIndex = headers.indexOf('attempt_id');
  const checksumIndex = headers.indexOf('checksum');
  const existing = {};
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
    values.forEach(function(row, index) {
      const key = learningSyncTextV2_(row[keyIndex]);
      if (!key) return;
      existing[key] = {
        rowNumber: index + 2,
        checksum: learningSyncTextV2_(row[checksumIndex])
      };
    });
  }

  const appendRows = [];
  let updated = 0;
  let skipped = 0;

  records.forEach(function(record) {
    const key = learningSyncTextV2_(record.attempt_id);
    const found = existing[key];
    if (!found) {
      appendRows.push(topik1ReadingSyncRowForHeadersV2_(record));
      return;
    }
    if (found.checksum === learningSyncTextV2_(record.checksum)) {
      skipped++;
      return;
    }
    sheet.getRange(found.rowNumber, 1, 1, headers.length)
      .setValues([topik1ReadingSyncRowForHeadersV2_(record)]);
    updated++;
  });

  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
  }

  return { inserted: appendRows.length, updated: updated, skipped: skipped };
}

function previewTopik1ReadingSyncV2() {
  const data = topik1ReadingSyncCollectV2_();
  const stats = topik1ReadingSyncPreviewStatsV2_(data.targetSs, data.records);
  const result = {
    mode: 'PREVIEW_TOPIK1_READING',
    eligibleStudentCount: data.eligibleStudentCount,
    scannedRows: data.scannedRows,
    matchedRows: data.records.length,
    changes: stats,
    note: '이 함수는 데이터를 수정하지 않습니다.'
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function topik1ReadingSyncRunV2_(runType) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    const error = new Error('다른 동기화 작업이 실행 중입니다. 잠시 후 다시 실행해 주세요.');
    error.code = 'TOPIK1_READING_SYNC_ALREADY_RUNNING';
    throw error;
  }

  const startedAt = new Date();
  let targetSs = null;
  try {
    const data = topik1ReadingSyncCollectV2_();
    targetSs = data.targetSs;
    const stats = topik1ReadingSyncUpsertV2_(targetSs, data.records);
    const finishedAt = new Date();
    const sourceStats = {
      scannedRows: data.scannedRows,
      matchedRows: data.records.length,
      insertedRows: stats.inserted,
      updatedRows: stats.updated,
      skippedRows: stats.skipped
    };

    learningSyncWriteStateV2_(targetSs, TOPIK1_READING_SYNC_V2_.SOURCE_TYPE, sourceStats, 'SUCCESS', '');
    learningSyncAppendLogV2_(targetSs, runType, TOPIK1_READING_SYNC_V2_.SOURCE_TYPE,
      startedAt, finishedAt, 'SUCCESS', sourceStats, '');

    const result = {
      mode: runType === 'TOPIK1_READING_AUTO' ? 'APPLY_TOPIK1_READING_AUTO' : 'APPLY_TOPIK1_READING_MANUAL',
      runType: runType,
      eligibleStudentCount: data.eligibleStudentCount,
      scannedRows: data.scannedRows,
      matchedRows: data.records.length,
      results: stats
    };
    console.log(JSON.stringify(result, null, 2));
    return result;
  } catch (err) {
    const finishedAt = new Date();
    const message = err && err.message ? err.message : String(err);
    if (targetSs) {
      try {
        learningSyncAppendLogV2_(targetSs, runType, TOPIK1_READING_SYNC_V2_.SOURCE_TYPE,
          startedAt, finishedAt, 'FAILED', {}, message);
      } catch (ignore) {}
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function applyTopik1ReadingSyncV2() {
  return topik1ReadingSyncRunV2_('TOPIK1_READING_MANUAL');
}

function runAutomaticTopik1ReadingSyncV2() {
  return topik1ReadingSyncRunV2_('TOPIK1_READING_AUTO');
}

function topik1ReadingSyncDeleteAutoTriggersV2_() {
  const handler = TOPIK1_READING_SYNC_V2_.AUTO_HANDLER;
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  triggers.forEach(function(trigger) {
    if (trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
      deleted++;
    }
  });
  return deleted;
}

function installTopik1ReadingAutoSync5MinV2() {
  const deletedExisting = topik1ReadingSyncDeleteAutoTriggersV2_();
  const trigger = ScriptApp.newTrigger(TOPIK1_READING_SYNC_V2_.AUTO_HANDLER)
    .timeBased()
    .everyMinutes(5)
    .create();
  const result = {
    installed: true,
    handler: TOPIK1_READING_SYNC_V2_.AUTO_HANDLER,
    everyMinutes: 5,
    deletedExisting: deletedExisting,
    triggerId: trigger.getUniqueId()
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function removeTopik1ReadingAutoSyncV2() {
  const deleted = topik1ReadingSyncDeleteAutoTriggersV2_();
  const result = { removed: true, handler: TOPIK1_READING_SYNC_V2_.AUTO_HANDLER, deleted: deleted };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function getTopik1ReadingAutoSyncStatusV2() {
  const handler = TOPIK1_READING_SYNC_V2_.AUTO_HANDLER;
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
