/**
 * Student Management V2 - Lesson history migration (Step 7C)
 *
 * Purpose:
 * - Read only the approved legacy lesson rows L-159..L-189 from the old Google-based system.
 * - Import only C-013 / C-014 into the current V2 Lessons sheet.
 * - Preserve legacy lessonId and createdAt values.
 * - Never modify the legacy source spreadsheet.
 * - Create a target Lessons backup before the first actual write.
 * - Avoid whole-sheet rewrite of the target Lessons sheet.
 *
 * Important ordering rule:
 * - The normal V2 list API scans from the sheet tail for recent lessons.
 * - Therefore historical rows are inserted before row 2, not appended after current V2 lessons.
 *   This keeps newer V2 lessons at the bottom and preserves recent-list behavior.
 */

const LESSON_MIGRATION_STEP7C_ = Object.freeze({
  VERSION: 'lesson-migration-step7c-20261004',
  SOURCE_SPREADSHEET_ID: '1Y5qoA0mQp-7EAQoXM7GsOF6y3MkdT5XhVEmjGjQc_TY',
  SOURCE_SHEET: 'Lessons',
  TARGET_SHEET: 'Lessons',
  ALLOWED_CLASS_IDS: Object.freeze(['C-013', 'C-014']),
  EXPECTED_COUNTS: Object.freeze({ 'C-013': 19, 'C-014': 12 }),
  EXPECTED_TOTAL: 31,
  CONFIRM_TEXT: 'IMPORT_C013_C014_31',
  LEGACY_IDS: Object.freeze([
    'L-159','L-160','L-161','L-162','L-163','L-164','L-165','L-166','L-167','L-168',
    'L-169','L-170','L-171','L-172','L-173','L-174','L-175','L-176','L-177','L-178',
    'L-179','L-180','L-181','L-182','L-183','L-184','L-185','L-186','L-187','L-188','L-189'
  ]),
  REQUIRED_HEADERS: Object.freeze([
    'lessonId','date','classId','topic','content','homework','nextPlan','progressCurrentCount','createdAt'
  ])
});

function lessonMigrationAssertAdminV2_(auth) {
  requireSuperAdminV2_(auth);
}

function lessonMigrationHeaderInfoV2_(sheet) {
  const headers = lessonHeadersV2_(sheet);
  const map = {};
  headers.forEach(function(h, i) {
    if (h) map[h] = i;
  });
  LESSON_MIGRATION_STEP7C_.REQUIRED_HEADERS.forEach(function(name) {
    if (map[name] == null) {
      const error = new Error(sheet.getName() + ' 시트에 필수 열이 없습니다: ' + name);
      error.code = 'LESSON_MIGRATION_SCHEMA_ERROR';
      throw error;
    }
  });
  return { headers: headers, map: map };
}

function lessonMigrationReadRowsV2_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  const headers = lessonHeadersV2_(sheet);
  const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  return values.filter(function(row) {
    return row.some(function(v) { return String(v || '').trim() !== ''; });
  }).map(function(row) {
    return lessonRowToObjectV2_(headers, row);
  });
}

function lessonMigrationExactKeyV2_(row) {
  return [
    lessonNormalizeDateV2_(row.date),
    String(row.classId || '').trim(),
    String(row.topic || '').trim(),
    String(row.content || '').trim(),
    String(row.homework || '').trim(),
    String(row.nextPlan || '').trim(),
    String(row.progressCurrentCount == null ? '' : row.progressCurrentCount).trim()
  ].join('||');
}

function lessonMigrationSourceRowsV2_() {
  const sourceSs = SpreadsheetApp.openById(LESSON_MIGRATION_STEP7C_.SOURCE_SPREADSHEET_ID);
  const sourceSheet = sourceSs.getSheetByName(LESSON_MIGRATION_STEP7C_.SOURCE_SHEET);
  if (!sourceSheet) {
    const error = new Error('기존 Google 기반 Lessons 시트를 찾을 수 없습니다.');
    error.code = 'LESSON_MIGRATION_SOURCE_NOT_FOUND';
    throw error;
  }
  lessonMigrationHeaderInfoV2_(sourceSheet);

  const allowedIds = {};
  LESSON_MIGRATION_STEP7C_.LEGACY_IDS.forEach(function(id) { allowedIds[id] = true; });
  const allowedClasses = {};
  LESSON_MIGRATION_STEP7C_.ALLOWED_CLASS_IDS.forEach(function(id) { allowedClasses[id] = true; });

  const rows = lessonMigrationReadRowsV2_(sourceSheet).filter(function(row) {
    const lessonId = String(row.lessonId || '').trim();
    const classId = String(row.classId || '').trim();
    return !!allowedIds[lessonId] && !!allowedClasses[classId];
  });

  const idCounts = {};
  const classCounts = {};
  rows.forEach(function(row) {
    const id = String(row.lessonId || '').trim();
    const classId = String(row.classId || '').trim();
    idCounts[id] = (idCounts[id] || 0) + 1;
    classCounts[classId] = (classCounts[classId] || 0) + 1;
  });

  const missingIds = LESSON_MIGRATION_STEP7C_.LEGACY_IDS.filter(function(id) { return !idCounts[id]; });
  const duplicateIds = Object.keys(idCounts).filter(function(id) { return idCounts[id] !== 1; });
  const countMismatch = rows.length !== LESSON_MIGRATION_STEP7C_.EXPECTED_TOTAL ||
    LESSON_MIGRATION_STEP7C_.ALLOWED_CLASS_IDS.some(function(id) {
      return Number(classCounts[id] || 0) !== Number(LESSON_MIGRATION_STEP7C_.EXPECTED_COUNTS[id] || 0);
    });

  if (missingIds.length || duplicateIds.length || countMismatch) {
    const error = new Error('기존 수업 원본이 Step 7B 승인 상태와 다릅니다. 실제 이관을 중단합니다.');
    error.code = 'LESSON_MIGRATION_SOURCE_CHANGED';
    error.details = {
      rowCount: rows.length,
      classCounts: classCounts,
      missingIds: missingIds,
      duplicateIds: duplicateIds
    };
    throw error;
  }

  rows.sort(function(a, b) {
    const da = lessonNormalizeDateV2_(a.date);
    const db = lessonNormalizeDateV2_(b.date);
    if (da !== db) return da < db ? -1 : 1;
    const ca = String(a.classId || '');
    const cb = String(b.classId || '');
    if (ca !== cb) return ca < cb ? -1 : 1;
    const ta = String(a.createdAt || '');
    const tb = String(b.createdAt || '');
    if (ta !== tb) return ta < tb ? -1 : 1;
    return String(a.lessonId || '').localeCompare(String(b.lessonId || ''));
  });

  return {
    spreadsheetTitle: sourceSs.getName(),
    rows: rows,
    classCounts: classCounts
  };
}

function lessonMigrationTargetContextV2_(auth) {
  const targetSs = getTeacherDataSpreadsheet_(auth.teacher);
  const targetSheet = getRequiredDataSheetV2_(targetSs, LESSON_MIGRATION_STEP7C_.TARGET_SHEET);
  lessonMigrationHeaderInfoV2_(targetSheet);

  const classes = readSheetObjects_(targetSs, 'Classes');
  const classIds = {};
  classes.forEach(function(row) {
    const id = String(row.classId || '').trim();
    if (id) classIds[id] = true;
  });
  const missingClasses = LESSON_MIGRATION_STEP7C_.ALLOWED_CLASS_IDS.filter(function(id) { return !classIds[id]; });
  if (missingClasses.length) {
    const error = new Error('V2에 이관 대상 반이 없습니다: ' + missingClasses.join(', '));
    error.code = 'LESSON_MIGRATION_TARGET_CLASS_MISSING';
    throw error;
  }

  return { ss: targetSs, sheet: targetSheet };
}

function lessonMigrationAnalyzeV2_(auth) {
  lessonMigrationAssertAdminV2_(auth);
  const source = lessonMigrationSourceRowsV2_();
  const target = lessonMigrationTargetContextV2_(auth);
  const existing = lessonMigrationReadRowsV2_(target.sheet);

  const existingIds = {};
  const existingKeys = {};
  existing.forEach(function(row) {
    const id = String(row.lessonId || '').trim();
    if (id) existingIds[id] = row;
    existingKeys[lessonMigrationExactKeyV2_(row)] = row;
  });

  const ready = [];
  const skippedIdCollision = [];
  const skippedExactDuplicate = [];
  source.rows.forEach(function(row) {
    const id = String(row.lessonId || '').trim();
    if (existingIds[id]) {
      skippedIdCollision.push({ lessonId: id, existingLessonId: id });
      return;
    }
    const key = lessonMigrationExactKeyV2_(row);
    if (existingKeys[key]) {
      skippedExactDuplicate.push({
        lessonId: id,
        duplicateV2LessonId: String(existingKeys[key].lessonId || '')
      });
      return;
    }
    ready.push(row);
  });

  const readyClassCounts = {};
  ready.forEach(function(row) {
    const id = String(row.classId || '').trim();
    readyClassCounts[id] = (readyClassCounts[id] || 0) + 1;
  });

  return {
    source: source,
    target: target,
    existing: existing,
    ready: ready,
    readyClassCounts: readyClassCounts,
    skippedIdCollision: skippedIdCollision,
    skippedExactDuplicate: skippedExactDuplicate
  };
}

function previewLegacyLessonsMigrationStep7CV2_(auth) {
  const analysis = lessonMigrationAnalyzeV2_(auth);
  return {
    version: LESSON_MIGRATION_STEP7C_.VERSION,
    readOnly: true,
    source: {
      spreadsheetTitle: analysis.source.spreadsheetTitle,
      approvedClasses: LESSON_MIGRATION_STEP7C_.ALLOWED_CLASS_IDS.slice(),
      approvedLegacyCount: analysis.source.rows.length,
      approvedClassCounts: analysis.source.classCounts
    },
    target: {
      spreadsheetTitle: analysis.target.ss.getName(),
      existingLessonCount: analysis.existing.length
    },
    readyCount: analysis.ready.length,
    readyClassCounts: analysis.readyClassCounts,
    skippedIdCollisionCount: analysis.skippedIdCollision.length,
    skippedExactDuplicateCount: analysis.skippedExactDuplicate.length,
    skippedIdCollisions: analysis.skippedIdCollision,
    skippedExactDuplicates: analysis.skippedExactDuplicate,
    approvedLessonIds: LESSON_MIGRATION_STEP7C_.LEGACY_IDS.slice(),
    insertionMode: 'insert-before-row-2',
    reasonForInsertionMode: '최근 목록 API가 시트 끝에서 역방향 조회하므로 과거 기록은 현재 V2 기록보다 위에 삽입합니다.',
    willCreateBackup: analysis.ready.length > 0,
    sourceWillBeModified: false,
    targetWholeSheetRewrite: false,
    confirmText: LESSON_MIGRATION_STEP7C_.CONFIRM_TEXT
  };
}

function lessonMigrationUniqueBackupNameV2_(ss) {
  const tz = Session.getScriptTimeZone() || 'Asia/Ulaanbaatar';
  const base = 'Lessons_BACKUP_STEP7C_' + Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  let name = base;
  let n = 2;
  while (ss.getSheetByName(name)) {
    name = base + '_' + n;
    n++;
  }
  return name;
}

function lessonMigrationBackupV2_(ss, targetSheet) {
  const copy = targetSheet.copyTo(ss);
  const name = lessonMigrationUniqueBackupNameV2_(ss);
  copy.setName(name);
  copy.hideSheet();
  return {
    sheetName: name,
    hidden: true,
    dataRows: Math.max(0, targetSheet.getLastRow() - 1)
  };
}

function lessonMigrationBuildTargetRowsV2_(targetSheet, rows) {
  const headers = lessonHeadersV2_(targetSheet);
  return rows.map(function(row) {
    return headers.map(function(header) {
      if (!header) return '';
      if (!Object.prototype.hasOwnProperty.call(row, header)) return '';
      if (header === 'date') return lessonNormalizeDateV2_(row[header]);
      return row[header];
    });
  });
}

function executeLegacyLessonsMigrationStep7CV2_(auth, confirmText) {
  lessonMigrationAssertAdminV2_(auth);
  if (String(confirmText || '').trim() !== LESSON_MIGRATION_STEP7C_.CONFIRM_TEXT) {
    const error = new Error('이관 확인 문구가 일치하지 않습니다.');
    error.code = 'LESSON_MIGRATION_CONFIRM_REQUIRED';
    throw error;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const analysis = lessonMigrationAnalyzeV2_(auth);
    const ready = analysis.ready;
    const beforeCount = analysis.existing.length;

    if (!ready.length) {
      return {
        version: LESSON_MIGRATION_STEP7C_.VERSION,
        mode: 'no-op',
        inserted: 0,
        beforeCount: beforeCount,
        afterCount: beforeCount,
        skippedIdCollisionCount: analysis.skippedIdCollision.length,
        skippedExactDuplicateCount: analysis.skippedExactDuplicate.length,
        backup: null,
        verified: true,
        sourceModified: false,
        targetWholeSheetRewrite: false
      };
    }

    const targetSheet = analysis.target.sheet;
    const backup = lessonMigrationBackupV2_(analysis.target.ss, targetSheet);
    const values = lessonMigrationBuildTargetRowsV2_(targetSheet, ready);

    // Historical rows must stay above the current V2 rows so tail-based recent lookup remains correct.
    targetSheet.insertRowsBefore(2, ready.length);
    targetSheet.getRange(2, 1, values.length, targetSheet.getLastColumn()).setValues(values);
    SpreadsheetApp.flush();

    const insertedIds = targetSheet.getRange(2, 1, ready.length, targetSheet.getLastColumn()).getDisplayValues().map(function(row) {
      const headers = lessonHeadersV2_(targetSheet);
      const idIndex = headers.indexOf('lessonId');
      return idIndex >= 0 ? String(row[idIndex] || '').trim() : '';
    });
    const expectedIds = ready.map(function(row) { return String(row.lessonId || '').trim(); });
    const insertedIdSet = {};
    insertedIds.forEach(function(id) { if (id) insertedIdSet[id] = true; });
    const missingAfterWrite = expectedIds.filter(function(id) { return !insertedIdSet[id]; });

    const afterRows = lessonMigrationReadRowsV2_(targetSheet);
    const afterCount = afterRows.length;
    const expectedAfter = beforeCount + ready.length;
    const classCounts = {};
    afterRows.forEach(function(row) {
      const id = String(row.lessonId || '').trim();
      if (LESSON_MIGRATION_STEP7C_.LEGACY_IDS.indexOf(id) < 0) return;
      const classId = String(row.classId || '').trim();
      classCounts[classId] = (classCounts[classId] || 0) + 1;
    });

    const classVerified = LESSON_MIGRATION_STEP7C_.ALLOWED_CLASS_IDS.every(function(id) {
      return Number(classCounts[id] || 0) === Number(LESSON_MIGRATION_STEP7C_.EXPECTED_COUNTS[id] || 0);
    });
    const verified = missingAfterWrite.length === 0 && afterCount === expectedAfter && classVerified;

    appendAuditLog_(
      auth.teacher.teacherId,
      'LESSON_MIGRATION_STEP7C',
      'Lessons',
      'C-013,C-014',
      verified ? 'SUCCESS' : 'VERIFY_FAILED',
      'inserted=' + ready.length + ',before=' + beforeCount + ',after=' + afterCount + ',backup=' + backup.sheetName
    );

    if (!verified) {
      const error = new Error('이관 후 검증값이 일치하지 않습니다. 백업 시트 ' + backup.sheetName + ' 을 유지한 채 작업을 중단합니다.');
      error.code = 'LESSON_MIGRATION_VERIFY_FAILED';
      error.details = {
        missingAfterWrite: missingAfterWrite,
        beforeCount: beforeCount,
        afterCount: afterCount,
        expectedAfter: expectedAfter,
        classCounts: classCounts,
        backup: backup
      };
      throw error;
    }

    return {
      version: LESSON_MIGRATION_STEP7C_.VERSION,
      mode: 'imported',
      inserted: ready.length,
      insertedClassCounts: classCounts,
      beforeCount: beforeCount,
      afterCount: afterCount,
      skippedIdCollisionCount: analysis.skippedIdCollision.length,
      skippedExactDuplicateCount: analysis.skippedExactDuplicate.length,
      backup: backup,
      insertionMode: 'insert-before-row-2',
      verified: true,
      sourceModified: false,
      targetWholeSheetRewrite: false,
      nextGeneratedLessonIdNote: '기존 ID를 보존하므로 이후 신규 lessonId는 현재 최대 숫자 다음 번호로 생성됩니다.'
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Step 7E - remove only the temporary lesson test rows created during Steps 3-5.
 *
 * Safety rules:
 * - SUPER_ADMIN only.
 * - Exact IDs only: L-00001, L-00002.
 * - Topic text must still match the known test markers before deletion.
 * - Back up Lessons and LessonAssignments before the first delete.
 * - Remove linked LessonAssignments only for those exact lessonIds.
 * - Do not rewrite either whole sheet.
 */
const LESSON_TEST_CLEANUP_STEP7E_ = Object.freeze({
  VERSION: 'lesson-test-cleanup-step7e-20261004',
  CONFIRM_TEXT: 'CLEANUP_LESSON_TEST_ROWS_2',
  TEST_ROWS: Object.freeze([
    Object.freeze({ lessonId: 'L-00001', topicIncludes: '[STEP3 TEST]' }),
    Object.freeze({ lessonId: 'L-00002', topicIncludes: '[실제 테스트]' })
  ])
});

function lessonCleanupHeaderMapV2_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return { headers: [], map: {} };
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) {
    return String(v || '').trim();
  });
  const map = {};
  headers.forEach(function(h, i) { if (h) map[h] = i; });
  return { headers: headers, map: map };
}

function lessonCleanupFindRowsByValueV2_(sheet, headerName, exactValue) {
  const info = lessonCleanupHeaderMapV2_(sheet);
  const index = info.map[headerName];
  if (index == null) {
    const error = new Error(sheet.getName() + ' 시트에 필수 열이 없습니다: ' + headerName);
    error.code = 'LESSON_CLEANUP_SCHEMA_ERROR';
    throw error;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const finder = sheet.getRange(2, index + 1, lastRow - 1, 1)
    .createTextFinder(String(exactValue || ''))
    .matchEntireCell(true);
  return finder.findAll().map(function(range) { return range.getRow(); });
}

function lessonCleanupRowObjectV2_(sheet, rowNumber) {
  const info = lessonCleanupHeaderMapV2_(sheet);
  const row = sheet.getRange(rowNumber, 1, 1, info.headers.length).getDisplayValues()[0];
  return lessonRowToObjectV2_(info.headers, row);
}

function lessonCleanupAnalyzeV2_(auth) {
  requireSuperAdminV2_(auth);
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const lessonsSheet = getRequiredDataSheetV2_(ss, 'Lessons');
  const assignmentsSheet = getRequiredDataSheetV2_(ss, 'LessonAssignments');

  const foundLessons = [];
  const missingLessonIds = [];
  const unsafeRows = [];

  LESSON_TEST_CLEANUP_STEP7E_.TEST_ROWS.forEach(function(expected) {
    const rows = lessonCleanupFindRowsByValueV2_(lessonsSheet, 'lessonId', expected.lessonId);
    if (rows.length !== 1) {
      if (!rows.length) missingLessonIds.push(expected.lessonId);
      else unsafeRows.push({ lessonId: expected.lessonId, reason: '동일 lessonId 행이 ' + rows.length + '개입니다.' });
      return;
    }
    const object = lessonCleanupRowObjectV2_(lessonsSheet, rows[0]);
    const topic = String(object.topic || '');
    if (topic.indexOf(expected.topicIncludes) < 0) {
      unsafeRows.push({ lessonId: expected.lessonId, reason: '테스트 주제 표식이 일치하지 않습니다.', topic: topic });
      return;
    }
    foundLessons.push({ lessonId: expected.lessonId, rowNumber: rows[0], lesson: object });
  });

  const linkedAssignments = [];
  LESSON_TEST_CLEANUP_STEP7E_.TEST_ROWS.forEach(function(expected) {
    const rows = lessonCleanupFindRowsByValueV2_(assignmentsSheet, 'lessonId', expected.lessonId);
    rows.forEach(function(rowNumber) {
      const info = lessonCleanupHeaderMapV2_(assignmentsSheet);
      const row = assignmentsSheet.getRange(rowNumber, 1, 1, info.headers.length).getDisplayValues()[0];
      const object = {};
      info.headers.forEach(function(h, i) { if (h) object[h] = row[i] == null ? '' : row[i]; });
      linkedAssignments.push({ rowNumber: rowNumber, assignment: object });
    });
  });

  const allLessons = lessonMigrationReadRowsV2_(lessonsSheet);
  const classCounts = {};
  allLessons.forEach(function(row) {
    const classId = String(row.classId || '').trim();
    if (classId) classCounts[classId] = (classCounts[classId] || 0) + 1;
  });

  return {
    ss: ss,
    lessonsSheet: lessonsSheet,
    assignmentsSheet: assignmentsSheet,
    foundLessons: foundLessons,
    missingLessonIds: missingLessonIds,
    unsafeRows: unsafeRows,
    linkedAssignments: linkedAssignments,
    beforeLessonCount: allLessons.length,
    beforeClassCounts: classCounts,
    safeToExecute: foundLessons.length === LESSON_TEST_CLEANUP_STEP7E_.TEST_ROWS.length && !missingLessonIds.length && !unsafeRows.length
  };
}

function previewLessonTestCleanupStep7EV2_(auth) {
  const a = lessonCleanupAnalyzeV2_(auth);
  return {
    version: LESSON_TEST_CLEANUP_STEP7E_.VERSION,
    readOnly: true,
    safeToExecute: a.safeToExecute,
    currentLessonCount: a.beforeLessonCount,
    currentClassCounts: a.beforeClassCounts,
    testLessonCount: a.foundLessons.length,
    testLessons: a.foundLessons.map(function(x) { return x.lesson; }),
    linkedAssignmentCount: a.linkedAssignments.length,
    linkedAssignments: a.linkedAssignments.map(function(x) { return x.assignment; }),
    missingLessonIds: a.missingLessonIds,
    unsafeRows: a.unsafeRows,
    expectedAfterLessonCount: a.beforeLessonCount - a.foundLessons.length,
    expectedAfterClassCounts: {
      'C-013': Number(a.beforeClassCounts['C-013'] || 0) - a.foundLessons.filter(function(x) { return String(x.lesson.classId || '') === 'C-013'; }).length,
      'C-014': Number(a.beforeClassCounts['C-014'] || 0) - a.foundLessons.filter(function(x) { return String(x.lesson.classId || '') === 'C-014'; }).length
    },
    confirmText: LESSON_TEST_CLEANUP_STEP7E_.CONFIRM_TEXT,
    willBackupLessons: a.safeToExecute,
    willBackupAssignments: a.safeToExecute,
    wholeSheetRewrite: false
  };
}

function lessonCleanupUniqueBackupNameV2_(ss, base) {
  const tz = Session.getScriptTimeZone() || 'Asia/Ulaanbaatar';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyyMMdd_HHmmss');
  let name = base + '_' + stamp;
  let n = 2;
  while (ss.getSheetByName(name)) {
    name = base + '_' + stamp + '_' + n;
    n++;
  }
  return name;
}

function lessonCleanupBackupSheetV2_(ss, sheet, base) {
  const copy = sheet.copyTo(ss);
  const name = lessonCleanupUniqueBackupNameV2_(ss, base);
  copy.setName(name);
  copy.hideSheet();
  return { sheetName: name, hidden: true, dataRows: Math.max(0, sheet.getLastRow() - 1) };
}

function executeLessonTestCleanupStep7EV2_(auth, confirmText) {
  requireSuperAdminV2_(auth);
  if (String(confirmText || '').trim() !== LESSON_TEST_CLEANUP_STEP7E_.CONFIRM_TEXT) {
    const error = new Error('테스트 기록 정리 확인 문구가 일치하지 않습니다.');
    error.code = 'LESSON_CLEANUP_CONFIRM_REQUIRED';
    throw error;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const a = lessonCleanupAnalyzeV2_(auth);
    if (!a.safeToExecute) {
      const error = new Error('테스트 기록이 예상 상태와 달라 자동 정리를 중단합니다.');
      error.code = 'LESSON_CLEANUP_UNSAFE_STATE';
      error.details = { missingLessonIds: a.missingLessonIds, unsafeRows: a.unsafeRows };
      throw error;
    }

    const lessonsBackup = lessonCleanupBackupSheetV2_(a.ss, a.lessonsSheet, 'Lessons_BACKUP_STEP7E');
    const assignmentsBackup = lessonCleanupBackupSheetV2_(a.ss, a.assignmentsSheet, 'LessonAssignments_BACKUP_STEP7E');

    const assignmentRows = a.linkedAssignments.map(function(x) { return x.rowNumber; }).sort(function(x, y) { return y - x; });
    assignmentRows.forEach(function(rowNumber) { a.assignmentsSheet.deleteRow(rowNumber); });

    const lessonRows = a.foundLessons.map(function(x) { return x.rowNumber; }).sort(function(x, y) { return y - x; });
    lessonRows.forEach(function(rowNumber) { a.lessonsSheet.deleteRow(rowNumber); });
    SpreadsheetApp.flush();

    const afterLessons = lessonMigrationReadRowsV2_(a.lessonsSheet);
    const remainingIds = {};
    const classCounts = {};
    afterLessons.forEach(function(row) {
      const lessonId = String(row.lessonId || '').trim();
      if (lessonId) remainingIds[lessonId] = true;
      const classId = String(row.classId || '').trim();
      if (classId) classCounts[classId] = (classCounts[classId] || 0) + 1;
    });
    const linkedRemaining = LESSON_TEST_CLEANUP_STEP7E_.TEST_ROWS.reduce(function(total, expected) {
      return total + lessonCleanupFindRowsByValueV2_(a.assignmentsSheet, 'lessonId', expected.lessonId).length;
    }, 0);

    const removedIds = LESSON_TEST_CLEANUP_STEP7E_.TEST_ROWS.map(function(x) { return x.lessonId; });
    const idsGone = removedIds.every(function(id) { return !remainingIds[id]; });
    const verified = afterLessons.length === 31 && Number(classCounts['C-013'] || 0) === 19 && Number(classCounts['C-014'] || 0) === 12 && idsGone && linkedRemaining === 0;

    appendAuditLog_(
      auth.teacher.teacherId,
      'LESSON_TEST_CLEANUP_STEP7E',
      'Lessons,LessonAssignments',
      removedIds.join(','),
      verified ? 'SUCCESS' : 'VERIFY_FAILED',
      'before=' + a.beforeLessonCount + ',after=' + afterLessons.length + ',assignmentsRemoved=' + assignmentRows.length + ',lessonsBackup=' + lessonsBackup.sheetName + ',assignmentsBackup=' + assignmentsBackup.sheetName
    );

    if (!verified) {
      const error = new Error('테스트 기록 정리 후 검증값이 일치하지 않습니다. 백업 시트를 유지한 채 작업을 중단합니다.');
      error.code = 'LESSON_CLEANUP_VERIFY_FAILED';
      error.details = {
        afterLessonCount: afterLessons.length,
        classCounts: classCounts,
        idsGone: idsGone,
        linkedRemaining: linkedRemaining,
        lessonsBackup: lessonsBackup,
        assignmentsBackup: assignmentsBackup
      };
      throw error;
    }

    return {
      version: LESSON_TEST_CLEANUP_STEP7E_.VERSION,
      mode: 'cleaned',
      removedLessonIds: removedIds,
      removedLessonCount: lessonRows.length,
      removedAssignmentCount: assignmentRows.length,
      beforeLessonCount: a.beforeLessonCount,
      afterLessonCount: afterLessons.length,
      afterClassCounts: classCounts,
      lessonsBackup: lessonsBackup,
      assignmentsBackup: assignmentsBackup,
      verified: true,
      wholeSheetRewrite: false,
      note: '과거 이관 31건은 유지하고 개발 중 생성한 테스트 수업 2건과 연결 과제만 제거했습니다.'
    };
  } finally {
    lock.releaseLock();
  }
}
