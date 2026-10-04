/**
 * Student Management V2 - Lesson archive/restore service (Step 8)
 *
 * Policy:
 * - No permanent delete in Step 8.
 * - Archiving moves one lesson and its linked LessonAssignments into archive sheets.
 * - Restoring moves them back.
 * - Active Lessons / LessonAssignments are changed only by the affected rows.
 * - Archive sheets are created lazily and never used during login/bootstrap.
 */

const LESSON_ARCHIVE_V2_ = Object.freeze({
  LESSON_SHEET: 'LessonsArchive',
  ASSIGNMENT_SHEET: 'LessonAssignmentsArchive',
  DEFAULT_LIMIT: 30,
  MAX_LIMIT: 100,
  CHUNK_SIZE: 200,
  MAX_SCAN_ROWS: 4000
});

function lessonArchiveMaxPrefixedNumberV2_(sheet, headerName, prefix) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const map = getSheetHeaderMapV2_(sheet);
  if (map[headerName] == null) return 0;
  const values = sheet.getRange(2, map[headerName] + 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  let max = 0;
  values.forEach(function(row) {
    const id = String(row[0] || '').trim();
    if (id.indexOf(prefix) !== 0) return;
    const n = Number(id.substring(prefix.length));
    if (Number.isFinite(n) && n > max) max = n;
  });
  return max;
}

function makeNextLessonIdIncludingArchiveV2_(ss, activeSheet) {
  const maxActive = lessonArchiveMaxPrefixedNumberV2_(activeSheet, 'lessonId', 'L-');
  const maxArchive = lessonArchiveMaxPrefixedNumberV2_(ss.getSheetByName(LESSON_ARCHIVE_V2_.LESSON_SHEET), 'lessonId', 'L-');
  return 'L-' + String(Math.max(maxActive, maxArchive) + 1).padStart(5, '0');
}

function makeNextLessonAssignmentIdIncludingArchiveV2_(ss, activeSheet) {
  const maxActive = lessonArchiveMaxPrefixedNumberV2_(activeSheet, 'assignmentId', 'LA-');
  const maxArchive = lessonArchiveMaxPrefixedNumberV2_(ss.getSheetByName(LESSON_ARCHIVE_V2_.ASSIGNMENT_SHEET), 'assignmentId', 'LA-');
  return 'LA-' + String(Math.max(maxActive, maxArchive) + 1).padStart(5, '0');
}

function lessonArchiveEnsureSheetV2_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function lessonArchiveLessonSheetV2_(ss) {
  return lessonArchiveEnsureSheetV2_(ss, LESSON_ARCHIVE_V2_.LESSON_SHEET, [
    'lessonId', 'date', 'classId', 'topic', 'content', 'homework', 'nextPlan',
    'progressCurrentCount', 'createdAt', 'updatedAt', 'archivedAt', 'archivedBy'
  ]);
}

function lessonArchiveAssignmentSheetV2_(ss) {
  return lessonArchiveEnsureSheetV2_(ss, LESSON_ARCHIVE_V2_.ASSIGNMENT_SHEET, [
    'assignmentId', 'lessonId', 'classId', 'studentId', 'assignmentText',
    'createdAt', 'legacyCreatedAt2', 'assignment', 'studentName',
    'archivedAt', 'archivedBy'
  ]);
}

function lessonArchiveFindRowByLessonIdV2_(sheet, lessonId) {
  const id = String(lessonId || '').trim();
  if (!id || sheet.getLastRow() < 2) return -1;
  const headers = getSheetHeaderMapV2_(sheet);
  if (headers.lessonId == null) return -1;
  const range = sheet.getRange(2, headers.lessonId + 1, sheet.getLastRow() - 1, 1);
  const found = range.createTextFinder(id).matchEntireCell(true).findNext();
  return found ? found.getRow() : -1;
}

function lessonArchiveReadRowV2_(sheet, rowNumber) {
  if (rowNumber < 2) return null;
  const headers = lessonHeadersV2_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  return lessonRowToObjectV2_(headers, values);
}

function lessonArchiveReadRecentV2_(ss, options) {
  options = options || {};
  const sheet = ss.getSheetByName(LESSON_ARCHIVE_V2_.LESSON_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return { rows: [], scannedRows: 0, hasMore: false };

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const headers = lessonHeadersV2_(sheet);
  const limit = Math.max(1, Math.min(LESSON_ARCHIVE_V2_.MAX_LIMIT, lessonNumberV2_(options.limit, LESSON_ARCHIVE_V2_.DEFAULT_LIMIT)));
  const classId = String(options.classId || '').trim();
  const startDate = lessonNormalizeDateV2_(options.startDate);
  const endDate = lessonNormalizeDateV2_(options.endDate);
  const rows = [];
  let scanned = 0;
  let endRow = lastRow;

  while (endRow >= 2 && scanned < LESSON_ARCHIVE_V2_.MAX_SCAN_ROWS && rows.length < limit) {
    const remaining = endRow - 1;
    const size = Math.min(LESSON_ARCHIVE_V2_.CHUNK_SIZE, remaining, LESSON_ARCHIVE_V2_.MAX_SCAN_ROWS - scanned);
    const startRow = endRow - size + 1;
    const values = sheet.getRange(startRow, 1, size, lastColumn).getDisplayValues();
    for (let i = values.length - 1; i >= 0; i--) {
      scanned++;
      const row = lessonRowToObjectV2_(headers, values[i]);
      const rowDate = lessonNormalizeDateV2_(row.date);
      if (!String(row.lessonId || '').trim()) continue;
      if (classId && String(row.classId || '').trim() !== classId) continue;
      if (startDate && (!rowDate || rowDate < startDate)) continue;
      if (endDate && (!rowDate || rowDate > endDate)) continue;
      rows.push(row);
      if (rows.length >= limit) break;
    }
    endRow = startRow - 1;
  }

  return { rows: rows, scannedRows: scanned, hasMore: endRow >= 2 };
}

function listArchivedLessonsV2_(auth, options) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const data = lessonArchiveReadRecentV2_(ss, options || {});
  const classMap = lessonClassMapV2_(ss);
  return {
    version: 'lesson-archive-step8-20261004',
    rows: data.rows.map(function(row) {
      const classId = String(row.classId || '').trim();
      const classInfo = classMap[classId] || {};
      return {
        lessonId: String(row.lessonId || ''),
        date: lessonNormalizeDateV2_(row.date),
        classId: classId,
        className: String(classInfo.className || classId),
        topic: String(row.topic || ''),
        content: String(row.content || ''),
        homework: String(row.homework || ''),
        nextPlan: String(row.nextPlan || ''),
        progressCurrentCount: row.progressCurrentCount === '' ? null : lessonNumberV2_(row.progressCurrentCount, null),
        progressUnit: String(classInfo.targetProgressUnit || ''),
        createdAt: String(row.createdAt || ''),
        archivedAt: String(row.archivedAt || ''),
        archivedBy: String(row.archivedBy || '')
      };
    }),
    count: data.rows.length,
    hasMore: data.hasMore,
    diagnostics: {
      scannedRows: data.scannedRows,
      maxScanRows: LESSON_ARCHIVE_V2_.MAX_SCAN_ROWS,
      chunkSize: LESSON_ARCHIVE_V2_.CHUNK_SIZE,
      fullSheetRead: false
    }
  };
}

function archiveLessonV2_(auth, lessonId) {
  const id = String(lessonId || '').trim();
  if (!id) {
    const error = new Error('보관할 수업 기록을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const lessonSheet = getRequiredDataSheetV2_(ss, 'Lessons');
  const assignmentSheet = getRequiredDataSheetV2_(ss, 'LessonAssignments');
  const archiveLessonSheet = lessonArchiveLessonSheetV2_(ss);
  const archiveAssignmentSheet = lessonArchiveAssignmentSheetV2_(ss);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (lessonArchiveFindRowByLessonIdV2_(archiveLessonSheet, id) >= 2) {
      const error = new Error('이미 보관된 수업 기록입니다.');
      error.code = 'LESSON_ALREADY_ARCHIVED';
      throw error;
    }

    const found = lessonGetRowV2_(lessonSheet, id);
    if (found.rowNumber < 2 || !found.row) {
      const error = new Error('보관할 수업 기록을 찾을 수 없습니다.');
      error.code = 'LESSON_NOT_FOUND';
      throw error;
    }

    const now = new Date();
    const teacherId = String(auth.teacher.teacherId || '');
    const lessonObj = Object.assign({}, found.row, {
      archivedAt: now,
      archivedBy: teacherId
    });
    appendObjectRowV2_(archiveLessonSheet, lessonObj);

    const linked = lessonAssignmentRowsForLessonV2_(assignmentSheet, id);
    linked.forEach(function(item) {
      const obj = Object.assign({}, item.row || {}, {
        archivedAt: now,
        archivedBy: teacherId
      });
      appendObjectRowV2_(archiveAssignmentSheet, obj);
    });

    linked.map(function(item) { return item.rowNumber; }).sort(function(a, b) { return b - a; }).forEach(function(rowNumber) {
      assignmentSheet.deleteRow(rowNumber);
    });
    lessonSheet.deleteRow(found.rowNumber);

    SpreadsheetApp.flush();
    appendAuditLog_(teacherId, 'LESSON_ARCHIVE', 'Lesson', id, 'SUCCESS', 'assignments=' + linked.length);

    return {
      version: 'lesson-archive-step8-20261004',
      lessonId: id,
      archivedAssignments: linked.length,
      diagnostics: {
        wholeSheetRewrite: false,
        rowOperation: 'copy-to-archive-and-delete-active-row',
        permanentDelete: false
      }
    };
  } finally {
    lock.releaseLock();
  }
}

function lessonArchiveInsertRowByDateV2_(sheet, obj) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    appendObjectRowV2_(sheet, obj);
    return sheet.getLastRow();
  }

  const headerMap = getSheetHeaderMapV2_(sheet);
  if (headerMap.date == null) {
    appendObjectRowV2_(sheet, obj);
    return sheet.getLastRow();
  }

  const restoredDate = lessonNormalizeDateV2_(obj.date);
  const count = Math.min(lastRow - 1, LESSON_READ_V2_.MAX_SCAN_ROWS);
  const values = sheet.getRange(2, headerMap.date + 1, count, 1).getDisplayValues();
  let insertAt = -1;
  for (let i = 0; i < values.length; i++) {
    const rowDate = lessonNormalizeDateV2_(values[i][0]);
    if (rowDate && restoredDate && rowDate > restoredDate) {
      insertAt = i + 2;
      break;
    }
  }

  if (insertAt < 2) {
    appendObjectRowV2_(sheet, obj);
    return sheet.getLastRow();
  }
  sheet.insertRowBefore(insertAt);
  writeObjectToRowV2_(sheet, insertAt, obj);
  return insertAt;
}

function restoreArchivedLessonV2_(auth, lessonId) {
  const id = String(lessonId || '').trim();
  if (!id) {
    const error = new Error('복원할 수업 기록을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const lessonSheet = getRequiredDataSheetV2_(ss, 'Lessons');
  const assignmentSheet = getRequiredDataSheetV2_(ss, 'LessonAssignments');
  const archiveLessonSheet = ss.getSheetByName(LESSON_ARCHIVE_V2_.LESSON_SHEET);
  const archiveAssignmentSheet = ss.getSheetByName(LESSON_ARCHIVE_V2_.ASSIGNMENT_SHEET);
  if (!archiveLessonSheet) {
    const error = new Error('보관된 수업 기록이 없습니다.');
    error.code = 'LESSON_ARCHIVE_NOT_FOUND';
    throw error;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (lessonGetRowV2_(lessonSheet, id).rowNumber >= 2) {
      const error = new Error('같은 lessonId의 활성 수업 기록이 이미 있습니다.');
      error.code = 'LESSON_RESTORE_COLLISION';
      throw error;
    }

    const archiveRowNumber = lessonArchiveFindRowByLessonIdV2_(archiveLessonSheet, id);
    if (archiveRowNumber < 2) {
      const error = new Error('복원할 보관 수업 기록을 찾을 수 없습니다.');
      error.code = 'LESSON_ARCHIVE_NOT_FOUND';
      throw error;
    }

    const archivedLesson = lessonArchiveReadRowV2_(archiveLessonSheet, archiveRowNumber) || {};
    const lessonObj = {
      lessonId: String(archivedLesson.lessonId || ''),
      date: lessonNormalizeDateV2_(archivedLesson.date),
      classId: String(archivedLesson.classId || ''),
      topic: String(archivedLesson.topic || ''),
      content: String(archivedLesson.content || ''),
      homework: String(archivedLesson.homework || ''),
      nextPlan: String(archivedLesson.nextPlan || ''),
      progressCurrentCount: archivedLesson.progressCurrentCount,
      createdAt: archivedLesson.createdAt || new Date(),
      updatedAt: new Date()
    };
    lessonArchiveInsertRowByDateV2_(lessonSheet, lessonObj);

    let restoredAssignments = 0;
    const archiveAssignmentRows = [];
    if (archiveAssignmentSheet && archiveAssignmentSheet.getLastRow() >= 2) {
      const map = getSheetHeaderMapV2_(archiveAssignmentSheet);
      if (map.lessonId != null) {
        const matches = archiveAssignmentSheet
          .getRange(2, map.lessonId + 1, archiveAssignmentSheet.getLastRow() - 1, 1)
          .createTextFinder(id).matchEntireCell(true).findAll();
        matches.forEach(function(cell) {
          const rowNumber = cell.getRow();
          archiveAssignmentRows.push(rowNumber);
          const row = lessonArchiveReadRowV2_(archiveAssignmentSheet, rowNumber) || {};
          let restoreAssignmentId = String(row.assignmentId || '').trim();
          if (restoreAssignmentId && findDataRowByIdV2_(assignmentSheet, 'assignmentId', restoreAssignmentId) >= 2) {
            restoreAssignmentId = makeNextLessonAssignmentIdIncludingArchiveV2_(ss, assignmentSheet);
          }
          appendObjectRowV2_(assignmentSheet, {
            assignmentId: restoreAssignmentId,
            lessonId: row.lessonId,
            classId: row.classId,
            studentId: row.studentId,
            assignmentText: row.assignmentText,
            createdAt: row.createdAt,
            legacyCreatedAt2: row.legacyCreatedAt2,
            assignment: row.assignment,
            studentName: row.studentName
          });
          restoredAssignments++;
        });
      }
    }

    archiveAssignmentRows.sort(function(a, b) { return b - a; }).forEach(function(rowNumber) {
      archiveAssignmentSheet.deleteRow(rowNumber);
    });
    archiveLessonSheet.deleteRow(archiveRowNumber);

    SpreadsheetApp.flush();
    appendAuditLog_(String(auth.teacher.teacherId || ''), 'LESSON_RESTORE', 'Lesson', id, 'SUCCESS', 'assignments=' + restoredAssignments);

    return {
      version: 'lesson-archive-step8-20261004',
      lessonId: id,
      restoredAssignments: restoredAssignments,
      diagnostics: {
        wholeSheetRewrite: false,
        permanentDelete: false,
        restoreOrder: 'date-aware-insert'
      }
    };
  } finally {
    lock.releaseLock();
  }
}
