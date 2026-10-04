/**
 * Student Management V2 - Lesson read-only service (Step 2)
 *
 * Principles:
 * - Lessons are not read during login/bootstrap.
 * - Read only when the lesson page requests data.
 * - Default recent limit: 30.
 * - Read from the sheet tail in bounded chunks instead of loading the whole sheet.
 * - No writes in this step.
 */

const LESSON_READ_V2_ = Object.freeze({
  DEFAULT_LIMIT: 30,
  MAX_LIMIT: 100,
  CHUNK_SIZE: 200,
  MAX_SCAN_ROWS: 4000
});

function lessonNormalizeDateV2_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  }
  return String(value).trim().split('T')[0].split(' ')[0];
}

function lessonNumberV2_(value, fallback) {
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function lessonHeadersV2_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) {
    return String(v || '').trim();
  });
}

function lessonRowToObjectV2_(headers, row) {
  const obj = {};
  headers.forEach(function(h, i) {
    if (h) obj[h] = row[i];
  });
  obj.date = lessonNormalizeDateV2_(obj.date);
  return obj;
}

function lessonClassMapV2_(ss) {
  const map = {};
  readSheetObjects_(ss, 'Classes').forEach(function(row) {
    const id = String(row.classId || '').trim();
    if (!id) return;
    map[id] = {
      classId: id,
      className: String(row.className || ''),
      targetProgressCount: lessonNumberV2_(row.targetProgressCount, 0),
      targetProgressUnit: String(row.targetProgressUnit || '')
    };
  });
  return map;
}

function lessonMatchesV2_(row, options) {
  const classId = String(options.classId || '').trim();
  const startDate = lessonNormalizeDateV2_(options.startDate);
  const endDate = lessonNormalizeDateV2_(options.endDate);
  const rowDate = lessonNormalizeDateV2_(row.date);

  if (classId && String(row.classId || '').trim() !== classId) return false;
  if (startDate && (!rowDate || rowDate < startDate)) return false;
  if (endDate && (!rowDate || rowDate > endDate)) return false;
  return true;
}

function lessonReadRecentRowsV2_(ss, options) {
  options = options || {};
  const sheet = ss.getSheetByName('Lessons');
  if (!sheet) {
    return { rows: [], scannedRows: 0, hasMore: false };
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) {
    return { rows: [], scannedRows: 0, hasMore: false };
  }

  const headers = lessonHeadersV2_(sheet);
  const limit = Math.max(1, Math.min(LESSON_READ_V2_.MAX_LIMIT, lessonNumberV2_(options.limit, LESSON_READ_V2_.DEFAULT_LIMIT)));
  const matched = [];
  let scanned = 0;
  let endRow = lastRow;

  while (endRow >= 2 && scanned < LESSON_READ_V2_.MAX_SCAN_ROWS && matched.length < limit) {
    const remaining = endRow - 1;
    const size = Math.min(LESSON_READ_V2_.CHUNK_SIZE, remaining, LESSON_READ_V2_.MAX_SCAN_ROWS - scanned);
    const startRow = endRow - size + 1;
    const values = sheet.getRange(startRow, 1, size, lastColumn).getDisplayValues();

    for (let i = values.length - 1; i >= 0; i--) {
      scanned++;
      const row = lessonRowToObjectV2_(headers, values[i]);
      if (!String(row.lessonId || '').trim()) continue;
      if (!lessonMatchesV2_(row, options)) continue;
      matched.push(row);
      if (matched.length >= limit) break;
    }

    endRow = startRow - 1;
  }

  const hasMore = endRow >= 2;
  return { rows: matched, scannedRows: scanned, hasMore: hasMore };
}

function getLessonsListV2_(auth, options) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const query = options || {};
  const result = lessonReadRecentRowsV2_(ss, query);
  const classMap = lessonClassMapV2_(ss);

  const rows = result.rows.map(function(row) {
    const classId = String(row.classId || '').trim();
    const classInfo = classMap[classId] || {};
    const current = lessonNumberV2_(row.progressCurrentCount, null);
    const target = lessonNumberV2_(classInfo.targetProgressCount, 0);
    const progressPercent = current !== null && target > 0 ? Math.round((current / target) * 1000) / 10 : null;

    return {
      lessonId: String(row.lessonId || ''),
      date: lessonNormalizeDateV2_(row.date),
      classId: classId,
      className: String(classInfo.className || classId),
      topic: String(row.topic || ''),
      content: String(row.content || ''),
      homework: String(row.homework || ''),
      nextPlan: String(row.nextPlan || ''),
      progressCurrentCount: current,
      progressUnit: String(classInfo.targetProgressUnit || ''),
      targetProgressCount: target,
      progressPercent: progressPercent,
      createdAt: String(row.createdAt || '')
    };
  });

  return {
    version: 'lesson-readonly-step2-20261004',
    readOnly: true,
    query: {
      classId: String(query.classId || ''),
      startDate: lessonNormalizeDateV2_(query.startDate),
      endDate: lessonNormalizeDateV2_(query.endDate),
      limit: Math.max(1, Math.min(LESSON_READ_V2_.MAX_LIMIT, lessonNumberV2_(query.limit, LESSON_READ_V2_.DEFAULT_LIMIT)))
    },
    rows: rows,
    count: rows.length,
    hasMore: result.hasMore,
    diagnostics: {
      scannedRows: result.scannedRows,
      maxScanRows: LESSON_READ_V2_.MAX_SCAN_ROWS,
      chunkSize: LESSON_READ_V2_.CHUNK_SIZE,
      fullSheetRead: false
    }
  };
}

function getPreviousLessonV2_(auth, classId, beforeDate) {
  const id = String(classId || '').trim();
  if (!id) {
    const error = new Error('반을 선택해 주세요.');
    error.code = 'LESSON_CLASS_REQUIRED';
    throw error;
  }

  const date = lessonNormalizeDateV2_(beforeDate) || '9999-12-31';
  const data = getLessonsListV2_(auth, {
    classId: id,
    endDate: date,
    limit: 30
  });

  const previous = data.rows.find(function(row) {
    return !row.date || row.date < date;
  }) || null;

  return {
    version: 'lesson-readonly-step2-20261004',
    readOnly: true,
    classId: id,
    beforeDate: date,
    previousLesson: previous,
    diagnostics: data.diagnostics
  };
}


/**
 * Student Management V2 - Lesson save/update service (Step 3)
 *
 * Write principles:
 * - New lesson: append one row only.
 * - Existing lesson: update only the row identified by lessonId.
 * - Never rewrite the whole Lessons sheet.
 * - LessonAssignments are intentionally NOT written in Step 3.
 */

function lessonAssertClassExistsV2_(ss, classId) {
  const id = String(classId || '').trim();
  if (!id) {
    const error = new Error('반을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const found = readSheetObjects_(ss, 'Classes').some(function(row) {
    return String(row.classId || '').trim() === id;
  });
  if (!found) {
    const error = new Error('선택한 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }
}

function lessonGetRowV2_(sheet, lessonId) {
  const id = String(lessonId || '').trim();
  if (!id) return { rowNumber: -1, row: null };

  const rowNumber = findDataRowByIdV2_(sheet, 'lessonId', id);
  if (rowNumber < 2) return { rowNumber: -1, row: null };

  const lastColumn = sheet.getLastColumn();
  const headers = lessonHeadersV2_(sheet);
  const values = sheet.getRange(rowNumber, 1, 1, lastColumn).getDisplayValues()[0];
  return {
    rowNumber: rowNumber,
    row: lessonRowToObjectV2_(headers, values)
  };
}

function lessonDecorateV2_(ss, row) {
  const classMap = lessonClassMapV2_(ss);
  const classId = String(row.classId || '').trim();
  const classInfo = classMap[classId] || {};
  const current = lessonNumberV2_(row.progressCurrentCount, null);
  const target = lessonNumberV2_(classInfo.targetProgressCount, 0);
  return {
    lessonId: String(row.lessonId || ''),
    date: lessonNormalizeDateV2_(row.date),
    classId: classId,
    className: String(classInfo.className || classId),
    topic: String(row.topic || ''),
    content: String(row.content || ''),
    homework: String(row.homework || ''),
    nextPlan: String(row.nextPlan || ''),
    progressCurrentCount: current,
    progressUnit: String(classInfo.targetProgressUnit || ''),
    targetProgressCount: target,
    progressPercent: current !== null && target > 0 ? Math.round((current / target) * 1000) / 10 : null,
    createdAt: String(row.createdAt || ''),
    updatedAt: String(row.updatedAt || '')
  };
}

function getLessonV2_(auth, lessonId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Lessons');
  const found = lessonGetRowV2_(sheet, lessonId);
  if (found.rowNumber < 2 || !found.row) {
    const error = new Error('수업 기록을 찾을 수 없습니다.');
    error.code = 'LESSON_NOT_FOUND';
    throw error;
  }
  return {
    version: 'lesson-save-step3-20261004',
    lesson: lessonDecorateV2_(ss, found.row)
  };
}

function saveLessonV2_(auth, payload) {
  payload = payload || {};
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Lessons');

  let lessonId = String(payload.lessonId || '').trim();
  const classId = String(payload.classId || '').trim();
  const date = lessonNormalizeDateV2_(payload.date);
  const topic = String(payload.topic || '').trim();
  const content = String(payload.content || '').trim();
  const homework = String(payload.homework || '').trim();
  const nextPlan = String(payload.nextPlan || '').trim();
  const progressRaw = payload.progressCurrentCount;
  const progressCurrentCount = progressRaw === '' || progressRaw == null ? '' : Number(progressRaw);

  lessonAssertClassExistsV2_(ss, classId);

  if (!date) {
    const error = new Error('날짜를 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (!topic) {
    const error = new Error('주제를 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }
  if (progressCurrentCount !== '' && (!Number.isFinite(progressCurrentCount) || progressCurrentCount < 0)) {
    const error = new Error('누적 진도는 0 이상의 숫자로 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let rowNumber = -1;
    let createdAt = new Date();

    if (lessonId) {
      const found = lessonGetRowV2_(sheet, lessonId);
      rowNumber = found.rowNumber;
      if (rowNumber < 2 || !found.row) {
        const error = new Error('수정할 수업 기록을 찾을 수 없습니다.');
        error.code = 'LESSON_NOT_FOUND';
        throw error;
      }
      const headers = getSheetHeaderMapV2_(sheet);
      if (headers.createdAt != null) {
        createdAt = sheet.getRange(rowNumber, headers.createdAt + 1).getValue() || createdAt;
      }
    } else {
      lessonId = makeNextLessonIdIncludingArchiveV2_(ss, sheet);
    }

    const now = new Date();
    const obj = {
      lessonId: lessonId,
      date: date,
      classId: classId,
      topic: topic,
      content: content,
      homework: homework,
      nextPlan: nextPlan,
      progressCurrentCount: progressCurrentCount,
      createdAt: createdAt,
      updatedAt: now
    };

    if (rowNumber >= 2) {
      writeObjectToRowV2_(sheet, rowNumber, obj);
    } else {
      appendObjectRowV2_(sheet, obj);
    }

    SpreadsheetApp.flush();
    appendAuditLog_(
      auth.teacher.teacherId,
      'LESSON_SAVE',
      'Lesson',
      lessonId,
      'SUCCESS',
      rowNumber >= 2 ? 'updated' : 'created'
    );

    return {
      version: 'lesson-save-step3-20261004',
      lessonId: lessonId,
      mode: rowNumber >= 2 ? 'updated' : 'created',
      lesson: lessonDecorateV2_(ss, obj),
      diagnostics: {
        wholeSheetRewrite: false,
        assignmentsWritten: false,
        rowOperation: rowNumber >= 2 ? 'single-row-update' : 'single-row-append'
      }
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Student Management V2 - LessonAssignments service (Step 5A)
 *
 * Separation principle:
 * - LessonAssignments stores per-student homework attached to one lesson.
 * - StudentDailyRecords remains the separate individual coaching log.
 * - This service never writes StudentDailyRecords.
 * - Existing LessonAssignments are updated row-by-row; blank text deletes only that assignment row.
 */

function lessonAssignmentHeaderMapV2_(sheet) {
  const headers = lessonHeadersV2_(sheet);
  const map = {};
  headers.forEach(function(h, i) {
    if (h) map[h] = i;
  });
  return map;
}

function lessonAssignmentRowObjectV2_(sheet, rowNumber) {
  const headers = lessonHeadersV2_(sheet);
  const lastColumn = sheet.getLastColumn();
  if (rowNumber < 2 || lastColumn < 1) return null;
  return lessonRowToObjectV2_(headers, sheet.getRange(rowNumber, 1, 1, lastColumn).getDisplayValues()[0]);
}

function lessonAssignmentRowsForLessonV2_(sheet, lessonId) {
  const id = String(lessonId || '').trim();
  if (!id || sheet.getLastRow() < 2) return [];

  const headerMap = lessonAssignmentHeaderMapV2_(sheet);
  if (headerMap.lessonId == null) {
    const error = new Error('LessonAssignments 시트에 lessonId 열이 없습니다.');
    error.code = 'LESSON_ASSIGNMENT_SCHEMA_ERROR';
    throw error;
  }

  const range = sheet.getRange(2, headerMap.lessonId + 1, sheet.getLastRow() - 1, 1);
  const matches = range.createTextFinder(id).matchEntireCell(true).findAll();
  return matches.map(function(cell) {
    const rowNumber = cell.getRow();
    return { rowNumber: rowNumber, row: lessonAssignmentRowObjectV2_(sheet, rowNumber) || {} };
  });
}

function lessonAssignmentStudentContextV2_(ss, classId) {
  const id = String(classId || '').trim();
  const students = readSheetObjects_(ss, 'Students');
  const map = {};
  students.forEach(function(row) {
    const studentId = String(row.studentId || '').trim();
    if (!studentId) return;
    map[studentId] = {
      studentId: studentId,
      classId: String(row.classId || '').trim(),
      name: String(row.name || ''),
      status: String(row.status || '')
    };
  });
  return { classId: id, studentMap: map };
}

function getLessonAssignmentsV2_(auth, lessonId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const lessonData = getLessonV2_(auth, lessonId);
  const lesson = lessonData.lesson || {};
  const sheet = getRequiredDataSheetV2_(ss, 'LessonAssignments');
  const ctx = lessonAssignmentStudentContextV2_(ss, lesson.classId);
  const found = lessonAssignmentRowsForLessonV2_(sheet, lesson.lessonId);

  const assignments = found.map(function(item) {
    const row = item.row || {};
    const studentId = String(row.studentId || '').trim();
    const student = ctx.studentMap[studentId] || {};
    const text = String(row.assignmentText || row.assignment || '').trim();
    return {
      assignmentId: String(row.assignmentId || ''),
      lessonId: String(row.lessonId || ''),
      classId: String(row.classId || lesson.classId || ''),
      studentId: studentId,
      studentName: String(row.studentName || student.name || studentId || ''),
      assignmentText: text,
      createdAt: String(row.createdAt || ''),
      rowNumber: item.rowNumber
    };
  }).sort(function(a, b) {
    return String(a.studentName || a.studentId).localeCompare(String(b.studentName || b.studentId), 'ko');
  });

  return {
    version: 'lesson-assignment-step5a-20261004',
    lesson: lesson,
    assignments: assignments,
    count: assignments.length,
    diagnostics: {
      fullSheetRead: false,
      lookup: 'lessonId-textfinder',
      dailyRecordsRead: false,
      dailyRecordsWritten: false
    }
  };
}

function saveLessonAssignmentsV2_(auth, payload) {
  payload = payload || {};
  const lessonId = String(payload.lessonId || '').trim();
  const requested = Array.isArray(payload.assignments) ? payload.assignments : [];

  if (!lessonId) {
    const error = new Error('수업 기록을 선택해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const lesson = (getLessonV2_(auth, lessonId).lesson || {});
  const classId = String(lesson.classId || '').trim();
  const sheet = getRequiredDataSheetV2_(ss, 'LessonAssignments');
  const ctx = lessonAssignmentStudentContextV2_(ss, classId);

  const normalized = [];
  const requestedStudentIds = {};
  requested.forEach(function(item) {
    item = item || {};
    const studentId = String(item.studentId || '').trim();
    const assignmentText = String(item.assignmentText != null ? item.assignmentText : item.assignment || '').trim();
    if (!studentId) return;
    if (requestedStudentIds[studentId]) {
      const error = new Error('같은 학생의 과제가 요청에 두 번 포함되어 있습니다: ' + studentId);
      error.code = 'LESSON_ASSIGNMENT_DUPLICATE_REQUEST';
      throw error;
    }
    requestedStudentIds[studentId] = true;

    const student = ctx.studentMap[studentId];
    if (!student) {
      const error = new Error('학생을 찾을 수 없습니다: ' + studentId);
      error.code = 'STUDENT_NOT_FOUND';
      throw error;
    }
    if (String(student.classId || '') !== classId) {
      const error = new Error('선택한 학생은 이 수업의 반에 속하지 않습니다: ' + (student.name || studentId));
      error.code = 'STUDENT_CLASS_MISMATCH';
      throw error;
    }
    normalized.push({
      studentId: studentId,
      studentName: String(student.name || studentId),
      assignmentText: assignmentText
    });
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const existingRows = lessonAssignmentRowsForLessonV2_(sheet, lessonId);
    const existingByStudent = {};
    existingRows.forEach(function(item) {
      const sid = String(item.row.studentId || '').trim();
      if (!sid) return;
      if (existingByStudent[sid]) {
        const error = new Error('기존 LessonAssignments에 같은 수업/학생의 중복 행이 있습니다: ' + sid);
        error.code = 'LESSON_ASSIGNMENT_DUPLICATE_EXISTING';
        throw error;
      }
      existingByStudent[sid] = item;
    });

    let created = 0;
    let updated = 0;
    let deleted = 0;
    const deleteRows = [];
    const resultAssignments = [];

    normalized.forEach(function(item) {
      const existing = existingByStudent[item.studentId] || null;

      if (!item.assignmentText) {
        if (existing) {
          deleteRows.push(existing.rowNumber);
          deleted++;
        }
        return;
      }

      let assignmentId = existing ? String(existing.row.assignmentId || '').trim() : '';
      let createdAt = new Date();
      if (existing) {
        const headers = getSheetHeaderMapV2_(sheet);
        if (headers.createdAt != null) {
          createdAt = sheet.getRange(existing.rowNumber, headers.createdAt + 1).getValue() || createdAt;
        }
      } else {
        assignmentId = makeNextLessonAssignmentIdIncludingArchiveV2_(ss, sheet);
      }

      const obj = {
        assignmentId: assignmentId,
        lessonId: lessonId,
        classId: classId,
        studentId: item.studentId,
        assignmentText: item.assignmentText,
        createdAt: createdAt,
        legacyCreatedAt2: '',
        assignment: item.assignmentText,
        studentName: item.studentName
      };

      if (existing) {
        writeObjectToRowV2_(sheet, existing.rowNumber, obj);
        updated++;
      } else {
        appendObjectRowV2_(sheet, obj);
        created++;
      }

      resultAssignments.push({
        assignmentId: assignmentId,
        lessonId: lessonId,
        classId: classId,
        studentId: item.studentId,
        studentName: item.studentName,
        assignmentText: item.assignmentText
      });
    });

    deleteRows.sort(function(a, b) { return b - a; }).forEach(function(rowNumber) {
      sheet.deleteRow(rowNumber);
    });

    SpreadsheetApp.flush();
    appendAuditLog_(
      auth.teacher.teacherId,
      'LESSON_ASSIGNMENTS_SAVE',
      'LessonAssignments',
      lessonId,
      'SUCCESS',
      'created=' + created + ',updated=' + updated + ',deleted=' + deleted
    );

    return {
      version: 'lesson-assignment-step5a-20261004',
      lessonId: lessonId,
      classId: classId,
      created: created,
      updated: updated,
      deleted: deleted,
      assignments: resultAssignments,
      diagnostics: {
        wholeSheetRewrite: false,
        dailyRecordsWritten: false,
        rowOperationsOnly: true
      }
    };
  } finally {
    lock.releaseLock();
  }
}

