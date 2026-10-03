/**
 * Student Management V2 - one-time cleanup of duplicate test classes.
 *
 * Keeps real operating classes:
 *   C-013 = 5시 수업반(2026년8월24)
 *   C-014 = 3시 수업반(20260901)
 *
 * Moves test students:
 *   S-008 -> C-014
 *   S-009 -> C-013
 *   S-010 -> C-013
 *
 * Removes test classes and their schedules:
 *   C-004, C-005
 *
 * No other sheets are modified by this script.
 */

const TEST_CLASS_CLEANUP_V2_ = Object.freeze({
  MOVES: Object.freeze({
    'S-008': 'C-014',
    'S-009': 'C-013',
    'S-010': 'C-013'
  }),
  DELETE_CLASS_IDS: Object.freeze(['C-004', 'C-005']),
  KEEP_CLASS_IDS: Object.freeze(['C-013', 'C-014'])
});

function cleanupTestClassesGetAdminSpreadsheetV2_() {
  const teacher = findTeacherById_(V2_CONFIG.ADMIN_TEACHER_ID);
  if (!teacher) throw new Error('관리자 교사 정보를 찾을 수 없습니다.');
  return getTeacherDataSpreadsheet_(teacher);
}

function cleanupTestClassesReadObjectsV2_(sheet) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim(); });
  return values.slice(1).map(function(row, i) {
    const obj = { __rowNumber: i + 2 };
    headers.forEach(function(h, j) { if (h) obj[h] = row[j]; });
    return obj;
  }).filter(function(obj) {
    return Object.keys(obj).some(function(k) {
      return k !== '__rowNumber' && String(obj[k] || '').trim() !== '';
    });
  });
}

function cleanupTestClassesHeaderMapV2_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    const key = String(h || '').trim();
    if (key) map[key] = i + 1;
  });
  return map;
}

function cleanupTestClassesPreviewV2_() {
  const ss = cleanupTestClassesGetAdminSpreadsheetV2_();
  const studentsSheet = ss.getSheetByName('Students');
  const classesSheet = ss.getSheetByName('Classes');
  const schedulesSheet = ss.getSheetByName('ClassSchedules');
  if (!studentsSheet || !classesSheet || !schedulesSheet) {
    throw new Error('Students / Classes / ClassSchedules 시트를 모두 찾을 수 있어야 합니다.');
  }

  const students = cleanupTestClassesReadObjectsV2_(studentsSheet);
  const classes = cleanupTestClassesReadObjectsV2_(classesSheet);
  const schedules = cleanupTestClassesReadObjectsV2_(schedulesSheet);

  const movePreview = Object.keys(TEST_CLASS_CLEANUP_V2_.MOVES).map(function(studentId) {
    const student = students.find(function(s) { return String(s.studentId || '') === studentId; });
    return {
      studentId: studentId,
      name: student ? String(student.name || '') : '',
      currentClassId: student ? String(student.classId || '') : '',
      targetClassId: TEST_CLASS_CLEANUP_V2_.MOVES[studentId],
      found: !!student
    };
  });

  const classesToDelete = classes.filter(function(c) {
    return TEST_CLASS_CLEANUP_V2_.DELETE_CLASS_IDS.indexOf(String(c.classId || '')) >= 0;
  }).map(function(c) {
    return { classId:c.classId, className:c.className, startTime:c.startTime, endTime:c.endTime, status:c.status };
  });

  const classesToKeep = classes.filter(function(c) {
    return TEST_CLASS_CLEANUP_V2_.KEEP_CLASS_IDS.indexOf(String(c.classId || '')) >= 0;
  }).map(function(c) {
    return { classId:c.classId, className:c.className, startTime:c.startTime, endTime:c.endTime, status:c.status };
  });

  const schedulesToDelete = schedules.filter(function(s) {
    return TEST_CLASS_CLEANUP_V2_.DELETE_CLASS_IDS.indexOf(String(s.classId || '')) >= 0;
  }).map(function(s) {
    return { scheduleId:s.scheduleId, classId:s.classId, dayOfWeek:s.dayOfWeek, startTime:s.startTime, endTime:s.endTime };
  });

  const result = {
    mode:'PREVIEW_ONLY',
    spreadsheetName:ss.getName(),
    studentsToMove:movePreview,
    classesToKeep:classesToKeep,
    classesToDelete:classesToDelete,
    schedulesToDeleteCount:schedulesToDelete.length,
    schedulesToDelete:schedulesToDelete,
    note:'이 함수는 아무 데이터도 수정하지 않습니다.'
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function cleanupTestClassesDeleteRowsBottomUpV2_(sheet, rowNumbers) {
  rowNumbers.slice().sort(function(a,b){ return b-a; }).forEach(function(rowNumber) {
    sheet.deleteRow(rowNumber);
  });
}

function applyCleanupDuplicateTestClassesV2() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = cleanupTestClassesGetAdminSpreadsheetV2_();
    const studentsSheet = ss.getSheetByName('Students');
    const classesSheet = ss.getSheetByName('Classes');
    const schedulesSheet = ss.getSheetByName('ClassSchedules');
    if (!studentsSheet || !classesSheet || !schedulesSheet) {
      throw new Error('Students / Classes / ClassSchedules 시트를 모두 찾을 수 있어야 합니다.');
    }

    const students = cleanupTestClassesReadObjectsV2_(studentsSheet);
    const classes = cleanupTestClassesReadObjectsV2_(classesSheet);
    const schedules = cleanupTestClassesReadObjectsV2_(schedulesSheet);
    const studentHeaders = cleanupTestClassesHeaderMapV2_(studentsSheet);
    if (!studentHeaders.classId || !studentHeaders.studentId) {
      throw new Error('Students 시트에 studentId/classId 열이 없습니다.');
    }

    // Safety: real classes must exist before moving students.
    TEST_CLASS_CLEANUP_V2_.KEEP_CLASS_IDS.forEach(function(classId) {
      const exists = classes.some(function(c) { return String(c.classId || '') === classId; });
      if (!exists) throw new Error('유지 대상 반이 없습니다: ' + classId);
    });

    // Safety: only move the three explicit test students.
    const moved = [];
    Object.keys(TEST_CLASS_CLEANUP_V2_.MOVES).forEach(function(studentId) {
      const student = students.find(function(s) { return String(s.studentId || '') === studentId; });
      if (!student) throw new Error('학생을 찾을 수 없습니다: ' + studentId);
      const targetClassId = TEST_CLASS_CLEANUP_V2_.MOVES[studentId];
      const currentClassId = String(student.classId || '');
      if (currentClassId !== targetClassId) {
        studentsSheet.getRange(student.__rowNumber, studentHeaders.classId).setValue(targetClassId);
      }
      moved.push({ studentId:studentId, name:String(student.name || ''), from:currentClassId, to:targetClassId });
    });

    // Re-read after moving and refuse to delete a test class if any student still belongs to it.
    const studentsAfterMove = cleanupTestClassesReadObjectsV2_(studentsSheet);
    const remaining = studentsAfterMove.filter(function(s) {
      return TEST_CLASS_CLEANUP_V2_.DELETE_CLASS_IDS.indexOf(String(s.classId || '')) >= 0;
    });
    if (remaining.length) {
      throw new Error('삭제 대상 테스트 반에 아직 학생이 남아 있습니다: ' + remaining.map(function(s){ return s.studentId; }).join(', '));
    }

    const scheduleRowsToDelete = schedules.filter(function(s) {
      return TEST_CLASS_CLEANUP_V2_.DELETE_CLASS_IDS.indexOf(String(s.classId || '')) >= 0;
    }).map(function(s) { return s.__rowNumber; });
    cleanupTestClassesDeleteRowsBottomUpV2_(schedulesSheet, scheduleRowsToDelete);

    const classRowsToDelete = classes.filter(function(c) {
      return TEST_CLASS_CLEANUP_V2_.DELETE_CLASS_IDS.indexOf(String(c.classId || '')) >= 0;
    }).map(function(c) { return c.__rowNumber; });
    cleanupTestClassesDeleteRowsBottomUpV2_(classesSheet, classRowsToDelete);

    const result = {
      mode:'APPLY_CLEANUP',
      movedStudents:moved,
      deletedClassIds:TEST_CLASS_CLEANUP_V2_.DELETE_CLASS_IDS.slice(),
      deletedScheduleRows:scheduleRowsToDelete.length,
      keptClassIds:TEST_CLASS_CLEANUP_V2_.KEEP_CLASS_IDS.slice(),
      note:'C-004/C-005 테스트 반과 일정만 삭제했습니다. 다른 시트 데이터는 수정하지 않았습니다.'
    };
    console.log(JSON.stringify(result, null, 2));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function previewCleanupDuplicateTestClassesV2() {
  return cleanupTestClassesPreviewV2_();
}
