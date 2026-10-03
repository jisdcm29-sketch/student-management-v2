const DATA_INTEGRATION_V2_ = Object.freeze({
  LEGACY_STUDENT_MANAGEMENT_ID: '1Y5qoA0mQp-7EAQoXM7GsOF6y3MkdT5XhVEmjGjQc_TY',
  MOBILE_LEARNING_ID: '1y4xaZD8SQUVLztDhBSytvi-_naVCGYX5gTOyZqmhMUE',
  MOBILE_SOURCE_SHEETS: Object.freeze([
    { sheetName: 'TestResults', phoneHeader: 'phone' },
    { sheetName: '워크북읽기평가', phoneHeader: '전화번호' },
    { sheetName: '워크북듣기평가', phoneHeader: '전화번호' },
    { sheetName: '진도현황', phoneHeader: '전화번호' }
  ]),
  ACTIVE_CLASS_STATUS: '운영중',
  ACTIVE_STUDENT_STATUS: '재학'
});

/**
 * Step 1 preview.
 *
 * Read only. No Google Sheet is changed.
 *
 * Match rule is intentionally strict:
 *   1) legacy class status === 운영중
 *   2) legacy student status === 재학
 *   3) legacy student has a phone number
 *   4) the normalized phone exists in the mobile learning spreadsheet
 *
 * Student name is NOT used for matching.
 */
function previewLegacyMobileMatchedStudentsV2() {
  const plan = buildLegacyMobileMigrationPlanV2_();
  const result = {
    mode: 'PREVIEW_ONLY',
    matchedStudentCount: plan.students.length,
    matchedClassCount: plan.classes.length,
    matchedScheduleCount: plan.schedules.length,
    students: plan.students.map(function(item) {
      return {
        studentId: item.studentId,
        classId: item.classId,
        name: item.name,
        phone: normalizeIntegrationPhoneV2_(item.phone)
      };
    }),
    classes: plan.classes.map(function(item) {
      return {
        classId: item.classId,
        className: item.className,
        status: item.status
      };
    })
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Step 1 apply.
 *
 * Imports ONLY matched operational students from the legacy student-management
 * spreadsheet into the currently linked V2 teacher spreadsheet.
 *
 * Safe behavior:
 * - Never deletes existing V2 rows.
 * - Never modifies the mobile source spreadsheet.
 * - Never matches by name.
 * - Existing exact studentId + phone is skipped.
 * - studentId/phone conflicts are reported and skipped.
 * - Only classes that contain at least one matched student are imported.
 * - Only schedules of those imported classes are imported.
 */
function applyLegacyMobileMatchedStudentsV2() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const plan = buildLegacyMobileMigrationPlanV2_();
    const targetSs = getDataIntegrationTargetSpreadsheetV2_();

    const classResult = importMatchedClassesV2_(targetSs, plan.classes);
    const scheduleResult = importMatchedSchedulesV2_(targetSs, plan.schedules);
    const studentResult = importMatchedStudentsV2_(targetSs, plan.students);

    const result = {
      mode: 'APPLY',
      source: {
        legacySpreadsheetId: DATA_INTEGRATION_V2_.LEGACY_STUDENT_MANAGEMENT_ID,
        mobileSpreadsheetId: DATA_INTEGRATION_V2_.MOBILE_LEARNING_ID
      },
      target: {
        spreadsheetId: targetSs.getId(),
        spreadsheetName: targetSs.getName()
      },
      selected: {
        students: plan.students.length,
        classes: plan.classes.length,
        schedules: plan.schedules.length
      },
      imported: {
        students: studentResult.inserted,
        classes: classResult.inserted,
        schedules: scheduleResult.inserted
      },
      skippedExisting: {
        students: studentResult.skippedExisting,
        classes: classResult.skippedExisting,
        schedules: scheduleResult.skippedExisting
      },
      conflicts: {
        students: studentResult.conflicts,
        classes: classResult.conflicts,
        schedules: scheduleResult.conflicts
      }
    };

    console.log(JSON.stringify(result, null, 2));
    return result;
  } finally {
    lock.releaseLock();
  }
}

function buildLegacyMobileMigrationPlanV2_() {
  const legacySs = SpreadsheetApp.openById(DATA_INTEGRATION_V2_.LEGACY_STUDENT_MANAGEMENT_ID);
  const mobileSs = SpreadsheetApp.openById(DATA_INTEGRATION_V2_.MOBILE_LEARNING_ID);

  const legacyClasses = readIntegrationSheetObjectsV2_(legacySs, 'Classes');
  const legacyStudents = readIntegrationSheetObjectsV2_(legacySs, 'Students');
  const legacySchedules = readIntegrationSheetObjectsV2_(legacySs, 'ClassSchedules');

  const activeClassIds = new Set(
    legacyClasses
      .filter(function(item) {
        return String(item.status || '').trim() === DATA_INTEGRATION_V2_.ACTIVE_CLASS_STATUS;
      })
      .map(function(item) { return String(item.classId || '').trim(); })
      .filter(Boolean)
  );

  const mobilePhones = collectMobilePhonesV2_(mobileSs);

  const matchedStudents = legacyStudents.filter(function(student) {
    const classId = String(student.classId || '').trim();
    const status = String(student.status || '').trim();
    const phone = normalizeIntegrationPhoneV2_(student.phone);

    return activeClassIds.has(classId) &&
      status === DATA_INTEGRATION_V2_.ACTIVE_STUDENT_STATUS &&
      !!phone &&
      mobilePhones.has(phone);
  });

  const matchedClassIds = new Set(
    matchedStudents.map(function(student) {
      return String(student.classId || '').trim();
    }).filter(Boolean)
  );

  const matchedClasses = legacyClasses.filter(function(item) {
    return matchedClassIds.has(String(item.classId || '').trim());
  });

  const matchedSchedules = legacySchedules.filter(function(item) {
    return matchedClassIds.has(String(item.classId || '').trim());
  });

  return {
    students: matchedStudents,
    classes: matchedClasses,
    schedules: matchedSchedules,
    mobilePhones: mobilePhones
  };
}

function collectMobilePhonesV2_(mobileSs) {
  const phones = new Set();

  DATA_INTEGRATION_V2_.MOBILE_SOURCE_SHEETS.forEach(function(source) {
    const sheet = mobileSs.getSheetByName(source.sheetName);
    if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return;

    const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
    const headers = values[0].map(function(value) { return String(value || '').trim(); });
    const phoneIndex = headers.indexOf(source.phoneHeader);

    if (phoneIndex < 0) {
      const error = new Error(source.sheetName + ' 시트에서 전화번호 열을 찾을 수 없습니다: ' + source.phoneHeader);
      error.code = 'MOBILE_PHONE_HEADER_NOT_FOUND';
      throw error;
    }

    for (let rowIndex = 1; rowIndex < values.length; rowIndex++) {
      const phone = normalizeIntegrationPhoneV2_(values[rowIndex][phoneIndex]);
      if (phone) phones.add(phone);
    }
  });

  return phones;
}

function getDataIntegrationTargetSpreadsheetV2_() {
  const teacher = findTeacherById_(V2_CONFIG.ADMIN_TEACHER_ID);
  if (!teacher) {
    const error = new Error('V2 관리자 교사 정보를 찾을 수 없습니다.');
    error.code = 'ADMIN_TEACHER_NOT_FOUND';
    throw error;
  }
  return getTeacherDataSpreadsheet_(teacher);
}

function readIntegrationSheetObjectsV2_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const error = new Error('필수 시트를 찾을 수 없습니다: ' + sheetName);
    error.code = 'INTEGRATION_SHEET_NOT_FOUND';
    throw error;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(value) { return String(value || '').trim(); });

  return values.slice(1).filter(function(row) {
    return row.some(function(value) { return String(value || '').trim() !== ''; });
  }).map(function(row, index) {
    const item = {};
    headers.forEach(function(header, columnIndex) {
      if (header) item[header] = row[columnIndex];
    });
    item.__sourceRow = index + 2;
    return item;
  });
}

function normalizeIntegrationPhoneV2_(value) {
  return String(value || '').replace(/\D/g, '');
}

function getTargetSheetInfoV2_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    const error = new Error('V2 대상 시트를 찾을 수 없습니다: ' + sheetName);
    error.code = 'TARGET_SHEET_NOT_FOUND';
    throw error;
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    const error = new Error('V2 대상 시트 헤더가 없습니다: ' + sheetName);
    error.code = 'TARGET_HEADER_NOT_FOUND';
    throw error;
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(value) { return String(value || '').trim(); });

  return { sheet: sheet, headers: headers };
}

function getTargetObjectsV2_(sheetInfo) {
  const sheet = sheetInfo.sheet;
  const headers = sheetInfo.headers;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
  return values.map(function(row, index) {
    const item = { __rowNumber: index + 2 };
    headers.forEach(function(header, columnIndex) {
      if (header) item[header] = row[columnIndex];
    });
    return item;
  }).filter(function(item) {
    return Object.keys(item).some(function(key) {
      return key !== '__rowNumber' && String(item[key] || '').trim() !== '';
    });
  });
}

function rowForTargetHeadersV2_(headers, source) {
  return headers.map(function(header) {
    return Object.prototype.hasOwnProperty.call(source, header) ? source[header] : '';
  });
}

function importMatchedClassesV2_(targetSs, sourceClasses) {
  const info = getTargetSheetInfoV2_(targetSs, 'Classes');
  const existing = getTargetObjectsV2_(info);
  const byId = new Map(existing.map(function(item) {
    return [String(item.classId || '').trim(), item];
  }));

  let inserted = 0;
  let skippedExisting = 0;
  const conflicts = [];

  sourceClasses.forEach(function(source) {
    const classId = String(source.classId || '').trim();
    const current = byId.get(classId);

    if (current) {
      if (String(current.className || '').trim() === String(source.className || '').trim()) {
        skippedExisting++;
      } else {
        conflicts.push({
          classId: classId,
          targetClassName: String(current.className || ''),
          sourceClassName: String(source.className || '')
        });
      }
      return;
    }

    info.sheet.appendRow(rowForTargetHeadersV2_(info.headers, source));
    inserted++;
    byId.set(classId, source);
    appendMigrationAuditV2_(targetSs, 'CLASS', source.__sourceRow, classId, classId, 'IMPORTED', '운영반 + 모바일 전화번호 일치 학생 포함 반');
  });

  return { inserted: inserted, skippedExisting: skippedExisting, conflicts: conflicts };
}

function importMatchedSchedulesV2_(targetSs, sourceSchedules) {
  const info = getTargetSheetInfoV2_(targetSs, 'ClassSchedules');
  const existing = getTargetObjectsV2_(info);
  const byId = new Map(existing.map(function(item) {
    return [String(item.scheduleId || '').trim(), item];
  }));

  let inserted = 0;
  let skippedExisting = 0;
  const conflicts = [];

  sourceSchedules.forEach(function(source) {
    const scheduleId = String(source.scheduleId || '').trim();
    const current = byId.get(scheduleId);

    if (current) {
      const same = String(current.classId || '').trim() === String(source.classId || '').trim() &&
        String(current.dayOfWeek || '').trim() === String(source.dayOfWeek || '').trim() &&
        String(current.startTime || '').trim() === String(source.startTime || '').trim() &&
        String(current.endTime || '').trim() === String(source.endTime || '').trim();

      if (same) skippedExisting++;
      else conflicts.push({ scheduleId: scheduleId, target: current, source: source });
      return;
    }

    info.sheet.appendRow(rowForTargetHeadersV2_(info.headers, source));
    inserted++;
    byId.set(scheduleId, source);
    appendMigrationAuditV2_(targetSs, 'CLASS_SCHEDULE', source.__sourceRow, scheduleId, scheduleId, 'IMPORTED', '연결 대상 운영반 일정');
  });

  return { inserted: inserted, skippedExisting: skippedExisting, conflicts: conflicts };
}

function importMatchedStudentsV2_(targetSs, sourceStudents) {
  const info = getTargetSheetInfoV2_(targetSs, 'Students');
  const existing = getTargetObjectsV2_(info);

  const byStudentId = new Map();
  const byPhone = new Map();
  existing.forEach(function(item) {
    const studentId = String(item.studentId || '').trim();
    const phone = normalizeIntegrationPhoneV2_(item.phone);
    if (studentId) byStudentId.set(studentId, item);
    if (phone) byPhone.set(phone, item);
  });

  let inserted = 0;
  let skippedExisting = 0;
  const conflicts = [];

  sourceStudents.forEach(function(source) {
    const studentId = String(source.studentId || '').trim();
    const phone = normalizeIntegrationPhoneV2_(source.phone);
    const existingById = byStudentId.get(studentId) || null;
    const existingByPhone = byPhone.get(phone) || null;

    if (existingById || existingByPhone) {
      const sameRecord = existingById && existingByPhone &&
        String(existingById.studentId || '').trim() === String(existingByPhone.studentId || '').trim() &&
        normalizeIntegrationPhoneV2_(existingById.phone) === phone;

      if (sameRecord) {
        skippedExisting++;
        return;
      }

      conflicts.push({
        sourceStudentId: studentId,
        sourceName: String(source.name || ''),
        sourcePhone: phone,
        existingStudentIdById: existingById ? String(existingById.studentId || '') : '',
        existingPhoneById: existingById ? normalizeIntegrationPhoneV2_(existingById.phone) : '',
        existingStudentIdByPhone: existingByPhone ? String(existingByPhone.studentId || '') : '',
        existingNameByPhone: existingByPhone ? String(existingByPhone.name || '') : ''
      });
      return;
    }

    const targetSource = Object.assign({}, source, { phone: phone });
    info.sheet.appendRow(rowForTargetHeadersV2_(info.headers, targetSource));
    inserted++;
    byStudentId.set(studentId, targetSource);
    byPhone.set(phone, targetSource);

    appendMigrationAuditV2_(
      targetSs,
      'STUDENT',
      source.__sourceRow,
      studentId,
      studentId,
      'IMPORTED',
      '운영반/재학/전화번호 모바일 정확 일치'
    );
  });

  return { inserted: inserted, skippedExisting: skippedExisting, conflicts: conflicts };
}

function appendMigrationAuditV2_(targetSs, entityType, sourceRow, sourceId, targetId, status, note) {
  const sheet = targetSs.getSheetByName('MigrationAudit');
  if (!sheet) return;

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return;

  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    .map(function(value) { return String(value || '').trim(); });

  const record = {
    entityType: entityType || '',
    sourceRow: sourceRow || '',
    sourceId: sourceId || '',
    targetId: targetId || '',
    status: status || '',
    checksum: '',
    note: note || '',
    verifiedAt: new Date()
  };

  sheet.appendRow(rowForTargetHeadersV2_(headers, record));
}
