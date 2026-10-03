function normalizeClassStatusV2_(status) {
  const s = String(status || '').trim();
  const lower = s.toLowerCase ? s.toLowerCase() : s;
  if (!s) return '운영중';
  if (s === '종료' || s === '종결' || s === '폐강' || s === '비활성' || lower === 'closed' || lower === 'inactive') return '종료';
  if (s === '예정') return '예정';
  return '운영중';
}

function requireSuperAdminV2_(auth) {
  if (!auth || !auth.teacher || String(auth.teacher.role || '').toUpperCase() !== 'SUPER_ADMIN') {
    const error = new Error('관리자 권한이 필요합니다.');
    error.code = 'FORBIDDEN';
    throw error;
  }
}

function getSheetHeaderMapV2_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function getRequiredDataSheetV2_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    const error = new Error(name + ' 시트를 찾을 수 없습니다.');
    error.code = 'SHEET_NOT_FOUND';
    throw error;
  }
  return sheet;
}

function makeNextPrefixedIdV2_(sheet, idHeader, prefix, digits) {
  const headers = getSheetHeaderMapV2_(sheet);
  if (headers[idHeader] == null) throw new Error(sheet.getName() + ' 시트에 ' + idHeader + ' 열이 없습니다.');
  const lastRow = sheet.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, headers[idHeader] + 1, lastRow - 1, 1).getDisplayValues();
    values.forEach(function(row) {
      const id = String(row[0] || '').trim();
      if (id.indexOf(prefix) !== 0) return;
      const tail = id.slice(prefix.length);
      if (!/^\d+$/.test(tail)) return;
      max = Math.max(max, Number(tail || 0));
    });
  }
  return prefix + String(max + 1).padStart(digits || 3, '0');
}

function findDataRowByIdV2_(sheet, idHeader, idValue) {
  const headers = getSheetHeaderMapV2_(sheet);
  if (headers[idHeader] == null) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, headers[idHeader] + 1, lastRow - 1, 1).getDisplayValues();
  const target = String(idValue || '').trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) return i + 2;
  }
  return -1;
}

function writeObjectToRowV2_(sheet, rowNumber, obj) {
  const headers = getSheetHeaderMapV2_(sheet);
  Object.keys(obj).forEach(function(key) {
    if (headers[key] != null) sheet.getRange(rowNumber, headers[key] + 1).setValue(obj[key]);
  });
}

function appendObjectRowV2_(sheet, obj) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  const row = headers.map(function(header) { return Object.prototype.hasOwnProperty.call(obj, header) ? obj[header] : ''; });
  sheet.appendRow(row);
}

function deleteRowsMatchingV2_(sheet, predicate) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return 0;
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim(); });
  let deleted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const obj = {};
    headers.forEach(function(h, idx) { if (h) obj[h] = values[i][idx]; });
    if (predicate(obj)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}

function summarizeScheduleGroupsV2_(groups) {
  const clean = (Array.isArray(groups) ? groups : []).map(function(group) {
    return {
      days: Array.isArray(group.days) ? group.days.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [],
      startTime: String(group.startTime || '').trim(),
      endTime: String(group.endTime || '').trim()
    };
  }).filter(function(group) { return group.days.length || group.startTime || group.endTime; });

  return {
    groups: clean,
    days: clean.map(function(g) { return g.days.join('/'); }).join(', '),
    startTime: clean.length ? clean[0].startTime : '',
    endTime: clean.length ? clean[0].endTime : ''
  };
}

function replaceClassSchedulesV2_(ss, classId, groups, now) {
  const sheet = getRequiredDataSheetV2_(ss, 'ClassSchedules');
  deleteRowsMatchingV2_(sheet, function(row) { return String(row.classId || '') === String(classId || ''); });

  const clean = summarizeScheduleGroupsV2_(groups).groups;
  clean.forEach(function(group) {
    group.days.forEach(function(day) {
      appendObjectRowV2_(sheet, {
        scheduleId: makeNextPrefixedIdV2_(sheet, 'scheduleId', 'SCH-', 3),
        classId: classId,
        dayOfWeek: day,
        startTime: group.startTime,
        endTime: group.endTime,
        createdAt: now
      });
    });
  });
}

function listClassesV2_(auth, options) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classes = readSheetObjects_(ss, 'Classes');
  const schedules = readSheetObjects_(ss, 'ClassSchedules');
  const includeClosed = !options || options.includeClosed !== false;
  const rows = classes.filter(function(row) {
    return includeClosed || normalizeClassStatusV2_(row.status) !== '종료';
  }).map(function(row) {
    const classId = String(row.classId || '');
    return Object.assign({}, row, {
      status: normalizeClassStatusV2_(row.status),
      schedules: schedules.filter(function(s) { return String(s.classId || '') === classId; })
    });
  });
  rows.sort(function(a, b) { return String(a.className || '').localeCompare(String(b.className || ''), 'ko'); });
  return { classes: rows };
}

function getClassV2_(auth, classId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classes = readSheetObjects_(ss, 'Classes');
  const info = classes.find(function(row) { return String(row.classId || '') === String(classId || ''); });
  if (!info) {
    const error = new Error('반 정보를 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }
  const schedules = readSheetObjects_(ss, 'ClassSchedules').filter(function(row) {
    return String(row.classId || '') === String(classId || '');
  });
  info.status = normalizeClassStatusV2_(info.status);
  return { classInfo: info, schedules: schedules };
}

function saveClassV2_(auth, payload) {
  payload = payload || {};
  const className = String(payload.className || '').trim();
  if (!className) {
    const error = new Error('반 이름을 입력해 주세요.');
    error.code = 'VALIDATION_ERROR';
    throw error;
  }

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classSheet = getRequiredDataSheetV2_(ss, 'Classes');
  const now = new Date();
  let classId = String(payload.classId || '').trim();
  let rowNumber = classId ? findDataRowByIdV2_(classSheet, 'classId', classId) : -1;
  const summary = summarizeScheduleGroupsV2_(payload.scheduleGroups || []);

  let teacherName = String(payload.teacherName || '').trim();
  if (String(auth.teacher.role || '').toUpperCase() !== 'SUPER_ADMIN') teacherName = String(auth.teacher.displayName || '').trim();
  if (!teacherName) teacherName = String(auth.teacher.displayName || '').trim();

  if (classId && rowNumber < 2) {
    const error = new Error('수정할 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }

  if (!classId) classId = makeNextPrefixedIdV2_(classSheet, 'classId', 'C-', 3);

  const obj = {
    classId: classId,
    className: className,
    teacherName: teacherName,
    days: summary.days,
    startTime: summary.startTime,
    endTime: summary.endTime,
    schedule: String(payload.schedule || '').trim(),
    classroom: String(payload.classroom || '').trim(),
    status: normalizeClassStatusV2_(payload.status || '운영중'),
    targetProgressCount: Number(payload.targetProgressCount || 0),
    targetProgressUnit: String(payload.targetProgressUnit || '').trim(),
    createdAt: now
  };

  if (rowNumber >= 2) writeObjectToRowV2_(classSheet, rowNumber, obj);
  else appendObjectRowV2_(classSheet, obj);

  replaceClassSchedulesV2_(ss, classId, summary.groups, now);
  SpreadsheetApp.flush();
  appendAuditLog_(auth.teacher.teacherId, 'CLASS_SAVE', 'Class', classId, 'SUCCESS', rowNumber >= 2 ? 'updated' : 'created');

  return { classId: classId, teacherName: teacherName, classInfo: obj };
}

function closeClassV2_(auth, classId) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const sheet = getRequiredDataSheetV2_(ss, 'Classes');
  const row = findDataRowByIdV2_(sheet, 'classId', classId);
  if (row < 2) {
    const error = new Error('폐강할 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }
  writeObjectToRowV2_(sheet, row, { status: '종료', createdAt: new Date() });
  SpreadsheetApp.flush();
  appendAuditLog_(auth.teacher.teacherId, 'CLASS_CLOSE', 'Class', classId, 'SUCCESS', '');
  return { classId: classId, status: '종료' };
}

function deleteClassV2_(auth, classId) {
  requireSuperAdminV2_(auth);
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classes = readSheetObjects_(ss, 'Classes');
  if (!classes.some(function(row) { return String(row.classId || '') === String(classId || ''); })) {
    const error = new Error('삭제할 반을 찾을 수 없습니다.');
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }

  const students = readSheetObjects_(ss, 'Students').filter(function(row) { return String(row.classId || '') === String(classId || ''); });
  const lessons = readSheetObjects_(ss, 'Lessons').filter(function(row) { return String(row.classId || '') === String(classId || ''); });
  const studentIds = {};
  const lessonIds = {};
  students.forEach(function(row) { studentIds[String(row.studentId || '')] = true; });
  lessons.forEach(function(row) { lessonIds[String(row.lessonId || '')] = true; });

  const assignments = ss.getSheetByName('LessonAssignments');
  deleteRowsMatchingV2_(assignments, function(row) {
    return String(row.classId || '') === String(classId || '') || !!studentIds[String(row.studentId || '')] || !!lessonIds[String(row.lessonId || '')];
  });
  deleteRowsMatchingV2_(ss.getSheetByName('Attendance'), function(row) {
    return String(row.classId || '') === String(classId || '') || !!studentIds[String(row.studentId || '')];
  });
  deleteRowsMatchingV2_(ss.getSheetByName('Scores'), function(row) {
    return String(row.classId || '') === String(classId || '') || !!studentIds[String(row.studentId || '')];
  });
  deleteRowsMatchingV2_(ss.getSheetByName('Lessons'), function(row) { return String(row.classId || '') === String(classId || ''); });
  deleteRowsMatchingV2_(ss.getSheetByName('Students'), function(row) { return String(row.classId || '') === String(classId || ''); });
  deleteRowsMatchingV2_(ss.getSheetByName('ClassSchedules'), function(row) { return String(row.classId || '') === String(classId || ''); });
  deleteRowsMatchingV2_(ss.getSheetByName('Classes'), function(row) { return String(row.classId || '') === String(classId || ''); });

  SpreadsheetApp.flush();
  appendAuditLog_(auth.teacher.teacherId, 'CLASS_DELETE', 'Class', classId, 'SUCCESS', 'Strong delete in teacher dataset');
  return { classId: classId, deleted: true };
}
