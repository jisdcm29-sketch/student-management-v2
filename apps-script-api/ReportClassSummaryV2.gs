/**
 * 학생관리 V2 - 조회/출력 Phase R5 반별 출결 / 수업 진행 조회
 * Step 15C / 2026-10-04
 *
 * 원칙
 * - 읽기 전용.
 * - 현재 V2 Attendance가 같은 학생/같은 날짜의 LegacyAttendanceHistory보다 우선한다.
 * - 현재 반의 재학생만 학생별 출결 요약 대상에 포함한다.
 * - classId + 기간으로 필요한 행만 읽는다.
 * - 기존 Lessons API 계약은 변경하지 않는다.
 */

function reportClassNormalizeDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function reportClassHeadersV2_(sheet) {
  const lastColumn = sheet ? sheet.getLastColumn() : 0;
  if (!sheet || lastColumn < 1) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) {
    return String(v || '').trim();
  });
}

function reportClassHeaderIndexV2_(headers) {
  const map = {};
  (headers || []).forEach(function(h, i) {
    if (h) map[h] = i;
  });
  return map;
}

function reportClassRowObjectV2_(headers, values, rowNumber) {
  const obj = { __sheetRow: rowNumber };
  headers.forEach(function(h, i) {
    if (h) obj[h] = values[i] == null ? '' : values[i];
  });
  return obj;
}

function reportClassFindRowsByExactValueV2_(sheet, headerName, exactValue) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const headers = reportClassHeadersV2_(sheet);
  const index = headers.indexOf(headerName);
  if (index < 0) return [];

  const value = String(exactValue || '').trim();
  if (!value) return [];

  const range = sheet.getRange(2, index + 1, lastRow - 1, 1);
  const matches = range.createTextFinder(value).matchEntireCell(true).findAll();
  if (!matches.length) return [];

  return matches.map(function(cell) {
    const rowNumber = cell.getRow();
    const values = sheet.getRange(rowNumber, 1, 1, lastColumn).getDisplayValues()[0];
    return reportClassRowObjectV2_(headers, values, rowNumber);
  });
}

function reportClassFindClassV2_(ss, classId) {
  const rows = reportClassFindRowsByExactValueV2_(ss.getSheetByName('Classes'), 'classId', classId);
  return rows.length ? rows[0] : null;
}

function reportClassActiveStudentsV2_(ss, classId) {
  return reportClassFindRowsByExactValueV2_(ss.getSheetByName('Students'), 'classId', classId)
    .filter(function(row) { return String(row.status || '').trim() === '재학'; })
    .sort(function(a, b) { return String(a.name || '').localeCompare(String(b.name || ''), 'ko'); });
}

function reportClassInRangeV2_(date, startDate, endDate) {
  const d = reportClassNormalizeDateV2_(date);
  if (!d) return false;
  if (startDate && d < startDate) return false;
  if (endDate && d > endDate) return false;
  return true;
}

function reportClassCompareRecencyV2_(a, b, updatedKey) {
  const av = String(a && a[updatedKey] || '').trim();
  const bv = String(b && b[updatedKey] || '').trim();
  if (av !== bv) return av.localeCompare(bv);
  const ar = Number(a && (a.sourceRow || a.__sheetRow) || 0);
  const br = Number(b && (b.sourceRow || b.__sheetRow) || 0);
  return ar - br;
}

function reportClassLatestByStudentDateV2_(rows, updatedKey) {
  const map = {};
  (rows || []).forEach(function(row) {
    const studentId = String(row.studentId || '').trim();
    const date = reportClassNormalizeDateV2_(row.date);
    if (!studentId || !date) return;
    const key = studentId + '|' + date;
    const found = map[key];
    if (!found || reportClassCompareRecencyV2_(found, row, updatedKey) <= 0) map[key] = row;
  });
  return map;
}

function reportClassStatusCounterV2_() {
  return { total: 0, present: 0, late: 0, absent: 0, early: 0 };
}

function reportClassCountStatusV2_(summary, status) {
  summary.total++;
  const s = String(status || '').trim();
  if (s === '출석') summary.present++;
  else if (s === '지각') summary.late++;
  else if (s === '결석') summary.absent++;
  else if (s === '조퇴') summary.early++;
}

function reportClassRateV2_(summary) {
  const total = Number(summary && summary.total || 0);
  if (!total) return null;
  return Math.round(((Number(summary.present || 0) + Number(summary.late || 0) + Number(summary.early || 0)) / total) * 1000) / 10;
}

function reportClassAttendanceV2_(ss, classId, activeStudents, startDate, endDate) {
  const activeIds = {};
  activeStudents.forEach(function(s) { activeIds[String(s.studentId || '').trim()] = true; });

  const currentRows = reportClassFindRowsByExactValueV2_(ss.getSheetByName('Attendance'), 'classId', classId)
    .filter(function(row) {
      const sid = String(row.studentId || '').trim();
      return sid && sid !== '__HOLIDAY__' && activeIds[sid] && String(row.status || '').trim() !== '휴무' &&
        reportClassInRangeV2_(row.date, startDate, endDate);
    });

  const legacyRows = reportClassFindRowsByExactValueV2_(ss.getSheetByName('LegacyAttendanceHistory'), 'currentClassId', classId)
    .filter(function(row) {
      const sid = String(row.studentId || '').trim();
      return sid && activeIds[sid] && reportClassInRangeV2_(row.date, startDate, endDate);
    });

  const currentMap = reportClassLatestByStudentDateV2_(currentRows, 'updatedAt');
  const legacyMap = reportClassLatestByStudentDateV2_(legacyRows, 'sourceUpdatedAt');
  const merged = {};

  Object.keys(legacyMap).forEach(function(key) {
    const row = legacyMap[key];
    merged[key] = {
      studentId: String(row.studentId || '').trim(),
      date: reportClassNormalizeDateV2_(row.date),
      status: String(row.status || '').trim(),
      memo: String(row.memo || '').trim(),
      source: 'LEGACY'
    };
  });
  Object.keys(currentMap).forEach(function(key) {
    const row = currentMap[key];
    merged[key] = {
      studentId: String(row.studentId || '').trim(),
      date: reportClassNormalizeDateV2_(row.date),
      status: String(row.status || '').trim(),
      memo: String(row.memo || '').trim(),
      source: 'CURRENT'
    };
  });

  const overall = reportClassStatusCounterV2_();
  const perStudent = {};
  activeStudents.forEach(function(s) {
    perStudent[String(s.studentId || '').trim()] = reportClassStatusCounterV2_();
  });

  Object.keys(merged).forEach(function(key) {
    const row = merged[key];
    reportClassCountStatusV2_(overall, row.status);
    if (perStudent[row.studentId]) reportClassCountStatusV2_(perStudent[row.studentId], row.status);
  });

  const studentRows = activeStudents.map(function(s) {
    const sid = String(s.studentId || '').trim();
    const summary = perStudent[sid] || reportClassStatusCounterV2_();
    return {
      studentId: sid,
      name: String(s.name || '').trim(),
      status: String(s.status || '').trim(),
      total: summary.total,
      present: summary.present,
      late: summary.late,
      absent: summary.absent,
      early: summary.early,
      attendanceRate: reportClassRateV2_(summary)
    };
  });

  const attendanceDates = {};
  Object.keys(merged).forEach(function(key) { attendanceDates[merged[key].date] = true; });

  return {
    summary: {
      total: overall.total,
      present: overall.present,
      late: overall.late,
      absent: overall.absent,
      early: overall.early,
      attendanceRate: reportClassRateV2_(overall),
      attendanceDays: Object.keys(attendanceDates).length
    },
    students: studentRows,
    diagnostics: {
      currentMatchedRows: currentRows.length,
      legacyMatchedRows: legacyRows.length,
      currentDedupedRows: Object.keys(currentMap).length,
      legacyDedupedRows: Object.keys(legacyMap).length,
      mergedRows: Object.keys(merged).length
    }
  };
}

function reportClassLessonsV2_(auth, classId, startDate, endDate) {
  const result = getLessonsListV2_(auth, {
    classId: classId,
    startDate: startDate,
    endDate: endDate,
    limit: 100
  }) || {};
  const rows = Array.isArray(result.rows) ? result.rows : [];
  const dayMap = {};
  rows.forEach(function(row) { if (row.date) dayMap[String(row.date)] = true; });
  const latest = rows.length ? rows[0] : null;
  return {
    rows: rows,
    count: rows.length,
    lessonDays: Object.keys(dayMap).length,
    latest: latest ? {
      lessonId: latest.lessonId || '',
      date: latest.date || '',
      topic: latest.topic || '',
      content: latest.content || '',
      progressCurrentCount: latest.progressCurrentCount,
      progressUnit: latest.progressUnit || ''
    } : null,
    hasMore: !!result.hasMore
  };
}

function getClassSummaryReportV2_(auth, classId, options) {
  const id = String(classId || '').trim();
  if (!id) {
    const error = new Error('classId가 필요합니다.');
    error.code = 'CLASS_ID_REQUIRED';
    throw error;
  }

  options = options || {};
  const startDate = reportClassNormalizeDateV2_(options.startDate);
  const endDate = reportClassNormalizeDateV2_(options.endDate);
  if (startDate && endDate && startDate > endDate) {
    const error = new Error('시작일이 종료일보다 늦을 수 없습니다.');
    error.code = 'INVALID_DATE_RANGE';
    throw error;
  }

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classInfo = reportClassFindClassV2_(ss, id);
  if (!classInfo) {
    const error = new Error('반 정보를 찾을 수 없습니다: ' + id);
    error.code = 'CLASS_NOT_FOUND';
    throw error;
  }

  const students = reportClassActiveStudentsV2_(ss, id);
  const attendance = reportClassAttendanceV2_(ss, id, students, startDate, endDate);
  const lessons = reportClassLessonsV2_(auth, id, startDate, endDate);

  return {
    version: 'report-class-summary-r5-20261004',
    readOnly: true,
    classInfo: {
      classId: String(classInfo.classId || '').trim(),
      className: String(classInfo.className || '').trim(),
      teacherName: String(classInfo.teacherName || '').trim(),
      schedule: String(classInfo.schedule || '').trim(),
      classroom: String(classInfo.classroom || '').trim(),
      status: String(classInfo.status || '').trim(),
      targetProgressCount: Number(classInfo.targetProgressCount || 0) || 0,
      targetProgressUnit: String(classInfo.targetProgressUnit || '').trim()
    },
    scope: {
      mode: String(options.mode || 'monthly'),
      startDate: startDate,
      endDate: endDate
    },
    activeStudentCount: students.length,
    attendance: attendance,
    lessons: lessons
  };
}
