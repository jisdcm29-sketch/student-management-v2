/**
 * Student Management V2 - Report / Output Phase R6
 * 학생 상태 통계 (읽기 전용)
 * Step 15E / 2026-10-04
 * - Students 시트 1회 읽기만 사용
 * - Classes 시트 재조회 제거
 * - 필요한 필드만 객체화하여 처리량 축소
 */

function reportStatusDateV2_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  var m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  return s.length >= 10 ? s.slice(0, 10) : s;
}

function reportStatusInRangeV2_(value, startDate, endDate) {
  var d = reportStatusDateV2_(value);
  if (!d) return false;
  if (startDate && d < startDate) return false;
  if (endDate && d > endDate) return false;
  return true;
}

function reportStatusIsArchivedV2_(status) {
  var s = String(status || '').trim();
  return s === '휴학' || s === '중단' || s === '중도포기';
}

function reportStatusRoundRateV2_(n, d) {
  n = Number(n || 0); d = Number(d || 0);
  if (!d) return null;
  return Math.round((n / d) * 1000) / 10;
}

function reportStatusReadStudentsFastV2_(ss, classId) {
  var sheet = ss.getSheetByName('Students');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  // 날짜를 문자열로 바로 받아 Date 객체 변환 비용을 줄인다.
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var headers = values[0].map(function(v) { return String(v || '').trim(); });
  var index = {};
  headers.forEach(function(h, i) { if (h) index[h] = i; });

  var required = ['studentId', 'classId', 'name', 'enrollmentDate', 'status', 'stopDate', 'note'];
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var studentId = index.studentId === undefined ? '' : String(row[index.studentId] || '').trim();
    if (!studentId) continue;
    var rowClassId = index.classId === undefined ? '' : String(row[index.classId] || '').trim();
    if (classId && rowClassId !== classId) continue;
    var obj = {};
    required.forEach(function(key) {
      var i = index[key];
      obj[key] = i === undefined ? '' : String(row[i] || '').trim();
    });
    rows.push(obj);
  }
  return rows;
}

function getStudentStatusStatsReportV2_(auth, options) {
  var startedAt = Date.now();
  options = options || {};
  var classId = String(options.classId || '').trim();
  var basis = String(options.basis || 'period_enrolled').trim();
  if (basis !== 'period_enrolled' && basis !== 'all_students') basis = 'period_enrolled';

  var startDate = reportStatusDateV2_(options.startDate);
  var endDate = reportStatusDateV2_(options.endDate);
  if (!startDate || !endDate) {
    var e1 = new Error('시작일과 종료일이 필요합니다.');
    e1.code = 'DATE_RANGE_REQUIRED';
    throw e1;
  }
  if (startDate > endDate) {
    var e2 = new Error('시작일이 종료일보다 늦을 수 없습니다.');
    e2.code = 'INVALID_DATE_RANGE';
    throw e2;
  }

  var ss = getTeacherDataSpreadsheet_(auth.teacher);
  var allRows = reportStatusReadStudentsFastV2_(ss, classId);

  var periodEnrolledRows = allRows.filter(function(row) {
    return reportStatusInRangeV2_(row.enrollmentDate, startDate, endDate);
  });
  var targetRows = basis === 'all_students' ? allRows : periodEnrolledRows;

  var current = { leave: 0, stopped: 0, withdrawn: 0, archived: 0 };
  targetRows.forEach(function(row) {
    var s = String(row.status || '').trim();
    if (s === '휴학') current.leave++;
    else if (s === '중단') current.stopped++;
    else if (s === '중도포기') current.withdrawn++;
    if (reportStatusIsArchivedV2_(s)) current.archived++;
  });

  var changedRows = targetRows.filter(function(row) {
    return reportStatusIsArchivedV2_(row.status) && reportStatusInRangeV2_(row.stopDate, startDate, endDate);
  }).map(function(row) {
    return {
      studentId: String(row.studentId || '').trim(),
      name: String(row.name || '').trim(),
      classId: String(row.classId || '').trim(),
      enrollmentDate: reportStatusDateV2_(row.enrollmentDate),
      status: String(row.status || '').trim(),
      stopDate: reportStatusDateV2_(row.stopDate),
      note: String(row.note || '').trim()
    };
  }).sort(function(a, b) {
    if (a.stopDate !== b.stopDate) return a.stopDate < b.stopDate ? 1 : -1;
    return String(a.name || '').localeCompare(String(b.name || ''), 'ko');
  });

  var targetCount = targetRows.length;
  return {
    version: 'report-student-status-r6-step15e-20261004',
    scope: {
      classId: classId,
      basis: basis,
      startDate: startDate,
      endDate: endDate
    },
    summary: {
      targetCount: targetCount,
      periodEnrolledCount: periodEnrolledRows.length,
      leaveCount: current.leave,
      stoppedCount: current.stopped,
      withdrawnCount: current.withdrawn,
      archivedCount: current.archived,
      changedInPeriodCount: changedRows.length,
      withdrawnRate: reportStatusRoundRateV2_(current.withdrawn, targetCount),
      statusChangeRate: reportStatusRoundRateV2_(changedRows.length, targetCount)
    },
    changedStudents: changedRows,
    diagnostics: {
      sourceRows: allRows.length,
      targetRows: targetRows.length,
      periodEnrolledRows: periodEnrolledRows.length,
      serverElapsedMs: Date.now() - startedAt,
      source: 'Students'
    }
  };
}
