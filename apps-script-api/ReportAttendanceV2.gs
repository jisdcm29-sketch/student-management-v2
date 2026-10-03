/**
 * 학생관리 V2 - 현재 출석 + 이전 출석 이력 통합 조회 준비
 * Step 10 / 2026-10-04
 *
 * 목적
 * - 기존 Attendance는 수정하지 않는다.
 * - LegacyAttendanceHistory도 수정하지 않는다.
 * - 학생 1명의 출석 이력을 조회할 때 두 시트를 읽어서 화면용 데이터만 합친다.
 * - 같은 학생/같은 날짜가 양쪽에 있으면 현재 V2 Attendance를 우선한다.
 * - LegacyAttendanceHistory 내부에서 같은 날짜가 여러 번 있으면 가장 최근 sourceUpdatedAt/sourceRow를 사용한다.
 */

const REPORT_ATTENDANCE_V2_PREVIEW_SPREADSHEET_ID_ = '1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q';
const REPORT_ATTENDANCE_V2_PREVIEW_STUDENT_ID_ = 'S-009'; // 운마랄

function reportAttendanceNormalizeDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '').trim().split('T')[0].split(' ')[0];
}

function reportAttendanceReadSheetObjectsIfExistsV2_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(value) {
    return String(value || '').trim();
  });

  return values.slice(1).filter(function(row) {
    return row.some(function(value) { return String(value || '').trim() !== ''; });
  }).map(function(row, index) {
    const obj = { __sheetRow: index + 2 };
    headers.forEach(function(header, columnIndex) {
      if (header) obj[header] = row[columnIndex] == null ? '' : row[columnIndex];
    });
    return obj;
  });
}

function reportAttendanceCompareRecencyV2_(a, b, updatedKey) {
  const av = String(a && a[updatedKey] || '').trim();
  const bv = String(b && b[updatedKey] || '').trim();
  if (av !== bv) return av.localeCompare(bv);

  const ar = Number(a && (a.sourceRow || a.__sheetRow) || 0);
  const br = Number(b && (b.sourceRow || b.__sheetRow) || 0);
  return ar - br;
}

function reportAttendanceSelectLatestByDateV2_(rows, updatedKey) {
  const byDate = {};

  (rows || []).forEach(function(row) {
    const date = reportAttendanceNormalizeDateV2_(row.date);
    if (!date) return;

    const found = byDate[date];
    if (!found || reportAttendanceCompareRecencyV2_(found, row, updatedKey) <= 0) {
      byDate[date] = row;
    }
  });

  return byDate;
}

function reportAttendanceBuildStudentHistoryV2_(ss, studentId, options) {
  options = options || {};
  const id = String(studentId || '').trim();
  if (!id) throw new Error('studentId가 필요합니다.');

  const students = reportAttendanceReadSheetObjectsIfExistsV2_(ss, 'Students');
  const student = students.find(function(row) {
    return String(row.studentId || '').trim() === id;
  });

  if (!student) {
    const error = new Error('학생 정보를 찾을 수 없습니다: ' + id);
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }

  const startDate = reportAttendanceNormalizeDateV2_(options.startDate);
  const endDate = reportAttendanceNormalizeDateV2_(options.endDate);

  const legacyRows = reportAttendanceReadSheetObjectsIfExistsV2_(ss, 'LegacyAttendanceHistory')
    .filter(function(row) {
      return String(row.studentId || '').trim() === id;
    });

  const currentRows = reportAttendanceReadSheetObjectsIfExistsV2_(ss, 'Attendance')
    .filter(function(row) {
      return String(row.studentId || '').trim() === id &&
        String(row.studentId || '').trim() !== '__HOLIDAY__' &&
        String(row.status || '').trim() !== '휴무';
    });

  const legacyByDate = reportAttendanceSelectLatestByDateV2_(legacyRows, 'sourceUpdatedAt');
  const currentByDate = reportAttendanceSelectLatestByDateV2_(currentRows, 'updatedAt');
  const dateKeys = {};
  Object.keys(legacyByDate).forEach(function(date) { dateKeys[date] = true; });
  Object.keys(currentByDate).forEach(function(date) { dateKeys[date] = true; });

  let mergedRows = Object.keys(dateKeys).map(function(date) {
    const current = currentByDate[date];
    if (current) {
      return {
        date: date,
        status: String(current.status || '').trim(),
        memo: String(current.memo || '').trim(),
        source: 'CURRENT',
        sourceLabel: '현재 V2',
        currentClassId: String(current.classId || '').trim(),
        sourceClassId: '',
        sourceStudentId: '',
        updatedAt: String(current.updatedAt || '').trim()
      };
    }

    const legacy = legacyByDate[date];
    return {
      date: date,
      status: String(legacy.status || '').trim(),
      memo: String(legacy.memo || '').trim(),
      source: 'LEGACY',
      sourceLabel: '이전 학생관리',
      currentClassId: String(legacy.currentClassId || student.classId || '').trim(),
      sourceClassId: String(legacy.sourceClassId || '').trim(),
      sourceStudentId: String(legacy.sourceStudentId || '').trim(),
      updatedAt: String(legacy.sourceUpdatedAt || '').trim()
    };
  });

  mergedRows = mergedRows.filter(function(row) {
    if (startDate && row.date < startDate) return false;
    if (endDate && row.date > endDate) return false;
    return true;
  }).sort(function(a, b) {
    return b.date.localeCompare(a.date);
  });

  const summary = {
    total: mergedRows.length,
    present: 0,
    late: 0,
    absent: 0,
    early: 0
  };

  mergedRows.forEach(function(row) {
    const status = String(row.status || '').trim();
    if (status === '지각') summary.late++;
    else if (status === '결석') summary.absent++;
    else if (status === '조퇴') summary.early++;
    else if (status === '출석') summary.present++;
  });

  return {
    student: {
      studentId: String(student.studentId || '').trim(),
      name: String(student.name || '').trim(),
      phone: String(student.phone || '').trim(),
      classId: String(student.classId || '').trim(),
      status: String(student.status || '').trim()
    },
    summary: summary,
    rows: mergedRows,
    diagnostics: {
      rawLegacyRows: legacyRows.length,
      dedupedLegacyDates: Object.keys(legacyByDate).length,
      rawCurrentRows: currentRows.length,
      dedupedCurrentDates: Object.keys(currentByDate).length,
      currentOverridesLegacyDates: Object.keys(currentByDate).filter(function(date) {
        return !!legacyByDate[date];
      }).length
    }
  };
}

/**
 * 향후 API 라우트에서 사용할 함수.
 * 현재 Step 10에서는 Code.gs에 아직 연결하지 않는다.
 */
function getStudentAttendanceHistoryV2_(auth, studentId, options) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  return reportAttendanceBuildStudentHistoryV2_(ss, studentId, options || {});
}

/**
 * Step 10 수동 검증 함수.
 * Apps Script 편집기에서 ReportAttendanceV2.gs를 열고 이 함수를 실행한다.
 */
function previewCombinedAttendanceHistoryV2() {
  const ss = SpreadsheetApp.openById(REPORT_ATTENDANCE_V2_PREVIEW_SPREADSHEET_ID_);
  const result = reportAttendanceBuildStudentHistoryV2_(
    ss,
    REPORT_ATTENDANCE_V2_PREVIEW_STUDENT_ID_,
    {}
  );
  console.log(JSON.stringify(result, null, 2));
  return result;
}
