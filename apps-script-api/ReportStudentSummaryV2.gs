/**
 * 학생관리 V2 - 학생 종합 조회/출력용 읽기 전용 통합 응답
 * Step 14 / 2026-10-04
 *
 * 원칙
 * - 기존 Attendance / LegacyAttendanceHistory / ScoreHub 전용 시트를 수정하지 않는다.
 * - 기존 통합 출석 함수와 ScoreHub 학생 요약 함수를 재사용한다.
 * - 학생 연결은 현재 Students.studentId를 시작점으로 하되, ScoreHub 내부에서는 전화번호 기준을 유지한다.
 * - TOPIK I WRONG_REVIEW 제외 규칙은 ScoreHubSourceV2의 기존 규칙을 그대로 따른다.
 */

function reportStudentSummaryStringV2_(value) {
  return String(value == null ? '' : value).trim();
}

function reportStudentSummaryDateV2_(value) {
  if (!value) return '';
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return reportStudentSummaryStringV2_(value).split('T')[0].split(' ')[0];
}

function reportStudentSummaryFindStudentV2_(ss, studentId) {
  const sid = reportStudentSummaryStringV2_(studentId);
  if (!sid) {
    const error = new Error('studentId가 필요합니다.');
    error.code = 'STUDENT_ID_REQUIRED';
    throw error;
  }

  const students = readSheetObjects_(ss, 'Students');
  const row = students.find(function(item) {
    return reportStudentSummaryStringV2_(item.studentId) === sid;
  });

  if (!row) {
    const error = new Error('학생 정보를 찾을 수 없습니다: ' + sid);
    error.code = 'STUDENT_NOT_FOUND';
    throw error;
  }

  return row;
}

function reportStudentSummaryFindClassV2_(ss, classId) {
  const cid = reportStudentSummaryStringV2_(classId);
  if (!cid) return null;

  const classes = readSheetObjects_(ss, 'Classes');
  return classes.find(function(item) {
    return reportStudentSummaryStringV2_(item.classId) === cid;
  }) || null;
}

function reportStudentSummaryBasicInfoV2_(studentRow, classRow) {
  return {
    studentId: reportStudentSummaryStringV2_(studentRow.studentId),
    name: reportStudentSummaryStringV2_(studentRow.name),
    phone: reportStudentSummaryStringV2_(studentRow.phone),
    classId: reportStudentSummaryStringV2_(studentRow.classId),
    className: classRow ? reportStudentSummaryStringV2_(classRow.className) : '',
    status: reportStudentSummaryStringV2_(studentRow.status),
    enrollmentDate: reportStudentSummaryDateV2_(studentRow.enrollmentDate),
    stopDate: reportStudentSummaryDateV2_(studentRow.stopDate),
    currentTopikLevel: reportStudentSummaryStringV2_(studentRow.currentTopikLevel),
    targetTopikLevel: reportStudentSummaryStringV2_(studentRow.targetTopikLevel),
    scholarshipType: reportStudentSummaryStringV2_(studentRow.scholarshipType),
    scholarshipStartDate: reportStudentSummaryDateV2_(studentRow.scholarshipStartDate),
    scholarshipEndDate: reportStudentSummaryDateV2_(studentRow.scholarshipEndDate),
    paidUntilDate: reportStudentSummaryDateV2_(studentRow.paidUntilDate),
    note: reportStudentSummaryStringV2_(studentRow.note)
  };
}

function reportStudentSummaryAttendanceRateV2_(summary) {
  const total = Number(summary && summary.total || 0);
  if (!total) return 0;
  const present = Number(summary && summary.present || 0);
  const late = Number(summary && summary.late || 0);
  const early = Number(summary && summary.early || 0);
  return Math.round(((present + late + early) / total) * 1000) / 10;
}

/**
 * 학생 종합 조회의 서버 핵심 함수.
 * 조회 전용이며 어떤 시트도 수정하지 않는다.
 */
function getStudentSummaryReportV2_(auth, studentId, options) {
  options = options || {};

  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const studentRow = reportStudentSummaryFindStudentV2_(ss, studentId);
  const classRow = reportStudentSummaryFindClassV2_(ss, studentRow.classId);

  const attendance = reportAttendanceBuildStudentHistoryV2_(ss, studentId, {
    startDate: options.startDate || '',
    endDate: options.endDate || ''
  });

  // 기존 성적관리와 동일한 데이터 소스를 그대로 사용한다.
  // ScoreHubSourceV2 내부에서 전화번호 기준 연결과 WRONG_REVIEW 제외 규칙을 유지한다.
  const scoreHub = getScoreHubStudentSummaryV2_(auth, studentId);

  const attendanceSummary = {
    total: Number(attendance.summary && attendance.summary.total || 0),
    present: Number(attendance.summary && attendance.summary.present || 0),
    late: Number(attendance.summary && attendance.summary.late || 0),
    absent: Number(attendance.summary && attendance.summary.absent || 0),
    early: Number(attendance.summary && attendance.summary.early || 0)
  };
  attendanceSummary.attendanceRate = reportStudentSummaryAttendanceRateV2_(attendanceSummary);

  return {
    version: 'student-summary-phase1-20261004',
    readOnly: true,
    query: {
      studentId: reportStudentSummaryStringV2_(studentId),
      startDate: reportStudentSummaryDateV2_(options.startDate),
      endDate: reportStudentSummaryDateV2_(options.endDate),
      allHistory: !options.startDate && !options.endDate
    },
    student: reportStudentSummaryBasicInfoV2_(studentRow, classRow),
    attendance: {
      summary: attendanceSummary,
      rows: attendance.rows || [],
      diagnostics: attendance.diagnostics || {}
    },
    learning: {
      progress: scoreHub.progress || null,
      snu: scoreHub.snu || null,
      workbookReading: scoreHub.workbookReading || null,
      workbookListening: scoreHub.workbookListening || null,
      topik1Collocation: scoreHub.topik1Collocation || null,
      topik1Grammar: scoreHub.topik1Grammar || null,
      topik1Reading: scoreHub.topik1Reading || null,
      activity: scoreHub.activity || null,
      syncInfo: scoreHub.syncInfo || {},
      sourceMode: scoreHub.sourceMode || ''
    },
    scope: {
      attendancePeriodApplied: true,
      scorePeriodApplied: false,
      scorePeriodNote: '1단계에서는 기존 ScoreHub 누적 성적/진도를 그대로 재사용합니다. 성적 기간 필터는 다음 단계에서 별도 적용합니다.',
      wrongReviewExcluded: true
    }
  };
}
