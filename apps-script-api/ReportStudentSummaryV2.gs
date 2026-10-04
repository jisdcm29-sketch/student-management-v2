/**
 * 학생관리 V2 - 학생 종합 조회/출력용 읽기 전용 통합 응답
 * Step 14C / 2026-10-04
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
 * 학생이 속한 반의 가장 최근 수업 1건을 읽기 전용으로 가져온다.
 * LessonServiceV2의 bounded tail-read helper를 재사용하므로 Lessons 전체 시트를 읽지 않는다.
 */
function reportStudentSummaryLatestClassLessonV2_(ss, studentRow, classRow) {
  const classId = reportStudentSummaryStringV2_(studentRow && studentRow.classId);
  if (!classId) return null;

  try {
    const result = lessonReadRecentRowsV2_(ss, { classId: classId, limit: 1 });
    const row = result && result.rows && result.rows.length ? result.rows[0] : null;
    if (!row) return null;

    const currentRaw = row.progressCurrentCount;
    const current = currentRaw === '' || currentRaw === null || currentRaw === undefined
      ? null
      : Number(currentRaw);

    return {
      lessonId: reportStudentSummaryStringV2_(row.lessonId),
      date: reportStudentSummaryDateV2_(row.date),
      classId: classId,
      className: classRow ? reportStudentSummaryStringV2_(classRow.className) : classId,
      topic: reportStudentSummaryStringV2_(row.topic),
      content: reportStudentSummaryStringV2_(row.content),
      homework: reportStudentSummaryStringV2_(row.homework),
      nextPlan: reportStudentSummaryStringV2_(row.nextPlan),
      progressCurrentCount: Number.isFinite(current) ? current : null,
      progressUnit: classRow ? reportStudentSummaryStringV2_(classRow.targetProgressUnit) : '',
      targetProgressCount: classRow && Number.isFinite(Number(classRow.targetProgressCount))
        ? Number(classRow.targetProgressCount)
        : 0
    };
  } catch (e) {
    // 현재 수업 진도는 보조 정보다. 이 조회 실패 때문에 학생 종합조회 전체를 막지 않는다.
    return null;
  }
}



/**
 * Phase R3 - 학생 종합조회용 수업/과제/개별지도 읽기 전용 조회
 * - Lessons: LessonServiceV2의 bounded tail read 재사용 (최대 100건)
 * - LessonAssignments / StudentDailyRecords: studentId 열 TextFinder로 필요한 학생 행만 조회
 * - 원본 시트 수정 없음
 */
function reportStudentSummaryRowsByExactColumnV2_(ss, sheetName, columnName, exactValue) {
  const value = reportStudentSummaryStringV2_(exactValue);
  if (!value) return [];
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(function(v) {
    return reportStudentSummaryStringV2_(v);
  });
  const columnIndex = headers.indexOf(columnName);
  if (columnIndex < 0) return [];

  const finderRange = sheet.getRange(2, columnIndex + 1, sheet.getLastRow() - 1, 1);
  const matches = finderRange.createTextFinder(value).matchEntireCell(true).findAll();
  return matches.map(function(cell) {
    const values = sheet.getRange(cell.getRow(), 1, 1, headers.length).getDisplayValues()[0];
    const obj = {};
    headers.forEach(function(h, i) { if (h) obj[h] = values[i]; });
    return obj;
  });
}

function reportStudentSummaryInDateRangeV2_(value, startDate, endDate) {
  const d = reportStudentSummaryDateV2_(value);
  if (!d) return !startDate && !endDate;
  if (startDate && d < startDate) return false;
  if (endDate && d > endDate) return false;
  return true;
}

function reportStudentSummaryTeachingV2_(ss, studentRow, classRow, options) {
  options = options || {};
  const studentId = reportStudentSummaryStringV2_(studentRow && studentRow.studentId);
  const classId = reportStudentSummaryStringV2_(studentRow && studentRow.classId);
  const startDate = reportStudentSummaryDateV2_(options.startDate);
  const endDate = reportStudentSummaryDateV2_(options.endDate);

  if (!studentId || !classId) {
    return { lessons: [], assignments: [], dailyRecords: [], hasMoreLessons: false };
  }

  let lessonResult = { rows: [], hasMore: false, scannedRows: 0 };
  try {
    lessonResult = lessonReadRecentRowsV2_(ss, {
      classId: classId,
      startDate: startDate,
      endDate: endDate,
      limit: 100
    }) || lessonResult;
  } catch (e) {}

  const lessons = (lessonResult.rows || []).map(function(row) {
    const currentRaw = row.progressCurrentCount;
    const current = currentRaw === '' || currentRaw === null || currentRaw === undefined ? null : Number(currentRaw);
    return {
      lessonId: reportStudentSummaryStringV2_(row.lessonId),
      date: reportStudentSummaryDateV2_(row.date),
      classId: classId,
      className: classRow ? reportStudentSummaryStringV2_(classRow.className) : classId,
      topic: reportStudentSummaryStringV2_(row.topic),
      content: reportStudentSummaryStringV2_(row.content),
      homework: reportStudentSummaryStringV2_(row.homework),
      nextPlan: reportStudentSummaryStringV2_(row.nextPlan),
      progressCurrentCount: Number.isFinite(current) ? current : null,
      progressUnit: classRow ? reportStudentSummaryStringV2_(classRow.targetProgressUnit) : ''
    };
  });

  const lessonById = {};
  lessons.forEach(function(row) { if (row.lessonId) lessonById[row.lessonId] = row; });

  let assignments = [];
  try {
    assignments = reportStudentSummaryRowsByExactColumnV2_(ss, 'LessonAssignments', 'studentId', studentId)
      .filter(function(row) {
        const lessonId = reportStudentSummaryStringV2_(row.lessonId);
        return lessonId && !!lessonById[lessonId] && (!row.classId || reportStudentSummaryStringV2_(row.classId) === classId);
      })
      .map(function(row) {
        const lessonId = reportStudentSummaryStringV2_(row.lessonId);
        const lesson = lessonById[lessonId] || {};
        return {
          assignmentId: reportStudentSummaryStringV2_(row.assignmentId),
          lessonId: lessonId,
          date: lesson.date || '',
          topic: lesson.topic || '',
          assignmentText: reportStudentSummaryStringV2_(row.assignmentText || row.assignment),
          createdAt: reportStudentSummaryStringV2_(row.createdAt)
        };
      })
      .filter(function(row) { return !!row.assignmentText; })
      .sort(function(a, b) { return String(b.date || '').localeCompare(String(a.date || '')); });
  } catch (e) { assignments = []; }

  const assignmentByLesson = {};
  assignments.forEach(function(row) { assignmentByLesson[row.lessonId] = row.assignmentText; });
  lessons.forEach(function(row) { row.studentAssignment = assignmentByLesson[row.lessonId] || ''; });

  let dailyRecords = [];
  try {
    dailyRecords = reportStudentSummaryRowsByExactColumnV2_(ss, 'StudentDailyRecords', 'studentId', studentId)
      .filter(function(row) {
        return (!row.classId || reportStudentSummaryStringV2_(row.classId) === classId) &&
          reportStudentSummaryInDateRangeV2_(row.date, startDate, endDate);
      })
      .map(function(row) {
        return {
          recordId: reportStudentSummaryStringV2_(row.recordId),
          date: reportStudentSummaryDateV2_(row.date),
          classId: reportStudentSummaryStringV2_(row.classId || classId),
          className: classRow ? reportStudentSummaryStringV2_(classRow.className) : classId,
          assignment: reportStudentSummaryStringV2_(row.assignment),
          comment: reportStudentSummaryStringV2_(row.comment),
          nextGuide: reportStudentSummaryStringV2_(row.nextGuide),
          teacherName: classRow ? reportStudentSummaryStringV2_(classRow.teacherName) : '',
          createdAt: reportStudentSummaryStringV2_(row.createdAt),
          updatedAt: reportStudentSummaryStringV2_(row.updatedAt)
        };
      })
      .sort(function(a, b) {
        if (a.date !== b.date) return String(b.date || '').localeCompare(String(a.date || ''));
        return String(b.recordId || '').localeCompare(String(a.recordId || ''));
      });
  } catch (e) { dailyRecords = []; }

  return {
    lessons: lessons,
    assignments: assignments,
    dailyRecords: dailyRecords,
    hasMoreLessons: !!lessonResult.hasMore,
    diagnostics: {
      lessonLimit: 100,
      lessonScannedRows: Number(lessonResult.scannedRows || 0),
      lessonFullSheetRead: false,
      assignmentLookup: 'studentId-textfinder',
      dailyRecordLookup: 'studentId-textfinder'
    }
  };
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

  // 학생이 속한 반의 최신 수업 1건. 화면 비교용 읽기 전용 정보이며 성적 판정에는 아직 사용하지 않는다.
  const latestClassLesson = reportStudentSummaryLatestClassLessonV2_(ss, studentRow, classRow);

  // Phase R3: 선택한 학생의 반 수업/개별 과제/개별 지도 기록을 조회 범위에 맞춰 읽기 전용으로 연결한다.
  const teaching = reportStudentSummaryTeachingV2_(ss, studentRow, classRow, options);

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
    version: 'student-summary-step15a-r3-20261004',
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
    teaching: teaching,
    learning: {
      progress: scoreHub.progress || null,
      classLearning: {
        latestLesson: latestClassLesson
      },
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
      scorePeriodNote: '성적은 기존 ScoreHub 누적 성적/진도를 재사용하며, 수업·과제·개별 지도 기록에는 선택한 조회 기간을 적용합니다.',
      wrongReviewExcluded: true
    }
  };
}
