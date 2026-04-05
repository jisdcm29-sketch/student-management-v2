/**
 * 신규 파일: V2_Consult_Report_Service.gs
 * 상담 리포트 생성 서비스
 */

function V2_ConsultReportService_validateStudentId_(studentId) {
  try {
    studentId = String(studentId || '').trim();

    if (!studentId) {
      throw new Error('studentId는 필수입니다.');
    }

    return studentId;
  } catch (error) {
    throw new Error('V2_ConsultReportService_validateStudentId_ 오류: ' + error.message);
  }
}

function V2_ConsultReportService_buildReport_(studentId) {
  try {
    var student = V2_ConsultReportRepository_getStudentById(studentId);

    if (!student) {
      throw new Error('학생을 찾을 수 없습니다.');
    }

    var attendance = V2_ConsultReportRepository_getAttendanceByStudentId(studentId);
    var scores = V2_ConsultReportRepository_getScoresByStudentId(studentId);
    var lessons = V2_ConsultReportRepository_getLessonsByStudentId(studentId);
    var comments = V2_ConsultReportRepository_getCommentsByStudentId(studentId);
    var absenceContacts = V2_ConsultReportRepository_getAbsenceContactsByStudentId(studentId);

    return {
      student: student,
      attendance: attendance,
      scores: scores,
      lessons: lessons,
      comments: comments,
      absenceContacts: absenceContacts,
      generatedAt: V2_nowText_(),
      teacherName: V2_getCurrentUserName_()
    };
  } catch (error) {
    throw new Error('V2_ConsultReportService_buildReport_ 오류: ' + error.message);
  }
}

function V2_ConsultReportService_getReport(studentId) {
  try {
    var validStudentId = V2_ConsultReportService_validateStudentId_(studentId);
    var report = V2_ConsultReportService_buildReport_(validStudentId);

    return V2_createSuccessResponse_(report);
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportService_getReport', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}
