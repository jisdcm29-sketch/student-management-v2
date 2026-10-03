function doGet(e) {
  const requestId = newRequestId_();
  try {
    const action = String(e && e.parameter && e.parameter.action || 'health');

    if (action === 'health') {
      return okResponse_(requestId, {
        app: V2_CONFIG.APP_NAME,
        apiVersion: V2_CONFIG.API_VERSION,
        status: 'ok',
        time: new Date().toISOString()
      });
    }

    return errorResponse_(requestId, 'NOT_FOUND', '지원하지 않는 GET action입니다.');
  } catch (e2) {
    return errorResponse_(requestId, e2.code || 'INTERNAL_ERROR', e2.message || String(e2));
  }
}

function doPost(e) {
  const requestId = newRequestId_();

  try {
    let body = {};
    if (e && e.postData && e.postData.contents) body = JSON.parse(e.postData.contents);
    const action = String(body.action || '').trim();

    if (action === 'login') return okResponse_(requestId, loginTeacher_(body.teacherId, body.password));
    if (action === 'logout') return okResponse_(requestId, logoutTeacher_(body.sessionToken));

    const auth = requireSession_(body.sessionToken);

    if (action === 'session') {
      return okResponse_(requestId, {
        valid: true,
        teacher: {
          teacherId: auth.teacher.teacherId,
          displayName: auth.teacher.displayName || '',
          role: auth.teacher.role || 'TEACHER'
        }
      });
    }
    if (action === 'bootstrap') return okResponse_(requestId, getBootstrapData_(auth));
    if (action === 'classes.list') return okResponse_(requestId, listClassesV2_(auth, body.options || {}));
    if (action === 'classes.get') return okResponse_(requestId, getClassV2_(auth, body.classId));
    if (action === 'classes.save') return okResponse_(requestId, saveClassV2_(auth, body.classData || {}));
    if (action === 'classes.close') return okResponse_(requestId, closeClassV2_(auth, body.classId));
    if (action === 'classes.delete') return okResponse_(requestId, deleteClassV2_(auth, body.classId));
    if (action === 'students.list') return okResponse_(requestId, listStudentsV2_(auth, body.options || {}));
    if (action === 'students.get') return okResponse_(requestId, getStudentV2_(auth, body.studentId));
    if (action === 'students.save') return okResponse_(requestId, saveStudentV2_(auth, body.studentData || {}));
    if (action === 'students.archived.list') return okResponse_(requestId, listArchivedStudentsV2_(auth, body.options || {}));
    if (action === 'students.restore') return okResponse_(requestId, restoreStudentV2_(auth, body.studentId));
    if (action === 'dailyRecords.list') return okResponse_(requestId, listDailyRecordsV2_(auth, body.options || {}));
    if (action === 'dailyRecords.get') return okResponse_(requestId, getDailyRecordV2_(auth, body.recordId));
    if (action === 'dailyRecords.save') return okResponse_(requestId, saveDailyRecordV2_(auth, body.recordData || {}));
    if (action === 'dailyRecords.delete') return okResponse_(requestId, deleteDailyRecordV2_(auth, body.recordId));
    if (action === 'attendance.load') return okResponse_(requestId, loadAttendanceV2_(auth, body.options || {}));
    if (action === 'attendance.save') return okResponse_(requestId, saveAttendanceV2_(auth, body.attendanceData || {}));
    if (action === 'attendance.holiday.save') return okResponse_(requestId, saveAttendanceHolidayV2_(auth, body.holidayData || {}));
    if (action === 'attendance.holiday.clear') return okResponse_(requestId, clearAttendanceHolidayV2_(auth, body.holidayData || {}));
    if (action === 'scores.students') return okResponse_(requestId, listScoreStudentsV2_(auth, body.classId));
    if (action === 'scores.list') return okResponse_(requestId, listScoresV2_(auth, body.options || {}));
    if (action === 'scores.get') return okResponse_(requestId, getScoreV2_(auth, body.recordId));
    if (action === 'scores.save') return okResponse_(requestId, saveScoreV2_(auth, body.scoreData || {}));
    if (action === 'scores.delete') return okResponse_(requestId, deleteScoreV2_(auth, body.recordId));
    if (action === 'qrAttendance.setup') return okResponse_(requestId, getQrAttendanceSetupV2_(auth, body.options || {}));
    if (action === 'qrAttendance.start') return okResponse_(requestId, startQrAttendanceSessionV2_(auth, body.sessionData || {}));
    if (action === 'qrAttendance.close') return okResponse_(requestId, closeQrAttendanceSessionV2_(auth, body.sessionData || {}));

    return errorResponse_(requestId, 'NOT_FOUND', '지원하지 않는 API action입니다.');
  } catch (err) {
    return errorResponse_(requestId, err && err.code ? err.code : 'INTERNAL_ERROR', err && err.message ? err.message : String(err));
  }
}
