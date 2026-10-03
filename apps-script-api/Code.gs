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

    return errorResponse_(requestId, 'NOT_FOUND', '지원하지 않는 API action입니다.');
  } catch (err) {
    return errorResponse_(requestId, err && err.code ? err.code : 'INTERNAL_ERROR', err && err.message ? err.message : String(err));
  }
}
