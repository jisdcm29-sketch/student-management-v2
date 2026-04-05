/**
 * V2_Student_Service.gs
 * V2 학생 비즈니스 로직 전용 Service
 * 검증 / 상태 변경 / 외부 노출 함수 담당
 */

function V2_StudentService_getAllowedStatuses_() {
  try {
    return [
      V2_CONFIG.STUDENT_STATUS.ACTIVE,
      V2_CONFIG.STUDENT_STATUS.ARCHIVED,
      V2_CONFIG.STUDENT_STATUS.DELETE_PENDING,
      V2_CONFIG.STUDENT_STATUS.DELETED
    ];
  } catch (error) {
    throw new Error('V2_StudentService_getAllowedStatuses_ 오류: ' + error.message);
  }
}

function V2_StudentService_normalizeInput_(input) {
  try {
    input = input || {};

    return {
      studentName: String(input.studentName || '').trim(),
      classId: String(input.classId || '').trim(),
      status: String(input.status || V2_CONFIG.STUDENT_STATUS.ACTIVE).trim(),
      parentName: String(input.parentName || '').trim(),
      parentPhone: String(input.parentPhone || '').trim(),
      notes: String(input.notes || '').trim()
    };
  } catch (error) {
    throw new Error('V2_StudentService_normalizeInput_ 오류: ' + error.message);
  }
}

function V2_StudentService_validateStudentInput_(input) {
  try {
    if (!input.studentName) {
      throw new Error('학생 이름은 필수입니다.');
    }

    if (!input.classId) {
      throw new Error('classId는 필수입니다.');
    }

    if (V2_StudentService_getAllowedStatuses_().indexOf(input.status) === -1) {
      throw new Error('허용되지 않은 학생 상태입니다: ' + input.status);
    }
  } catch (error) {
    throw new Error('V2_StudentService_validateStudentInput_ 오류: ' + error.message);
  }
}

function V2_StudentService_buildStudentEntity_(input) {
  try {
    var now = V2_nowText_();

    return {
      studentId: V2_StudentService_generateStudentId_(),
      studentName: input.studentName,
      classId: input.classId,
      status: input.status,
      parentName: input.parentName,
      parentPhone: input.parentPhone,
      notes: input.notes,
      createdAt: now,
      updatedAt: now
    };
  } catch (error) {
    throw new Error('V2_StudentService_buildStudentEntity_ 오류: ' + error.message);
  }
}

function V2_StudentService_generateStudentId_() {
  try {
    var students = V2_StudentRepository_findAllStudents();
    var maxNumber = 0;

    for (var i = 0; i < students.length; i++) {
      var student = students[i] || {};
      var studentId = String(student.studentId || '').trim();
      var match = studentId.match(/^STU_TEST_(\d+)$/i);

      if (!match) {
        continue;
      }

      maxNumber = Math.max(maxNumber, Number(match[1] || 0));
    }

    return 'STU_TEST_' + ('000' + (maxNumber + 1)).slice(-3);
  } catch (error) {
    throw new Error('V2_StudentService_generateStudentId_ 오류: ' + error.message);
  }
}

function V2_createStudent(input) {
  try {
    var normalized = V2_StudentService_normalizeInput_(input);
    V2_StudentService_validateStudentInput_(normalized);

    var duplicated = V2_StudentRepository_existsActiveStudentByNameAndClass(
      normalized.studentName,
      normalized.classId
    );

    if (duplicated) {
      throw new Error('같은 반에 동일한 이름의 학생이 이미 존재합니다.');
    }

    var student = V2_StudentService_buildStudentEntity_(normalized);
    var saved = V2_StudentRepository_insertStudent(student);

    V2_log_(
      'INFO',
      'V2_createStudent',
      '학생 등록 완료',
      JSON.stringify({
        studentId: saved.studentId,
        studentName: saved.studentName,
        classId: saved.classId
      })
    );

    return V2_createSuccessResponse_(saved);
  } catch (error) {
    V2_log_('ERROR', 'V2_createStudent', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_getStudentList(includeDeleted) {
  try {
    var students = V2_StudentRepository_findAllStudents();

    if (!includeDeleted) {
      students = students.filter(function(student) {
        return String(student.status) !== V2_CONFIG.STUDENT_STATUS.DELETED;
      });
    }

    students.sort(function(a, b) {
      return String(a.studentName).localeCompare(String(b.studentName), 'ko');
    });

    return V2_createSuccessResponse_(students);
  } catch (error) {
    V2_log_('ERROR', 'V2_getStudentList', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

/**
 * 학생 목록 팝업 전용 함수
 * 현재 학생 목록 JS에서 직접 호출하는 함수명과 반드시 일치해야 한다.
 */
function V2_StudentService_getStudentListForPopup(includeDeleted) {
  try {
    var result = V2_getStudentList(includeDeleted);

    return {
      success: result && result.success === true,
      message: result && result.message ? result.message : '학생 목록 조회 완료',
      data: result && Array.isArray(result.data) ? result.data : []
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_StudentService_getStudentListForPopup', error.message, error.stack || '');
    return {
      success: false,
      message: error.message || '학생 목록 팝업 조회 중 오류가 발생했습니다.',
      data: []
    };
  }
}

function V2_getStudentsByClassId(classId, includeDeleted) {
  try {
    classId = String(classId || '').trim();

    if (!classId) {
      throw new Error('classId는 필수입니다.');
    }

    var students = V2_StudentRepository_findStudentsBy_(function(student) {
      var sameClass = String(student.classId) === classId;
      var allowed = includeDeleted ? true : String(student.status) !== V2_CONFIG.STUDENT_STATUS.DELETED;
      return sameClass && allowed;
    });

    students.sort(function(a, b) {
      return String(a.studentName).localeCompare(String(b.studentName), 'ko');
    });

    return V2_createSuccessResponse_(students);
  } catch (error) {
    V2_log_('ERROR', 'V2_getStudentsByClassId', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_getStudentById(studentId) {
  try {
    studentId = String(studentId || '').trim();

    if (!studentId) {
      throw new Error('studentId는 필수입니다.');
    }

    var student = V2_StudentRepository_findStudentById(studentId);

    if (!student) {
      throw new Error('학생을 찾을 수 없습니다.');
    }

    return V2_createSuccessResponse_(student);
  } catch (error) {
    V2_log_('ERROR', 'V2_getStudentById', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_updateStudent(studentId, input) {
  try {
    studentId = String(studentId || '').trim();

    if (!studentId) {
      throw new Error('studentId는 필수입니다.');
    }

    var current = V2_StudentRepository_findStudentById(studentId);

    if (!current) {
      throw new Error('수정 대상 학생을 찾을 수 없습니다.');
    }

    input = input || {};

    var normalized = {
      studentName: input.studentName !== undefined ? String(input.studentName).trim() : String(current.studentName).trim(),
      classId: input.classId !== undefined ? String(input.classId).trim() : String(current.classId).trim(),
      status: input.status !== undefined ? String(input.status).trim() : String(current.status).trim(),
      parentName: input.parentName !== undefined ? String(input.parentName).trim() : String(current.parentName).trim(),
      parentPhone: input.parentPhone !== undefined ? String(input.parentPhone).trim() : String(current.parentPhone).trim(),
      notes: input.notes !== undefined ? String(input.notes).trim() : String(current.notes).trim()
    };

    V2_StudentService_validateStudentInput_(normalized);

    var duplicated = V2_StudentRepository_findStudentsBy_(function(student) {
      if (String(student.studentId) === studentId) {
        return false;
      }

      return String(student.studentName).trim() === normalized.studentName &&
             String(student.classId).trim() === normalized.classId &&
             String(student.status).trim() !== V2_CONFIG.STUDENT_STATUS.DELETED;
    });

    if (duplicated.length > 0) {
      throw new Error('같은 반에 동일한 이름의 다른 학생이 이미 존재합니다.');
    }

    var updated = Object.assign({}, current, normalized, {
      updatedAt: V2_nowText_()
    });

    var saved = V2_StudentRepository_updateStudentById(studentId, updated);

    V2_log_(
      'INFO',
      'V2_updateStudent',
      '학생 정보 수정 완료',
      JSON.stringify({
        studentId: saved.studentId,
        studentName: saved.studentName,
        classId: saved.classId,
        status: saved.status
      })
    );

    return V2_createSuccessResponse_(saved);
  } catch (error) {
    V2_log_('ERROR', 'V2_updateStudent', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_updateStudentStatus(studentId, newStatus) {
  try {
    studentId = String(studentId || '').trim();
    newStatus = String(newStatus || '').trim();

    if (!studentId) {
      throw new Error('studentId는 필수입니다.');
    }

    if (V2_StudentService_getAllowedStatuses_().indexOf(newStatus) === -1) {
      throw new Error('허용되지 않은 상태값입니다: ' + newStatus);
    }

    var current = V2_StudentRepository_findStudentById(studentId);

    if (!current) {
      throw new Error('상태 변경 대상 학생을 찾을 수 없습니다.');
    }

    var updated = V2_StudentRepository_patchStudentById(studentId, {
      status: newStatus,
      updatedAt: V2_nowText_()
    });

    V2_log_(
      'INFO',
      'V2_updateStudentStatus',
      '학생 상태 변경 완료',
      JSON.stringify({
        studentId: studentId,
        oldStatus: current.status,
        newStatus: newStatus
      })
    );

    return V2_createSuccessResponse_(updated);
  } catch (error) {
    V2_log_('ERROR', 'V2_updateStudentStatus', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_archiveStudent(studentId) {
  return V2_updateStudentStatus(studentId, V2_CONFIG.STUDENT_STATUS.ARCHIVED);
}

function V2_markStudentDeletePending(studentId) {
  return V2_updateStudentStatus(studentId, V2_CONFIG.STUDENT_STATUS.DELETE_PENDING);
}

function V2_markStudentDeleted(studentId) {
  return V2_updateStudentStatus(studentId, V2_CONFIG.STUDENT_STATUS.DELETED);
}

function V2_getStudentFormClassList() {
  try {
    var items = V2_StudentService_getClassListItems_();

    return V2_createSuccessResponse_({
      items: items
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_getStudentFormClassList', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_StudentService_getClassListItems_() {
  try {
    var spreadsheet = V2_getSpreadsheet_();
    var classesSheetName = V2_CONFIG && V2_CONFIG.SHEETS && V2_CONFIG.SHEETS.CLASSES
      ? V2_CONFIG.SHEETS.CLASSES
      : 'V2_Classes';

    var sheet = spreadsheet.getSheetByName(classesSheetName);
    if (!sheet) {
      return [];
    }

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) {
      return [];
    }

    var headers = values[0];
    var items = [];
    var seen = {};

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowObject = V2_StudentService_rowToObjectByHeaders_(headers, row);
      var classId = String(V2_StudentService_pickValue_(rowObject, ['classid', 'id', '반id', '반코드']) || '').trim();
      var className = String(V2_StudentService_pickValue_(rowObject, ['classname', 'name', '반명']) || '').trim();
      var status = String(V2_StudentService_pickValue_(rowObject, ['status', '상태', 'classstatus']) || '').trim();

      if (!classId && !className) {
        continue;
      }

      if (!classId) {
        classId = className;
      }

      if (status && (status === '삭제' || status === '삭제대기' || status === '완전삭제')) {
        continue;
      }

      if (seen[classId]) {
        continue;
      }

      seen[classId] = true;

      items.push({
        classId: classId,
        className: className || classId,
        status: status || ''
      });
    }

    items.sort(function(a, b) {
      return String(a.className || '').localeCompare(String(b.className || ''), 'ko');
    });

    return items;
  } catch (error) {
    throw new Error('반 목록 조회 중 오류가 발생했습니다. ' + error.message);
  }
}

function V2_StudentService_rowToObjectByHeaders_(headers, row) {
  try {
    headers = Array.isArray(headers) ? headers : [];
    row = Array.isArray(row) ? row : [];

    var obj = {};

    for (var i = 0; i < headers.length; i++) {
      obj[V2_StudentService_normalizeHeader_(headers[i])] = row[i];
    }

    return obj;
  } catch (error) {
    return {};
  }
}

function V2_StudentService_pickValue_(rowObject, aliases) {
  try {
    rowObject = rowObject || {};
    aliases = Array.isArray(aliases) ? aliases : [];

    for (var i = 0; i < aliases.length; i++) {
      var key = V2_StudentService_normalizeHeader_(aliases[i]);
      if (Object.prototype.hasOwnProperty.call(rowObject, key)) {
        return rowObject[key];
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_StudentService_normalizeHeader_(value) {
  try {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '')
      .replace(/_/g, '')
      .replace(/-/g, '');
  } catch (error) {
    return '';
  }
}

function V2_createStudentSample() {
  try {
    return V2_createStudent({
      studentName: '테스트학생1',
      classId: 'CLASS_A',
      status: V2_CONFIG.STUDENT_STATUS.ACTIVE,
      parentName: '학부모1',
      parentPhone: '010-1234-5678',
      notes: '학생 등록 기능 테스트 데이터'
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_createStudentSample', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_testGetStudentList() {
  try {
    var result = V2_getStudentList(true);
    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log(error.message);
  }
}
