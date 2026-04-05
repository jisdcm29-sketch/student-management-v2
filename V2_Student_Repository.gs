/**
 * V2_Student_Repository.gs
 * V2 학생 데이터 접근 전용 Repository
 * 시트 직접 접근은 이 파일에서만 처리
 */

/**
 * 학생 시트 반환
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function V2_StudentRepository_getSheet_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.STUDENTS);

    if (!sheet) {
      throw new Error('V2_Students 시트를 찾을 수 없습니다. 먼저 초기 설정을 실행하세요.');
    }

    return sheet;
  } catch (error) {
    throw new Error('V2_StudentRepository_getSheet_ 오류: ' + error.message);
  }
}

/**
 * 학생 헤더 목록 반환
 * @returns {string[]}
 */
function V2_StudentRepository_getHeaders_() {
  try {
    return V2_SHEET_SCHEMAS[V2_CONFIG.SHEETS.STUDENTS].slice();
  } catch (error) {
    throw new Error('V2_StudentRepository_getHeaders_ 오류: ' + error.message);
  }
}

/**
 * 헤더 인덱스 맵 반환
 * @returns {Object}
 */
function V2_StudentRepository_getHeaderMap_() {
  try {
    var headers = V2_StudentRepository_getHeaders_();
    var map = {};

    headers.forEach(function(header, index) {
      map[header] = index;
    });

    return map;
  } catch (error) {
    throw new Error('V2_StudentRepository_getHeaderMap_ 오류: ' + error.message);
  }
}

/**
 * 전체 데이터 행 반환
 * @returns {Array[]}
 */
function V2_StudentRepository_getAllRows_() {
  try {
    var sheet = V2_StudentRepository_getSheet_();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
  } catch (error) {
    throw new Error('V2_StudentRepository_getAllRows_ 오류: ' + error.message);
  }
}

/**
 * 배열 행을 객체로 변환
 * @param {Array} row
 * @returns {Object}
 */
function V2_StudentRepository_rowToObject_(row) {
  try {
    var headers = V2_StudentRepository_getHeaders_();
    var obj = {};

    headers.forEach(function(header, index) {
      obj[header] = row[index];
    });

    return obj;
  } catch (error) {
    throw new Error('V2_StudentRepository_rowToObject_ 오류: ' + error.message);
  }
}

/**
 * 학생 객체를 시트 저장용 배열로 변환
 * @param {Object} student
 * @returns {Array}
 */
function V2_StudentRepository_objectToRow_(student) {
  try {
    var headers = V2_StudentRepository_getHeaders_();

    return headers.map(function(header) {
      return student[header] !== undefined ? student[header] : '';
    });
  } catch (error) {
    throw new Error('V2_StudentRepository_objectToRow_ 오류: ' + error.message);
  }
}

/**
 * 학생 추가
 * @param {Object} student
 * @returns {Object}
 */
function V2_StudentRepository_insertStudent(student) {
  try {
    var sheet = V2_StudentRepository_getSheet_();
    var row = V2_StudentRepository_objectToRow_(student);

    sheet.appendRow(row);

    return student;
  } catch (error) {
    throw new Error('V2_StudentRepository_insertStudent 오류: ' + error.message);
  }
}

/**
 * 학생 전체 조회
 * @returns {Object[]}
 */
function V2_StudentRepository_findAllStudents() {
  try {
    var rows = V2_StudentRepository_getAllRows_();

    return rows.map(function(row) {
      return V2_StudentRepository_rowToObject_(row);
    });
  } catch (error) {
    throw new Error('V2_StudentRepository_findAllStudents 오류: ' + error.message);
  }
}

/**
 * 조건 기반 학생 조회
 * @param {function(Object): boolean} predicate
 * @returns {Object[]}
 */
function V2_StudentRepository_findStudentsBy_(predicate) {
  try {
    var students = V2_StudentRepository_findAllStudents();

    return students.filter(function(student) {
      return predicate(student);
    });
  } catch (error) {
    throw new Error('V2_StudentRepository_findStudentsBy_ 오류: ' + error.message);
  }
}

/**
 * studentId로 학생 조회
 * @param {string} studentId
 * @returns {Object|null}
 */
function V2_StudentRepository_findStudentById(studentId) {
  try {
    var found = V2_StudentRepository_findStudentsBy_(function(student) {
      return String(student.studentId) === String(studentId);
    });

    return found.length > 0 ? found[0] : null;
  } catch (error) {
    throw new Error('V2_StudentRepository_findStudentById 오류: ' + error.message);
  }
}

/**
 * studentId 기준 시트 실제 행 번호 조회
 * @param {string} studentId
 * @returns {number}
 */
function V2_StudentRepository_findRowNumberByStudentId_(studentId) {
  try {
    var rows = V2_StudentRepository_getAllRows_();
    var map = V2_StudentRepository_getHeaderMap_();

    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][map.studentId]) === String(studentId)) {
        return i + 2;
      }
    }

    return -1;
  } catch (error) {
    throw new Error('V2_StudentRepository_findRowNumberByStudentId_ 오류: ' + error.message);
  }
}

/**
 * studentId 기준 전체 학생 객체 업데이트
 * @param {string} studentId
 * @param {Object} student
 * @returns {Object}
 */
function V2_StudentRepository_updateStudentById(studentId, student) {
  try {
    var sheet = V2_StudentRepository_getSheet_();
    var rowNumber = V2_StudentRepository_findRowNumberByStudentId_(studentId);

    if (rowNumber < 0) {
      throw new Error('업데이트 대상 학생을 찾을 수 없습니다. studentId=' + studentId);
    }

    var row = V2_StudentRepository_objectToRow_(student);
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

    return student;
  } catch (error) {
    throw new Error('V2_StudentRepository_updateStudentById 오류: ' + error.message);
  }
}

/**
 * studentId 기준 특정 컬럼만 업데이트
 * @param {string} studentId
 * @param {Object} updates
 * @returns {Object}
 */
function V2_StudentRepository_patchStudentById(studentId, updates) {
  try {
    var current = V2_StudentRepository_findStudentById(studentId);

    if (!current) {
      throw new Error('수정 대상 학생을 찾을 수 없습니다. studentId=' + studentId);
    }

    var updated = Object.assign({}, current, updates);
    return V2_StudentRepository_updateStudentById(studentId, updated);
  } catch (error) {
    throw new Error('V2_StudentRepository_patchStudentById 오류: ' + error.message);
  }
}

/**
 * 이름 중복 검사
 * classId + studentName + 삭제 상태 제외
 * @param {string} studentName
 * @param {string} classId
 * @returns {boolean}
 */
function V2_StudentRepository_existsActiveStudentByNameAndClass(studentName, classId) {
  try {
    var students = V2_StudentRepository_findStudentsBy_(function(student) {
      var sameName = String(student.studentName).trim() === String(studentName).trim();
      var sameClass = String(student.classId).trim() === String(classId).trim();
      var notDeleted = String(student.status).trim() !== V2_CONFIG.STUDENT_STATUS.DELETED;

      return sameName && sameClass && notDeleted;
    });

    return students.length > 0;
  } catch (error) {
    throw new Error('V2_StudentRepository_existsActiveStudentByNameAndClass 오류: ' + error.message);
  }
}
