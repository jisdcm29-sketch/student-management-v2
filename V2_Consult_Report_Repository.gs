/**
 * 수정된 전체 파일: V2_Consult_Report_Repository.gs
 * 상담 리포트용 데이터 통합 조회 Repository
 */

function V2_ConsultReportRepository_getSheetByName_(sheetName) {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(sheetName + ' 시트를 찾을 수 없습니다.');
    }

    return sheet;
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getSheetByName_ 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getSchemaBySheetName_(sheetName) {
  try {
    return V2_getSheetSchema_(sheetName);
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getSchemaBySheetName_ 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_rowToObject_(headers, row) {
  try {
    var obj = {};

    headers.forEach(function(header, index) {
      var value = row[index];

      if (value === null || value === undefined) {
        obj[header] = '';
        return;
      }

      if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
        if (header === 'date' || header === 'contactDate' || header === 'examDate' || header === 'lessonDate' || header === 'commentDate') {
          obj[header] = Utilities.formatDate(value, V2_CONFIG.TIMEZONE, V2_CONFIG.DATE_FORMAT);
          return;
        }

        obj[header] = Utilities.formatDate(value, V2_CONFIG.TIMEZONE, V2_CONFIG.DATETIME_FORMAT);
        return;
      }

      obj[header] = value;
    });

    return obj;
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_rowToObject_ 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getAllBySheetName_(sheetName) {
  try {
    var sheet = V2_ConsultReportRepository_getSheetByName_(sheetName);
    var headers = V2_ConsultReportRepository_getSchemaBySheetName_(sheetName);
    var lastRow = sheet.getLastRow();
    var lastColumn = headers.length;

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    return rows.map(function(row) {
      return V2_ConsultReportRepository_rowToObject_(headers, row);
    });
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getAllBySheetName_ 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_sortByDateDesc_(list, fieldName) {
  try {
    list = Array.isArray(list) ? list.slice() : [];
    fieldName = String(fieldName || '').trim();

    list.sort(function(a, b) {
      var aValue = String(a && a[fieldName] ? a[fieldName] : '');
      var bValue = String(b && b[fieldName] ? b[fieldName] : '');

      if (aValue === bValue) {
        return 0;
      }

      return aValue > bValue ? -1 : 1;
    });

    return list;
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_sortByDateDesc_ 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getStudentById(studentId) {
  try {
    studentId = String(studentId || '').trim();

    if (!studentId) {
      throw new Error('studentId는 필수입니다.');
    }

    var students = V2_ConsultReportRepository_getAllBySheetName_(V2_CONFIG.SHEETS.STUDENTS);
    var found = students.filter(function(item) {
      return String(item.studentId || '').trim() === studentId;
    });

    return found.length > 0 ? found[0] : null;
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getStudentById 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getAttendanceByStudentId(studentId) {
  try {
    studentId = String(studentId || '').trim();

    var list = V2_ConsultReportRepository_getAllBySheetName_(V2_CONFIG.SHEETS.ATTENDANCE);

    list = list.filter(function(item) {
      return String(item.studentId || '').trim() === studentId;
    });

    return V2_ConsultReportRepository_sortByDateDesc_(list, 'date');
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getAttendanceByStudentId 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getScoresByStudentId(studentId) {
  try {
    studentId = String(studentId || '').trim();

    var list = V2_ConsultReportRepository_getAllBySheetName_(V2_CONFIG.SHEETS.SCORES);

    list = list.filter(function(item) {
      return String(item.studentId || '').trim() === studentId;
    });

    return V2_ConsultReportRepository_sortByDateDesc_(list, 'examDate');
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getScoresByStudentId 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getLessonsByStudentId(studentId) {
  try {
    studentId = String(studentId || '').trim();

    var list = V2_ConsultReportRepository_getAllBySheetName_(V2_CONFIG.SHEETS.LESSONS);

    list = list.filter(function(item) {
      return String(item.studentId || '').trim() === studentId;
    });

    return V2_ConsultReportRepository_sortByDateDesc_(list, 'lessonDate');
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getLessonsByStudentId 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getCommentsByStudentId(studentId) {
  try {
    studentId = String(studentId || '').trim();

    var list = V2_ConsultReportRepository_getAllBySheetName_(V2_CONFIG.SHEETS.COMMENTS);

    list = list.filter(function(item) {
      return String(item.studentId || '').trim() === studentId;
    });

    return V2_ConsultReportRepository_sortByDateDesc_(list, 'commentDate');
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getCommentsByStudentId 오류: ' + error.message);
  }
}

function V2_ConsultReportRepository_getAbsenceContactsByStudentId(studentId) {
  try {
    studentId = String(studentId || '').trim();

    var list = V2_ConsultReportRepository_getAllBySheetName_(V2_CONFIG.SHEETS.ABSENCE_CONTACTS);

    list = list.filter(function(item) {
      return String(item.studentId || '').trim() === studentId;
    });

    return V2_ConsultReportRepository_sortByDateDesc_(list, 'contactDate');
  } catch (error) {
    throw new Error('V2_ConsultReportRepository_getAbsenceContactsByStudentId 오류: ' + error.message);
  }
}
