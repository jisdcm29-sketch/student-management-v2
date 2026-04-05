/**
 * V2_Attendance_Repository.gs
 * 출석 데이터 조회 전용 Repository
 */

function V2_AttendanceRepository_getSheet_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.ATTENDANCE);

    if (!sheet) {
      throw new Error('V2_Attendance 시트를 찾을 수 없습니다.');
    }

    return sheet;
  } catch (error) {
    throw new Error('V2_AttendanceRepository_getSheet_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_getStudentSheet_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.STUDENTS);

    if (!sheet) {
      throw new Error('V2_Students 시트를 찾을 수 없습니다.');
    }

    return sheet;
  } catch (error) {
    throw new Error('V2_AttendanceRepository_getStudentSheet_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_getHeaders_() {
  try {
    return V2_getSheetSchema_(V2_CONFIG.SHEETS.ATTENDANCE);
  } catch (error) {
    throw new Error('V2_AttendanceRepository_getHeaders_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_getStudentHeaders_() {
  try {
    return V2_getSheetSchema_(V2_CONFIG.SHEETS.STUDENTS);
  } catch (error) {
    throw new Error('V2_AttendanceRepository_getStudentHeaders_ 오류: ' + error.message);
  }
}

/**
 * 날짜 값을 yyyy-MM-dd 문자열로 정규화
 * @param {*} value
 * @returns {string}
 */
function V2_AttendanceRepository_normalizeDate_(value) {
  try {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, V2_CONFIG.TIMEZONE, V2_CONFIG.DATE_FORMAT);
    }

    var text = String(value).trim();
    if (!text) {
      return '';
    }

    var directDate = new Date(text);
    if (!isNaN(directDate.getTime())) {
      return Utilities.formatDate(directDate, V2_CONFIG.TIMEZONE, V2_CONFIG.DATE_FORMAT);
    }

    var match = text.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
    if (match) {
      var year = match[1];
      var month = ('0' + match[2]).slice(-2);
      var day = ('0' + match[3]).slice(-2);
      return year + '-' + month + '-' + day;
    }

    return text;
  } catch (error) {
    throw new Error('V2_AttendanceRepository_normalizeDate_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_rowToObject_(headers, row) {
  try {
    var obj = {};

    headers.forEach(function(header, index) {
      var value = row[index];

      if (header === 'date') {
        obj[header] = V2_AttendanceRepository_normalizeDate_(value);
      } else {
        obj[header] = value;
      }
    });

    return obj;
  } catch (error) {
    throw new Error('V2_AttendanceRepository_rowToObject_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_getAllAttendance() {
  try {
    var sheet = V2_AttendanceRepository_getSheet_();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var headers = V2_AttendanceRepository_getHeaders_();
    var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    return rows.map(function(row) {
      return V2_AttendanceRepository_rowToObject_(headers, row);
    });
  } catch (error) {
    throw new Error('V2_AttendanceRepository_getAllAttendance 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_getAllStudents_() {
  try {
    var sheet = V2_AttendanceRepository_getStudentSheet_();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var headers = V2_AttendanceRepository_getStudentHeaders_();
    var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    return rows.map(function(row) {
      return V2_AttendanceRepository_rowToObject_(headers, row);
    });
  } catch (error) {
    throw new Error('V2_AttendanceRepository_getAllStudents_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_findAttendanceByDateAndClass(dateText, classId) {
  try {
    var targetDate = V2_AttendanceRepository_normalizeDate_(dateText);
    var targetClassId = String(classId || '').trim();
    var all = V2_AttendanceRepository_getAllAttendance();

    return all.filter(function(item) {
      return V2_AttendanceRepository_normalizeDate_(item.date) === targetDate &&
             String(item.classId).trim() === targetClassId;
    });
  } catch (error) {
    throw new Error('V2_AttendanceRepository_findAttendanceByDateAndClass 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_findAttendanceByDateRangeAndClass(dateFrom, dateTo, classId) {
  try {
    var normalizedFrom = V2_AttendanceRepository_normalizeDate_(dateFrom);
    var normalizedTo = V2_AttendanceRepository_normalizeDate_(dateTo);
    var targetClassId = String(classId || '').trim();

    var fromTime = new Date(normalizedFrom + 'T00:00:00').getTime();
    var toTime = new Date(normalizedTo + 'T23:59:59').getTime();
    var all = V2_AttendanceRepository_getAllAttendance();

    return all.filter(function(item) {
      var normalizedItemDate = V2_AttendanceRepository_normalizeDate_(item.date);
      var itemTime = new Date(normalizedItemDate + 'T00:00:00').getTime();

      return String(item.classId).trim() === targetClassId &&
             !isNaN(itemTime) &&
             itemTime >= fromTime &&
             itemTime <= toTime;
    });
  } catch (error) {
    throw new Error('V2_AttendanceRepository_findAttendanceByDateRangeAndClass 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_findStudentById_(studentId) {
  try {
    var students = V2_AttendanceRepository_getAllStudents_();
    var found = students.filter(function(student) {
      return String(student.studentId).trim() === String(studentId).trim();
    });

    return found.length > 0 ? found[0] : null;
  } catch (error) {
    throw new Error('V2_AttendanceRepository_findStudentById_ 오류: ' + error.message);
  }
}

function V2_AttendanceRepository_findActiveClasses_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.CLASSES);

    if (!sheet) {
      return [];
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var headers = V2_getSheetSchema_(V2_CONFIG.SHEETS.CLASSES);
    var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    return rows.map(function(row) {
      return V2_AttendanceRepository_rowToObject_(headers, row);
    }).filter(function(item) {
      var value = String(item.isActive).trim().toLowerCase();
      return value !== 'false';
    });
  } catch (error) {
    throw new Error('V2_AttendanceRepository_findActiveClasses_ 오류: ' + error.message);
  }
}
