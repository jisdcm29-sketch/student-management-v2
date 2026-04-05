/**
 * V2_Attendance_Service.gs
 * 결석 학생 조회용 Service
 */

function V2_AttendanceService_validateDate_(dateText) {
  try {
    if (!dateText || !String(dateText).trim()) {
      throw new Error('날짜는 필수입니다.');
    }

    var value = String(dateText).trim();
    var date = new Date(value + 'T00:00:00');

    if (isNaN(date.getTime())) {
      throw new Error('날짜 형식이 올바르지 않습니다. 예: 2026-04-03');
    }

    return value;
  } catch (error) {
    throw new Error('V2_AttendanceService_validateDate_ 오류: ' + error.message);
  }
}

function V2_AttendanceService_validateClassId_(classId) {
  try {
    if (!classId || !String(classId).trim()) {
      throw new Error('classId는 필수입니다.');
    }

    return String(classId).trim();
  } catch (error) {
    throw new Error('V2_AttendanceService_validateClassId_ 오류: ' + error.message);
  }
}

function V2_AttendanceService_getWeekRange_(dateText) {
  try {
    var baseDate = new Date(dateText + 'T00:00:00');
    var day = baseDate.getDay();
    var mondayOffset = day === 0 ? -6 : 1 - day;

    var monday = new Date(baseDate);
    monday.setDate(baseDate.getDate() + mondayOffset);

    var sunday = new Date(monday);
    sunday.setDate(monday.getDate() + 6);

    return {
      dateFrom: Utilities.formatDate(monday, V2_CONFIG.TIMEZONE, V2_CONFIG.DATE_FORMAT),
      dateTo: Utilities.formatDate(sunday, V2_CONFIG.TIMEZONE, V2_CONFIG.DATE_FORMAT)
    };
  } catch (error) {
    throw new Error('V2_AttendanceService_getWeekRange_ 오류: ' + error.message);
  }
}

function V2_AttendanceService_mergeAbsenceWithStudent_(attendanceList) {
  try {
    return attendanceList.map(function(item) {
      var student = V2_AttendanceRepository_findStudentById_(item.studentId);

      return {
        attendanceId: item.attendanceId || '',
        studentId: item.studentId || '',
        studentName: student ? student.studentName || '' : '',
        classId: item.classId || '',
        date: item.date || '',
        status: item.status || '',
        teacherName: item.teacherName || '',
        attendanceMemo: item.memo || '',
        parentName: student ? student.parentName || '' : '',
        parentPhone: student ? student.parentPhone || '' : '',
        studentStatus: student ? student.status || '' : ''
      };
    });
  } catch (error) {
    throw new Error('V2_AttendanceService_mergeAbsenceWithStudent_ 오류: ' + error.message);
  }
}

function V2_AttendanceService_getAbsentStudentsByDateAndClass(dateText, classId) {
  try {
    var validDate = V2_AttendanceService_validateDate_(dateText);
    var validClassId = V2_AttendanceService_validateClassId_(classId);

    var attendanceList = V2_AttendanceRepository_findAttendanceByDateAndClass(validDate, validClassId);
    var absences = attendanceList.filter(function(item) {
      return String(item.status).trim() === V2_CONFIG.ATTENDANCE_STATUS.ABSENT;
    });

    var merged = V2_AttendanceService_mergeAbsenceWithStudent_(absences);
    return V2_createSuccessResponse_(merged);
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceService_getAbsentStudentsByDateAndClass', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_AttendanceService_getAbsentStudentsByWeekAndClass(baseDateText, classId) {
  try {
    var validDate = V2_AttendanceService_validateDate_(baseDateText);
    var validClassId = V2_AttendanceService_validateClassId_(classId);
    var range = V2_AttendanceService_getWeekRange_(validDate);

    var attendanceList = V2_AttendanceRepository_findAttendanceByDateRangeAndClass(
      range.dateFrom,
      range.dateTo,
      validClassId
    );

    var absences = attendanceList.filter(function(item) {
      return String(item.status).trim() === V2_CONFIG.ATTENDANCE_STATUS.ABSENT;
    });

    var merged = V2_AttendanceService_mergeAbsenceWithStudent_(absences);

    return V2_createSuccessResponse_({
      dateFrom: range.dateFrom,
      dateTo: range.dateTo,
      items: merged
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceService_getAbsentStudentsByWeekAndClass', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_AttendanceService_getClassOptions() {
  try {
    var classes = V2_AttendanceRepository_findActiveClasses_().map(function(item) {
      return {
        classId: item.classId || '',
        className: item.className || ''
      };
    });

    return V2_createSuccessResponse_(classes);
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceService_getClassOptions', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

/**
 * 테스트용 샘플 출석 데이터 생성
 * 결석 조회 UI 연결 전 테스트용
 */
function V2_AttendanceService_createSampleAbsence() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.ATTENDANCE);

    if (!sheet) {
      throw new Error('V2_Attendance 시트를 찾을 수 없습니다.');
    }

    var today = V2_todayText_();
    var rows = sheet.getLastRow() >= 2
      ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues()
      : [];

    var alreadyExists = rows.some(function(row) {
      var rowDate = V2_AttendanceRepository_normalizeDate_(row[3]);

      return String(row[1]).trim() !== '' &&
             String(row[2]).trim() === 'CLASS_A' &&
             rowDate === today &&
             String(row[4]).trim() === V2_CONFIG.ATTENDANCE_STATUS.ABSENT;
    });

    if (alreadyExists) {
      return V2_createSuccessResponse_('오늘 결석 샘플 데이터가 이미 존재합니다.');
    }

    var studentSheet = ss.getSheetByName(V2_CONFIG.SHEETS.STUDENTS);
    if (!studentSheet || studentSheet.getLastRow() < 2) {
      throw new Error('학생 데이터가 없습니다.');
    }

    var studentRows = studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 9).getValues();
    var student = null;

    for (var i = 0; i < studentRows.length; i++) {
      if (String(studentRows[i][2]).trim() === 'CLASS_A') {
        student = studentRows[i];
        break;
      }
    }

    if (!student) {
      throw new Error('CLASS_A 학생이 없습니다.');
    }

    sheet.appendRow([
      V2_createId_(),
      student[0],
      student[2],
      today,
      V2_CONFIG.ATTENDANCE_STATUS.ABSENT,
      V2_getCurrentUserName_(),
      '결석 조회 테스트용 샘플',
      V2_nowText_(),
      V2_nowText_()
    ]);

    return V2_createSuccessResponse_('결석 샘플 출석 데이터가 생성되었습니다.');
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceService_createSampleAbsence', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}
