/**
 * V2_Consult_Report_Print_Repository.gs
 * 상담 리포트 출력용 원본 데이터 조회 Repository
 * 역할:
 * - 시트 접근 전용
 * - studentId 기준 데이터 수집
 * - 기간 필터 적용 가능
 * - 날짜 비교를 "날짜 문자열(yyyy-MM-dd)" 기준으로 안정화
 */

function V2_CRP_getPrintSourceDataByStudentId_(studentId, options) {
  try {
    if (!studentId) {
      throw new Error('studentId가 필요합니다.');
    }

    var normalizedOptions = V2_CRPR_normalizeOptions_(options);

    var student = V2_CRP_getStudentBaseByStudentId_(studentId);
    if (!student) {
      throw new Error('학생 정보를 찾을 수 없습니다. studentId=' + studentId);
    }

    return {
      student: student,
      classInfo: V2_CRP_getClassInfoByClassId_(student.classId),
      teacherMap: V2_CRP_getTeacherMap_(),
      attendanceRows: V2_CRP_getAttendanceRowsByStudentId_(studentId, normalizedOptions),
      scoreRows: V2_CRP_getScoreRowsByStudentId_(studentId, normalizedOptions),
      lessonRows: V2_CRP_getLessonRowsByStudentId_(studentId, normalizedOptions),
      commentRows: V2_CRP_getCommentRowsByStudentId_(studentId, normalizedOptions),
      absenceContactRows: V2_CRP_getAbsenceContactRowsByStudentId_(studentId, normalizedOptions),
      options: normalizedOptions
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getPrintSourceDataByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getStudentBaseByStudentId_(studentId) {
  try {
    var sheetName = V2_CRPR_getSheetName_('STUDENTS', 'V2_Students');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (V2_CRPR_isSameId_(row.studentId || row.id, studentId)) {
        return {
          studentId: row.studentId || row.id || '',
          studentName: row.studentName || row.name || '',
          classId: row.classId || row.classCode || row.class || '',
          status: row.status || '',
          parentName: row.parentName || row.guardianName || '',
          parentPhone: row.parentPhone || row.phone || '',
          memo: row.memo || row.note || '',
          createdAt: row.createdAt || '',
          updatedAt: row.updatedAt || ''
        };
      }
    }

    return null;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getStudentBaseByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getClassInfoByClassId_(classId) {
  try {
    if (!classId) {
      return {
        classId: '',
        className: '',
        teacherId: '',
        teacherName: ''
      };
    }

    var sheetName = V2_CRPR_getSheetName_('CLASSES', 'V2_Classes');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var currentClassId = row.classId || row.id || row.classCode || '';

      if (V2_CRPR_isSameId_(currentClassId, classId)) {
        return {
          classId: currentClassId,
          className: row.className || row.name || currentClassId,
          teacherId: row.teacherId || '',
          teacherName: row.teacherName || ''
        };
      }
    }

    return {
      classId: classId,
      className: classId,
      teacherId: '',
      teacherName: ''
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getClassInfoByClassId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getTeacherMap_() {
  try {
    var sheetName = V2_CRPR_getSheetName_('TEACHERS', 'V2_Teachers');
    var rows = V2_CRPR_getSheetObjects_(sheetName);
    var map = {};

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var teacherId = row.teacherId || row.id || '';
      if (!teacherId) {
        continue;
      }

      map[teacherId] = {
        teacherId: teacherId,
        teacherName: row.teacherName || row.name || '',
        email: row.email || row.teacherEmail || '',
        active: row.isActive,
        createdAt: row.createdAt || '',
        updatedAt: row.updatedAt || ''
      };
    }

    return map;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getTeacherMap_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getAttendanceRowsByStudentId_(studentId, options) {
  try {
    var sheetName = V2_CRPR_getSheetName_('ATTENDANCE', 'V2_Attendance');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    return V2_CRPR_filterAndSortByStudentId_(rows, studentId, options, {
      dateFields: ['attendanceDate', 'date', 'lessonDate', 'createdAt']
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getAttendanceRowsByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getScoreRowsByStudentId_(studentId, options) {
  try {
    var sheetName = V2_CRPR_getSheetName_('SCORES', 'V2_Scores');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    return V2_CRPR_filterAndSortByStudentId_(rows, studentId, options, {
      dateFields: ['scoreDate', 'examDate', 'date', 'createdAt']
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getScoreRowsByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getLessonRowsByStudentId_(studentId, options) {
  try {
    var sheetName = V2_CRPR_getSheetName_('LESSONS', 'V2_Lessons');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    return V2_CRPR_filterAndSortByStudentId_(rows, studentId, options, {
      dateFields: ['lessonDate', 'date', 'createdAt']
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getLessonRowsByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getCommentRowsByStudentId_(studentId, options) {
  try {
    var sheetName = V2_CRPR_getSheetName_('COMMENTS', 'V2_Comments');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    return V2_CRPR_filterAndSortByStudentId_(rows, studentId, options, {
      dateFields: ['commentDate', 'date', 'createdAt']
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getCommentRowsByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRP_getAbsenceContactRowsByStudentId_(studentId, options) {
  try {
    var sheetName = V2_CRPR_getSheetName_('ABSENCE_CONTACTS', 'V2_Absence_Contacts');
    var rows = V2_CRPR_getSheetObjects_(sheetName);

    return V2_CRPR_filterAndSortByStudentId_(rows, studentId, options, {
      dateFields: ['contactDate', 'absenceDate', 'date', 'createdAt']
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_CRP_getAbsenceContactRowsByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRPR_filterAndSortByStudentId_(rows, studentId, options, config) {
  try {
    var result = [];
    var normalizedStudentId = V2_CRPR_toText_(studentId);
    var dateFields = (config && config.dateFields) ? config.dateFields : ['date', 'createdAt'];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var rowStudentId = V2_CRPR_toText_(row.studentId || row.idStudent || row.targetStudentId || '');

      if (rowStudentId !== normalizedStudentId) {
        continue;
      }

      var baseDate = V2_CRPR_pickFirstValue_(row, dateFields);
      if (!V2_CRPR_isDateInRange_(baseDate, options.startDate, options.endDate)) {
        continue;
      }

      row._baseDate = baseDate;
      result.push(row);
    }

    result.sort(function(a, b) {
      var timeA = V2_CRPR_toTime_(a._baseDate);
      var timeB = V2_CRPR_toTime_(b._baseDate);
      return timeA - timeB;
    });

    return result;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRPR_filterAndSortByStudentId_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRPR_getSheetObjects_(sheetName) {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      return [];
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    var objects = [];

    for (var i = 0; i < values.length; i++) {
      var row = {};
      for (var j = 0; j < header.length; j++) {
        row[String(header[j]).trim()] = values[i][j];
      }
      objects.push(row);
    }

    return objects;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRPR_getSheetObjects_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRPR_normalizeOptions_(options) {
  try {
    options = options || {};

    return {
      startDate: V2_CRPR_toComparableDateText_(options.startDate),
      endDate: V2_CRPR_toComparableDateText_(options.endDate),
      includeAttendance: options.includeAttendance !== false,
      includeScores: options.includeScores !== false,
      includeLessons: options.includeLessons !== false,
      includeComments: options.includeComments !== false,
      includeAbsenceContacts: options.includeAbsenceContacts !== false
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_CRPR_normalizeOptions_', V2_CRPR_getErrorMessage_(error), error && error.stack ? error.stack : '');
    throw new Error(V2_CRPR_getErrorMessage_(error));
  }
}

function V2_CRPR_getSheetName_(configKey, fallbackName) {
  try {
    if (
      typeof V2_CONFIG !== 'undefined' &&
      V2_CONFIG &&
      V2_CONFIG.SHEETS &&
      V2_CONFIG.SHEETS[configKey]
    ) {
      return V2_CONFIG.SHEETS[configKey];
    }

    return fallbackName;
  } catch (error) {
    return fallbackName;
  }
}

function V2_CRPR_pickFirstValue_(obj, keys) {
  try {
    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (obj[key] !== undefined && obj[key] !== null && obj[key] !== '') {
        return obj[key];
      }
    }
    return '';
  } catch (error) {
    return '';
  }
}

function V2_CRPR_isDateInRange_(value, startDate, endDate) {
  try {
    var currentDateText = V2_CRPR_toComparableDateText_(value);

    if (!currentDateText) {
      return true;
    }

    if (startDate && currentDateText < startDate) {
      return false;
    }

    if (endDate && currentDateText > endDate) {
      return false;
    }

    return true;
  } catch (error) {
    return true;
  }
}

function V2_CRPR_toComparableDateText_(value) {
  try {
    if (!value) {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    var text = String(value).trim();
    if (!text) {
      return '';
    }

    var match = text.match(/(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})/);
    if (match) {
      return [
        match[1],
        ('0' + match[2]).slice(-2),
        ('0' + match[3]).slice(-2)
      ].join('-');
    }

    var dateObj = V2_CRPR_toDateObject_(value);
    if (!dateObj) {
      return '';
    }

    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    return '';
  }
}

function V2_CRPR_toDateOnlyText_(value) {
  try {
    return V2_CRPR_toComparableDateText_(value);
  } catch (error) {
    return '';
  }
}

function V2_CRPR_toDateObject_(value) {
  try {
    if (!value) {
      return null;
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return new Date(value.getTime());
    }

    var text = String(value).trim();
    if (!text) {
      return null;
    }

    var match = text.match(
      /^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})(?:\s+(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?)?$/
    );

    if (match) {
      var year = Number(match[1]);
      var month = Number(match[2]) - 1;
      var day = Number(match[3]);
      var hour = Number(match[4] || 0);
      var minute = Number(match[5] || 0);
      var second = Number(match[6] || 0);

      return new Date(year, month, day, hour, minute, second);
    }

    var normalizedText = text.replace(/\./g, '-').replace(/\//g, '-');
    var dateObj = new Date(normalizedText);

    if (isNaN(dateObj.getTime())) {
      return null;
    }

    return dateObj;
  } catch (error) {
    return null;
  }
}

function V2_CRPR_toTime_(value) {
  try {
    var dateObj = V2_CRPR_toDateObject_(value);
    return dateObj ? dateObj.getTime() : 0;
  } catch (error) {
    return 0;
  }
}

function V2_CRPR_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  } catch (error) {
    return '';
  }
}

function V2_CRPR_isSameId_(a, b) {
  try {
    return V2_CRPR_toText_(a) === V2_CRPR_toText_(b);
  } catch (error) {
    return false;
  }
}

function V2_CRPR_getErrorMessage_(error) {
  try {
    if (!error) {
      return '알 수 없는 오류가 발생했습니다.';
    }

    if (error.message) {
      return String(error.message);
    }

    return String(error);
  } catch (innerError) {
    return '오류 메시지를 읽을 수 없습니다.';
  }
}
