/**
 * V2_Config.gs
 * 학생관리 시스템 V2 공통 설정 / 상수 / 시트 스키마 정의
 * 안정성 중심 설계
 */

var V2_CONFIG = Object.freeze({
  APP_NAME: '학생관리 시스템 V2',
  VERSION: '0.2.0',
  TIMEZONE: Session.getScriptTimeZone() || 'Asia/Seoul',

  SHEETS: Object.freeze({
    TEACHERS: 'V2_Teachers',
    CLASSES: 'V2_Classes',
    STUDENTS: 'V2_Students',
    ATTENDANCE: 'V2_Attendance',
    SCORES: 'V2_Scores',
    LESSONS: 'V2_Lessons',
    COMMENTS: 'V2_Comments',
    ABSENCE_CONTACTS: 'V2_Absence_Contacts',
    LOGS: 'V2_Logs'
  }),

  STUDENT_STATUS: Object.freeze({
    ACTIVE: '재학',
    ARCHIVED: '보관',
    DELETE_PENDING: '삭제대기',
    DELETED: '완전삭제'
  }),

  ATTENDANCE_STATUS: Object.freeze({
    PRESENT: '출석',
    ABSENT: '결석',
    LATE: '지각',
    EXCUSED: '인정결석'
  }),

  ABSENCE_CONTACT_STATUS: Object.freeze({
    NOT_CONTACTED: '미연락',
    CONTACTED: '연락완료',
    NO_ANSWER: '부재중',
    CALLBACK_NEEDED: '재연락필요'
  }),

  DATE_FORMAT: 'yyyy-MM-dd',
  DATETIME_FORMAT: 'yyyy-MM-dd HH:mm:ss'
});

var V2_SHEET_SCHEMAS = Object.freeze({
  V2_Teachers: [
    'teacherId',
    'teacherName',
    'teacherEmail',
    'isActive',
    'createdAt',
    'updatedAt'
  ],

  V2_Classes: [
    'classId',
    'className',
    'teacherId',
    'sortOrder',
    'isActive',
    'createdAt',
    'updatedAt'
  ],

  V2_Students: [
    'studentId',
    'studentName',
    'classId',
    'status',
    'parentName',
    'parentPhone',
    'notes',
    'createdAt',
    'updatedAt'
  ],

  V2_Attendance: [
    'attendanceId',
    'studentId',
    'classId',
    'date',
    'status',
    'teacherName',
    'memo',
    'createdAt',
    'updatedAt'
  ],

  V2_Scores: [
    'scoreId',
    'studentId',
    'classId',
    'subject',
    'score',
    'examDate',
    'teacherName',
    'memo',
    'createdAt',
    'updatedAt'
  ],

  V2_Lessons: [
    'lessonId',
    'studentId',
    'classId',
    'lessonDate',
    'topic',
    'content',
    'teacherName',
    'createdAt',
    'updatedAt'
  ],

  V2_Comments: [
    'commentId',
    'studentId',
    'classId',
    'commentDate',
    'commentType',
    'content',
    'teacherName',
    'createdAt',
    'updatedAt'
  ],

  V2_Absence_Contacts: [
    'contactId',
    'attendanceId',
    'studentId',
    'classId',
    'contactDate',
    'contactStatus',
    'parentPhone',
    'memo',
    'teacherName',
    'createdAt',
    'updatedAt'
  ],

  V2_Logs: [
    'logId',
    'level',
    'source',
    'message',
    'details',
    'createdAt'
  ]
});

/**
 * 현재 스프레드시트 반환
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet}
 */
function V2_getSpreadsheet_() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      throw new Error('활성 스프레드시트를 찾을 수 없습니다.');
    }

    return ss;
  } catch (error) {
    throw new Error('V2_getSpreadsheet_ 오류: ' + error.message);
  }
}

/**
 * 현재 시간 문자열 반환
 * @returns {string}
 */
function V2_nowText_() {
  try {
    return Utilities.formatDate(
      new Date(),
      V2_CONFIG.TIMEZONE,
      V2_CONFIG.DATETIME_FORMAT
    );
  } catch (error) {
    throw new Error('V2_nowText_ 오류: ' + error.message);
  }
}

/**
 * 현재 날짜 문자열 반환
 * @returns {string}
 */
function V2_todayText_() {
  try {
    return Utilities.formatDate(
      new Date(),
      V2_CONFIG.TIMEZONE,
      V2_CONFIG.DATE_FORMAT
    );
  } catch (error) {
    throw new Error('V2_todayText_ 오류: ' + error.message);
  }
}

/**
 * UUID 생성
 * @returns {string}
 */
function V2_createId_() {
  try {
    return Utilities.getUuid();
  } catch (error) {
    throw new Error('V2_createId_ 오류: ' + error.message);
  }
}

/**
 * 이메일 기반 현재 사용자명 추정
 * @returns {string}
 */
function V2_getCurrentUserName_() {
  try {
    var email = Session.getActiveUser().getEmail() || '';

    if (!email) {
      return 'Unknown Teacher';
    }

    var idPart = email.split('@')[0] || '';

    if (!idPart) {
      return email;
    }

    return idPart;
  } catch (error) {
    return 'Unknown Teacher';
  }
}

/**
 * 이메일 반환
 * @returns {string}
 */
function V2_getCurrentUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (error) {
    return '';
  }
}

/**
 * 시트 스키마 반환
 * @param {string} sheetName
 * @returns {string[]}
 */
function V2_getSheetSchema_(sheetName) {
  try {
    var schema = V2_SHEET_SCHEMAS[sheetName];

    if (!schema) {
      throw new Error('정의되지 않은 시트 스키마입니다: ' + sheetName);
    }

    return schema.slice();
  } catch (error) {
    throw new Error('V2_getSheetSchema_ 오류: ' + error.message);
  }
}

/**
 * 헤더 맵 생성
 * @param {string[]} headers
 * @returns {Object}
 */
function V2_createHeaderMap_(headers) {
  try {
    var map = {};

    (headers || []).forEach(function(header, index) {
      map[header] = index;
    });

    return map;
  } catch (error) {
    throw new Error('V2_createHeaderMap_ 오류: ' + error.message);
  }
}

/**
 * 로그 기록
 * 실패해도 전체 기능은 죽지 않도록 설계
 * @param {string} level
 * @param {string} source
 * @param {string} message
 * @param {string=} details
 */
function V2_log_(level, source, message, details) {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.LOGS);

    if (!sheet) {
      return;
    }

    sheet.appendRow([
      V2_createId_(),
      level || 'INFO',
      source || '',
      message || '',
      details || '',
      V2_nowText_()
    ]);
  } catch (error) {
    Logger.log('V2_log_ 실패: ' + error.message);
  }
}

/**
 * 공통 성공 응답
 * @param {*} data
 * @returns {{success:boolean,data:*}}
 */
function V2_createSuccessResponse_(data) {
  return {
    success: true,
    data: data
  };
}

/**
 * 공통 실패 응답
 * @param {Error|string} error
 * @returns {{success:boolean,message:string}}
 */
function V2_createErrorResponse_(error) {
  var message = error && error.message ? error.message : String(error);

  return {
    success: false,
    message: message
  };
}
