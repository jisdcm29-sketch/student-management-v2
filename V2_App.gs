/**
 * V2_App.gs
 * 완전 자동 초기화 + 학생/결석 조회 + 상담 리포트 + 대시보드 + 반별 기간 출석부 + 홈/교사 패널 연결 버전
 */

function onOpen() {
  try {
    V2_buildMenu_();
  } catch (error) {
    V2_log_('ERROR', 'onOpen', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('V2 메뉴 생성 오류: ' + error.message);
  }
}

function V2_buildMenu_() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('V2 시스템')
      .addItem('🚀 완전 초기화 (추천)', 'V2_initializeAll')
      .addItem('초기 설정만', 'V2_runInitialSetup')
      .addItem('교사 동기화', 'V2_syncCurrentTeacher')
      .addSeparator()
      .addItem('기존 홈 대시보드', 'V2_openDashboard')
      .addItem('신규 홈 대시보드', 'V2_openHomeDashboard')
      .addItem('교사 정보 패널', 'V2_openTeacherPanel')
      .addItem('반별 기간 출석부', 'V2_openAttendanceSheet')
      .addSeparator()
      .addItem('학생 등록', 'V2_openStudentForm')
      .addItem('학생 목록', 'V2_openStudentList')
      .addSeparator()
      .addItem('결석 연락 조회', 'V2_openAbsenceList')
      .addItem('학부모 상담 리포트', 'V2_openConsultReport')
      .addToUi();
  } catch (error) {
    V2_log_('ERROR', 'V2_buildMenu_', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('V2 메뉴 구성 오류: ' + error.message);
  }
}

/**
 * 가장 중요한 함수
 * 이 함수 하나로 전체 초기화
 */
function V2_initializeAll() {
  try {
    V2_runInitialSetup();
    V2_syncCurrentTeacher();
    V2_insertSampleData_();

    SpreadsheetApp.getUi().alert(
      '✅ V2 전체 초기화 완료\n\n' +
      '✔ 시트 생성 완료\n' +
      '✔ 교사 등록 완료\n' +
      '✔ 샘플 데이터 생성 완료'
    );
  } catch (error) {
    V2_log_('ERROR', 'V2_initializeAll', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('초기화 실패: ' + error.message);
  }
}

/**
 * 시트 생성 + 헤더 자동 생성
 * 기존 데이터는 보존
 */
function V2_runInitialSetup() {
  try {
    var ss = V2_getSpreadsheet_();

    Object.keys(V2_SHEET_SCHEMAS).forEach(function(name) {
      var schema = V2_SHEET_SCHEMAS[name];
      V2_createSheet_(ss, name, schema);
    });

    SpreadsheetApp.getUi().alert('초기 설정 완료');
  } catch (error) {
    V2_log_('ERROR', 'V2_runInitialSetup', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('초기 설정 실패: ' + error.message);
  }
}

/**
 * 시트 생성 함수
 * 헤더만 보정하고 기존 데이터는 유지
 */
function V2_createSheet_(ss, name, header) {
  try {
    var sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    var lastColumn = sheet.getLastColumn();
    var currentHeader = lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      : [];

    var currentHeaderText = currentHeader.join('||');
    var targetHeaderText = header.join('||');

    if (currentHeaderText !== targetHeaderText) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);

      if (lastColumn > header.length) {
        sheet.getRange(1, header.length + 1, 1, lastColumn - header.length).clearContent();
      }
    }

    sheet.setFrozenRows(1);
  } catch (error) {
    throw new Error('V2_createSheet_ 오류 [' + name + ']: ' + error.message);
  }
}

/**
 * 교사 자동 등록 / 갱신
 */
function V2_syncCurrentTeacher() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.TEACHERS);

    if (!sheet) {
      throw new Error('V2_Teachers 시트를 찾을 수 없습니다.');
    }

    var email = V2_getCurrentUserEmail_();
    var name = V2_getCurrentUserName_();
    var now = V2_nowText_();

    if (!email) {
      throw new Error('현재 사용자 이메일을 확인할 수 없습니다.');
    }

    var lastRow = sheet.getLastRow();
    var values = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 6).getValues() : [];
    var foundRow = -1;

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][2]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
        foundRow = i + 2;
        break;
      }
    }

    if (foundRow > -1) {
      sheet.getRange(foundRow, 2, 1, 5).setValues([[
        name,
        email,
        true,
        values[foundRow - 2][4] || now,
        now
      ]]);

      SpreadsheetApp.getUi().alert('교사 정보가 갱신되었습니다.');
      return;
    }

    sheet.appendRow([
      V2_createId_(),
      name,
      email,
      true,
      now,
      now
    ]);

    SpreadsheetApp.getUi().alert('교사 정보가 등록되었습니다.');
  } catch (error) {
    V2_log_('ERROR', 'V2_syncCurrentTeacher', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('교사 동기화 실패: ' + error.message);
  }
}

/**
 * 샘플 데이터 자동 생성
 * 중복 생성 방지
 */
function V2_insertSampleData_() {
  try {
    var ss = V2_getSpreadsheet_();
    var studentSheet = ss.getSheetByName(V2_CONFIG.SHEETS.STUDENTS);

    if (!studentSheet) {
      throw new Error('V2_Students 시트를 찾을 수 없습니다.');
    }

    var lastRow = studentSheet.getLastRow();
    var existingRows = lastRow >= 2
      ? studentSheet.getRange(2, 1, lastRow - 1, 9).getValues()
      : [];

    var alreadyExists = existingRows.some(function(row) {
      return String(row[1]).trim() === '홍길동' && String(row[2]).trim() === 'CLASS_A';
    });

    if (alreadyExists) {
      return;
    }

    studentSheet.appendRow([
      V2_createId_(),
      '홍길동',
      'CLASS_A',
      V2_CONFIG.STUDENT_STATUS.ACTIVE,
      '부모님',
      '010-0000-0000',
      '샘플 학생',
      V2_nowText_(),
      V2_nowText_()
    ]);
  } catch (error) {
    V2_log_('ERROR', 'V2_insertSampleData_', error.message, error.stack || '');
    throw error;
  }
}

/**
 * 홈 열기
 */
function V2_openHome() {
  try {
    var html = HtmlService.createHtmlOutput('<h2>V2 정상 작동 중</h2>');
    SpreadsheetApp.getUi().showModalDialog(html, 'V2');
  } catch (error) {
    V2_log_('ERROR', 'V2_openHome', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('홈 열기 실패: ' + error.message);
  }
}

/**
 * 기존 홈 대시보드 창 열기
 */
function V2_openDashboard() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Dashboard').evaluate()
      .setWidth(1200)
      .setHeight(720);

    SpreadsheetApp.getUi().showModalDialog(html, '홈 대시보드');
  } catch (error) {
    V2_log_('ERROR', 'V2_openDashboard', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('홈 대시보드 창 열기 실패: ' + error.message);
  }
}

/**
 * 신규 홈 대시보드 창 열기
 */
function V2_openHomeDashboard() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Home').evaluate()
      .setWidth(1200)
      .setHeight(820);

    SpreadsheetApp.getUi().showModalDialog(html, 'V2 홈 대시보드');
  } catch (error) {
    V2_log_('ERROR', 'V2_openHomeDashboard', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('신규 홈 대시보드 창 열기 실패: ' + error.message);
  }
}

/**
 * 교사 정보 패널 창 열기
 */
function V2_openTeacherPanel() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Teacher_Panel').evaluate()
      .setWidth(1100)
      .setHeight(820);

    SpreadsheetApp.getUi().showModalDialog(html, '교사 정보 패널');
  } catch (error) {
    V2_log_('ERROR', 'V2_openTeacherPanel', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('교사 정보 패널 창 열기 실패: ' + error.message);
  }
}

/**
 * 반별 기간 출석부 창 열기
 */
function V2_openAttendanceSheet() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Attendance_Sheet').evaluate()
      .setWidth(1200)
      .setHeight(720);

    SpreadsheetApp.getUi().showModalDialog(html, '반별 기간 출석부');
  } catch (error) {
    V2_log_('ERROR', 'V2_openAttendanceSheet', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('반별 기간 출석부 창 열기 실패: ' + error.message);
  }
}

/**
 * 학생 등록 창 열기
 */
function V2_openStudentForm() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Student_Form').evaluate()
      .setWidth(450)
      .setHeight(520);

    SpreadsheetApp.getUi().showModalDialog(html, '학생 등록');
  } catch (error) {
    V2_log_('ERROR', 'V2_openStudentForm', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('학생 등록 창 열기 실패: ' + error.message);
  }
}

/**
 * 학생 목록 창 열기
 */
function V2_openStudentList() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Student_List').evaluate()
      .setWidth(1100)
      .setHeight(650);

    SpreadsheetApp.getUi().showModalDialog(html, '학생 목록');
  } catch (error) {
    V2_log_('ERROR', 'V2_openStudentList', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('학생 목록 창 열기 실패: ' + error.message);
  }
}

/**
 * 결석 연락 조회 창 열기
 */
function V2_openAbsenceList() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Absence_List').evaluate()
      .setWidth(1100)
      .setHeight(700);

    SpreadsheetApp.getUi().showModalDialog(html, '결석 연락 조회');
  } catch (error) {
    V2_log_('ERROR', 'V2_openAbsenceList', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('결석 연락 조회 창 열기 실패: ' + error.message);
  }
}

/**
 * 학부모 상담 리포트 창 열기
 */
function V2_openConsultReport() {
  try {
    var html = HtmlService.createTemplateFromFile('V2_Consult_Report').evaluate()
      .setWidth(1200)
      .setHeight(720);

    SpreadsheetApp.getUi().showModalDialog(html, '학부모 상담 리포트');
  } catch (error) {
    V2_log_('ERROR', 'V2_openConsultReport', error.message, error.stack || '');
    SpreadsheetApp.getUi().alert('학부모 상담 리포트 창 열기 실패: ' + error.message);
  }
}

/**
 * HTML include helper
 * HTML 파일 내부에서 <?!= V2_include('파일명'); ?> 형태로 사용
 */
function V2_include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    V2_log_('ERROR', 'V2_include', error.message, error.stack || '');
    return '';
  }
}

function V2_App_getStudentSearchPopupHtml() {
  try {
    return HtmlService.createHtmlOutputFromFile('V2_Student_Search_Popup')
      .getContent();
  } catch (error) {
    throw error;
  }
}
