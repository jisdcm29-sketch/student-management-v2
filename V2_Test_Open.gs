/**
 * V2_Test_Open.gs
 * 테스트 오픈용
 */

function testOpenConsultReport() {
  try {
    var html = HtmlService
      .createTemplateFromFile('V2_Consult_Report')
      .evaluate()
      .setWidth(1200)
      .setHeight(900);

    SpreadsheetApp.getUi().showModalDialog(html, '학부모 상담 리포트');
  } catch (error) {
    SpreadsheetApp.getUi().alert('상담 리포트 열기 실패: ' + error.message);
    throw error;
  }
}

function testOpenHome() {
  try {
    var html = HtmlService
      .createTemplateFromFile('V2_Home')
      .evaluate()
      .setWidth(1200)
      .setHeight(820);

    SpreadsheetApp.getUi().showModalDialog(html, 'V2 홈 대시보드');
  } catch (error) {
    SpreadsheetApp.getUi().alert('홈 대시보드 열기 실패: ' + error.message);
    throw error;
  }
}

function testOpenTeacherPanel() {
  try {
    var html = HtmlService
      .createTemplateFromFile('V2_Teacher_Panel')
      .evaluate()
      .setWidth(1100)
      .setHeight(820);

    SpreadsheetApp.getUi().showModalDialog(html, '교사 정보 패널');
  } catch (error) {
    SpreadsheetApp.getUi().alert('교사 정보 패널 열기 실패: ' + error.message);
    throw error;
  }
}

function testPing() {
  Logger.log('ping');
}
