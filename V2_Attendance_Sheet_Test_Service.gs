// V2_Attendance_Sheet_Test_Service.gs

function V2_AttendanceSheetTestService_runReadOnlyFlow(params) {
  try {
    var request = V2_AttendanceSheetTestData_buildRequest_(params);
    V2_AttendanceSheetTestData_validateRequest_(request);

    var resultItems = [];

    var sheetResponse = V2_AttendanceSheetService_getClassAttendanceSheet(request);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(sheetResponse, '출석부 조회');
      var sheetData = sheetResponse.data || {};
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '출석부 조회',
          true,
          sheetResponse.message || '조회 성공',
          {
            classId: request.classId,
            startDate: request.startDate,
            endDate: request.endDate,
            totalStudentCount: Number((sheetData.summary && sheetData.summary.totalStudentCount) || 0),
            totalAttendanceCount: Number((sheetData.summary && sheetData.summary.totalAttendanceCount) || 0),
            dateColumnCount: V2_AttendanceSheetTestData_getArray_(sheetData.dateColumns).length
          }
        )
      );
    } catch (stepError1) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('출석부 조회', false, stepError1.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    var exportResponse = V2_AttendanceSheetService_getAttendanceSheetExportData(request);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(exportResponse, '엑셀 출력 데이터 생성');
      var exportData = exportResponse.data || {};
      var exportRows = V2_AttendanceSheetTestData_extractExportRows_(exportResponse);

      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '엑셀 출력 데이터 생성',
          true,
          exportResponse.message || '엑셀 출력 데이터 생성 성공',
          {
            fileName: String(exportData.fileName || ''),
            mimeType: String(exportData.mimeType || ''),
            rowCount: exportRows.length
          }
        )
      );
    } catch (stepError2) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('엑셀 출력 데이터 생성', false, stepError2.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    var validationRequest = V2_AttendanceSheetTestData_buildValidationRequest_(
      request,
      'V2_attendance_readonly_test.csv',
      V2_AttendanceSheetTestData_extractExportRows_(exportResponse)
    );

    var validationResponse = V2_AttendanceSheetService_validateAttendanceSheetUpload(validationRequest);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(validationResponse, '업로드 검증');
      var validationSummary = V2_AttendanceSheetTestData_extractValidationSummary_(validationResponse);

      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '업로드 검증',
          validationSummary.errorCount === 0,
          validationResponse.message || '업로드 검증 완료',
          validationSummary
        )
      );
    } catch (stepError3) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('업로드 검증', false, stepError3.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
  } catch (error) {
    return V2_AttendanceSheetTestData_buildFinalResult_([
      V2_AttendanceSheetTestData_buildResultRow_(
        '읽기 전용 통합 테스트',
        false,
        error.message || '읽기 전용 통합 테스트 실패',
        {}
      )
    ]);
  }
}

function V2_AttendanceSheetTestService_runApplyFlow(params) {
  try {
    params = params || {};

    var request = V2_AttendanceSheetTestData_buildRequest_(params);
    V2_AttendanceSheetTestData_validateRequest_(request);

    var allowWrite = params.allowWrite === true;
    var resultItems = [];

    if (!allowWrite) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '업로드 반영 사전 차단',
          false,
          'allowWrite=true 일 때만 실제 반영 테스트를 실행한다.',
          {
            allowWrite: false
          }
        )
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    var exportResponse = V2_AttendanceSheetService_getAttendanceSheetExportData(request);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(exportResponse, '반영 테스트용 엑셀 출력 데이터 생성');
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '반영 테스트용 엑셀 출력 데이터 생성',
          true,
          exportResponse.message || '테스트용 엑셀 출력 데이터 생성 성공',
          {
            rowCount: V2_AttendanceSheetTestData_extractExportRows_(exportResponse).length
          }
        )
      );
    } catch (stepError1) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('반영 테스트용 엑셀 출력 데이터 생성', false, stepError1.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    var exportRows = V2_AttendanceSheetTestData_extractExportRows_(exportResponse);

    var validationRequest = V2_AttendanceSheetTestData_buildValidationRequest_(
      request,
      'V2_attendance_apply_test.csv',
      exportRows
    );

    var validationResponse = V2_AttendanceSheetService_validateAttendanceSheetUpload(validationRequest);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(validationResponse, '업로드 반영 전 검증');
      var validationSummary = V2_AttendanceSheetTestData_extractValidationSummary_(validationResponse);

      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '업로드 반영 전 검증',
          validationSummary.errorCount === 0,
          validationResponse.message || '업로드 반영 전 검증 완료',
          validationSummary
        )
      );

      if (validationSummary.errorCount > 0) {
        return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
      }
    } catch (stepError2) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('업로드 반영 전 검증', false, stepError2.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    var applyRequest = V2_AttendanceSheetTestData_buildApplyRequest_(
      request,
      'V2_attendance_apply_test.csv',
      exportRows
    );

    var applyResponse = V2_AttendanceSheetService_applyAttendanceSheetUpload(applyRequest);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(applyResponse, '업로드 반영');
      var applySummary = V2_AttendanceSheetTestData_extractApplySummary_(applyResponse);

      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '업로드 반영',
          true,
          applyResponse.message || '업로드 반영 완료',
          applySummary
        )
      );
    } catch (stepError3) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('업로드 반영', false, stepError3.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    var reloadResponse = V2_AttendanceSheetService_getClassAttendanceSheet(request);
    try {
      V2_AttendanceSheetTestData_assertSuccessResponse_(reloadResponse, '업로드 반영 후 재조회');
      var reloadData = reloadResponse.data || {};
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_(
          '업로드 반영 후 재조회',
          true,
          reloadResponse.message || '업로드 반영 후 재조회 성공',
          {
            totalStudentCount: Number((reloadData.summary && reloadData.summary.totalStudentCount) || 0),
            totalAttendanceCount: Number((reloadData.summary && reloadData.summary.totalAttendanceCount) || 0),
            dateColumnCount: V2_AttendanceSheetTestData_getArray_(reloadData.dateColumns).length
          }
        )
      );
    } catch (stepError4) {
      resultItems.push(
        V2_AttendanceSheetTestData_buildResultRow_('업로드 반영 후 재조회', false, stepError4.message, {})
      );
      return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
    }

    return V2_AttendanceSheetTestData_buildFinalResult_(resultItems);
  } catch (error) {
    return V2_AttendanceSheetTestData_buildFinalResult_([
      V2_AttendanceSheetTestData_buildResultRow_(
        '업로드 반영 통합 테스트',
        false,
        error.message || '업로드 반영 통합 테스트 실패',
        {}
      )
    ]);
  }
}

function V2_AttendanceSheetTestService_runAll(params) {
  try {
    params = params || {};

    var readOnlyResult = V2_AttendanceSheetTestService_runReadOnlyFlow(params);
    var applyResult = null;

    if (params.allowWrite === true) {
      applyResult = V2_AttendanceSheetTestService_runApplyFlow(params);
    }

    return {
      success: true,
      message: '반별 기간 출석부 통합 테스트 실행 완료',
      data: {
        readOnlyResult: readOnlyResult,
        applyResult: applyResult
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '반별 기간 출석부 통합 테스트 실행 실패',
      data: null
    };
  }
}

function V2_AttendanceSheetTestService_runSampleReadOnly() {
  try {
    return V2_AttendanceSheetTestService_runReadOnlyFlow({
      classId: 'CLASS_A',
      startDate: '2026-04-01',
      endDate: '2026-04-05',
      includeInactive: false
    });
  } catch (error) {
    return {
      success: false,
      message: error.message || '샘플 읽기 전용 테스트 실패',
      data: null
    };
  }
}

function V2_AttendanceSheetTestService_runSampleApply() {
  try {
    return V2_AttendanceSheetTestService_runApplyFlow({
      classId: 'CLASS_A',
      startDate: '2026-04-01',
      endDate: '2026-04-05',
      includeInactive: false,
      allowWrite: true
    });
  } catch (error) {
    return {
      success: false,
      message: error.message || '샘플 반영 테스트 실패',
      data: null
    };
  }
}
function V2_AttendanceSheetTestService_logSampleReadOnlyResult() {
  try {
    var result = V2_AttendanceSheetTestService_runSampleReadOnly();
    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log(error.message || error);
  }
}

function V2_AttendanceSheetTestService_logSampleApplyResult() {
  try {
    var result = V2_AttendanceSheetTestService_runSampleApply();
    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log(error.message || error);
  }
}
