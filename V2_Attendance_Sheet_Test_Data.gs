// V2_Attendance_Sheet_Test_Data.gs

function V2_AttendanceSheetTestData_buildRequest_(params) {
  try {
    params = params || {};

    return {
      classId: String(params.classId || '').trim(),
      startDate: String(params.startDate || '').trim(),
      endDate: String(params.endDate || '').trim(),
      includeInactive: params.includeInactive === true
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_buildRequest_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_validateRequest_(request) {
  try {
    request = request || {};

    if (!request.classId) {
      throw new Error('classId가 비어 있다.');
    }

    if (!request.startDate || !request.endDate) {
      throw new Error('startDate 또는 endDate가 비어 있다.');
    }

    if (request.startDate > request.endDate) {
      throw new Error('startDate가 endDate보다 늦다.');
    }

    return request;
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_validateRequest_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_cloneJson_(value) {
  try {
    return JSON.parse(JSON.stringify(value === undefined ? null : value));
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_cloneJson_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_isSuccessResponse_(response) {
  try {
    return !!(response && response.success === true);
  } catch (error) {
    return false;
  }
}

function V2_AttendanceSheetTestData_assertSuccessResponse_(response, label) {
  try {
    if (!V2_AttendanceSheetTestData_isSuccessResponse_(response)) {
      throw new Error((label || '응답') + ' 실패: ' + V2_AttendanceSheetTestData_extractMessage_(response));
    }

    return true;
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_assertSuccessResponse_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_extractMessage_(response) {
  try {
    if (!response) {
      return '응답 객체가 없다.';
    }

    if (response.message) {
      return String(response.message);
    }

    return '메시지가 없다.';
  } catch (error) {
    return '메시지 추출 실패';
  }
}

function V2_AttendanceSheetTestData_getArray_(value) {
  try {
    return Array.isArray(value) ? value : [];
  } catch (error) {
    return [];
  }
}

function V2_AttendanceSheetTestData_buildValidationRequest_(request, fileName, rows) {
  try {
    request = V2_AttendanceSheetTestData_validateRequest_(request);

    return {
      classId: request.classId,
      startDate: request.startDate,
      endDate: request.endDate,
      includeInactive: request.includeInactive === true,
      fileName: String(fileName || 'V2_attendance_test.csv'),
      rows: V2_AttendanceSheetTestData_cloneJson_(V2_AttendanceSheetTestData_getArray_(rows))
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_buildValidationRequest_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_buildApplyRequest_(request, fileName, rows) {
  try {
    request = V2_AttendanceSheetTestData_validateRequest_(request);

    return {
      classId: request.classId,
      startDate: request.startDate,
      endDate: request.endDate,
      includeInactive: request.includeInactive === true,
      fileName: String(fileName || 'V2_attendance_test.csv'),
      rows: V2_AttendanceSheetTestData_cloneJson_(V2_AttendanceSheetTestData_getArray_(rows))
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_buildApplyRequest_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_extractExportRows_(exportResponse) {
  try {
    if (!exportResponse || !exportResponse.data) {
      return [];
    }

    return V2_AttendanceSheetTestData_getArray_(exportResponse.data.rows);
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_extractExportRows_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_extractValidationSummary_(validationResponse) {
  try {
    var data = validationResponse && validationResponse.data ? validationResponse.data : {};
    var summary = data.summary || {};

    return {
      totalRowCount: Number(summary.totalRowCount || 0),
      dataRowCount: Number(summary.dataRowCount || 0),
      validStudentRowCount: Number(summary.validStudentRowCount || 0),
      errorCount: Number(summary.errorCount || 0),
      warningCount: Number(summary.warningCount || 0),
      isValid: data.isValid === true
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_extractValidationSummary_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_extractApplySummary_(applyResponse) {
  try {
    var data = applyResponse && applyResponse.data ? applyResponse.data : {};
    var summary = data.summary || {};

    return {
      appliedCount: Number(summary.appliedCount || 0),
      insertedCount: Number(summary.insertedCount || 0),
      updatedCount: Number(summary.updatedCount || 0)
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_extractApplySummary_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_buildResultRow_(stepName, success, message, detail) {
  try {
    return {
      stepName: String(stepName || ''),
      success: success === true,
      message: String(message || ''),
      detail: detail || {}
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_buildResultRow_ 오류: ' + error.message);
  }
}

function V2_AttendanceSheetTestData_buildFinalResult_(items) {
  try {
    items = Array.isArray(items) ? items : [];

    var hasFailure = items.some(function(item) {
      return !item || item.success !== true;
    });

    return {
      success: !hasFailure,
      executedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      totalStepCount: items.length,
      successStepCount: items.filter(function(item) {
        return item && item.success === true;
      }).length,
      failStepCount: items.filter(function(item) {
        return !item || item.success !== true;
      }).length,
      items: items
    };
  } catch (error) {
    throw new Error('V2_AttendanceSheetTestData_buildFinalResult_ 오류: ' + error.message);
  }
}
