/**
 * V2_Log_Repository.gs
 * 공통 로그 Repository
 * - V2_Logs 시트 저장 전용
 * - 시트 없으면 안전하게 자동 생성
 * - 구조화 detailJson / detailType / detailPreview 저장 지원
 */

function V2_LogRepository_saveLog(logItem) {
  try {
    logItem = logItem || {};

    var ss = V2_LogRepository_getSpreadsheet_();
    var sheet = V2_LogRepository_getOrCreateLogSheet_(ss);
    var normalizedLogItem = V2_LogRepository_normalizeLogItem_(logItem);
    var row = V2_LogRepository_buildRow_(normalizedLogItem);

    sheet.appendRow(row);

    return {
      success: true,
      message: '로그가 저장되었습니다.',
      data: {
        logId: V2_LogRepository_toText_(normalizedLogItem.logId),
        level: V2_LogRepository_toText_(normalizedLogItem.level),
        source: V2_LogRepository_toText_(normalizedLogItem.source),
        createdAt: V2_LogRepository_toText_(normalizedLogItem.createdAt),
        detailType: V2_LogRepository_toText_(normalizedLogItem.detailType)
      }
    };
  } catch (error) {
    try {
      Logger.log('[V2_LogRepository_saveLog ERROR] ' + error.message);
    } catch (innerError) {}

    return {
      success: false,
      message: error.message || '로그 저장 중 오류가 발생했습니다.',
      data: null
    };
  }
}

function V2_LogRepository_getRecentLogs(limit) {
  try {
    var rowLimit = V2_LogRepository_toNumber_(limit);
    if (rowLimit < 1) {
      rowLimit = 50;
    }

    var ss = V2_LogRepository_getSpreadsheet_();
    var sheet = V2_LogRepository_getOrCreateLogSheet_(ss);
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return {
        success: true,
        message: '조회할 로그가 없습니다.',
        data: []
      };
    }

    var startRow = Math.max(2, lastRow - rowLimit + 1);
    var values = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastColumn).getValues();
    var header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var items = [];

    for (var i = values.length - 1; i >= 0; i--) {
      var item = {};
      for (var j = 0; j < header.length; j++) {
        item[V2_LogRepository_toText_(header[j])] = values[i][j];
      }
      items.push(item);
    }

    return {
      success: true,
      message: '최근 로그를 조회했습니다.',
      data: items
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '로그 조회 중 오류가 발생했습니다.',
      data: []
    };
  }
}

function V2_LogRepository_getOrCreateLogSheet_(ss) {
  try {
    var sheetName = 'V2_Logs';
    var headers = [
      'logId',
      'level',
      'source',
      'message',
      'detail',
      'detailType',
      'detailJson',
      'detailPreview',
      'userEmail',
      'createdAt'
    ];

    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    V2_LogRepository_ensureHeader_(sheet, headers);
    sheet.setFrozenRows(1);

    return sheet;
  } catch (error) {
    throw new Error('V2_LogRepository_getOrCreateLogSheet_ 오류: ' + error.message);
  }
}

function V2_LogRepository_ensureHeader_(sheet, headers) {
  try {
    headers = headers || [];

    var lastColumn = sheet.getLastColumn();
    var currentHeader = lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      : [];

    var currentHeaderText = currentHeader.join('||');
    var targetHeaderText = headers.join('||');

    if (currentHeaderText !== targetHeaderText) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      if (lastColumn > headers.length) {
        sheet.getRange(1, headers.length + 1, 1, lastColumn - headers.length).clearContent();
      }
    }
  } catch (error) {
    throw new Error('V2_LogRepository_ensureHeader_ 오류: ' + error.message);
  }
}

function V2_LogRepository_normalizeLogItem_(logItem) {
  try {
    logItem = logItem || {};

    var detailText = V2_LogRepository_toText_(logItem.detail);
    var detailInfo = V2_LogRepository_buildDetailInfo_(detailText);

    return {
      logId: V2_LogRepository_toText_(logItem.logId),
      level: V2_LogRepository_toText_(logItem.level),
      source: V2_LogRepository_toText_(logItem.source),
      message: V2_LogRepository_toText_(logItem.message),
      detail: detailText,
      detailType: detailInfo.detailType,
      detailJson: detailInfo.detailJson,
      detailPreview: detailInfo.detailPreview,
      userEmail: V2_LogRepository_toText_(logItem.userEmail),
      createdAt: V2_LogRepository_toText_(logItem.createdAt)
    };
  } catch (error) {
    return {
      logId: '',
      level: 'ERROR',
      source: 'V2_LogRepository_normalizeLogItem_',
      message: error.message || '로그 정규화 실패',
      detail: '',
      detailType: 'TEXT',
      detailJson: '',
      detailPreview: '',
      userEmail: '',
      createdAt: V2_LogRepository_nowText_()
    };
  }
}

function V2_LogRepository_buildDetailInfo_(detailText) {
  try {
    detailText = V2_LogRepository_toText_(detailText);

    if (!detailText) {
      return {
        detailType: 'EMPTY',
        detailJson: '',
        detailPreview: ''
      };
    }

    var parsedJson = V2_LogRepository_tryParseJson_(detailText);
    if (parsedJson !== null) {
      var jsonText = JSON.stringify(parsedJson);
      return {
        detailType: V2_LogRepository_detectDetailType_(parsedJson),
        detailJson: jsonText,
        detailPreview: V2_LogRepository_buildDetailPreview_(jsonText)
      };
    }

    return {
      detailType: 'TEXT',
      detailJson: '',
      detailPreview: V2_LogRepository_buildDetailPreview_(detailText)
    };
  } catch (error) {
    return {
      detailType: 'TEXT',
      detailJson: '',
      detailPreview: V2_LogRepository_buildDetailPreview_(detailText)
    };
  }
}

function V2_LogRepository_detectDetailType_(parsedValue) {
  try {
    if (Array.isArray(parsedValue)) {
      return 'JSON_ARRAY';
    }

    if (parsedValue && typeof parsedValue === 'object') {
      return 'JSON_OBJECT';
    }

    return 'JSON_VALUE';
  } catch (error) {
    return 'JSON_VALUE';
  }
}

function V2_LogRepository_tryParseJson_(text) {
  try {
    text = V2_LogRepository_toText_(text);
    if (!text) {
      return null;
    }

    var firstChar = text.charAt(0);
    if (firstChar !== '{' && firstChar !== '[' && firstChar !== '"') {
      return null;
    }

    return JSON.parse(text);
  } catch (error) {
    return null;
  }
}

function V2_LogRepository_buildDetailPreview_(text) {
  try {
    text = V2_LogRepository_toText_(text).replace(/\s+/g, ' ');
    if (text.length <= 300) {
      return text;
    }
    return text.substring(0, 300);
  } catch (error) {
    return '';
  }
}

function V2_LogRepository_buildRow_(logItem) {
  try {
    logItem = logItem || {};

    return [
      V2_LogRepository_toText_(logItem.logId),
      V2_LogRepository_toText_(logItem.level),
      V2_LogRepository_toText_(logItem.source),
      V2_LogRepository_toText_(logItem.message),
      V2_LogRepository_toText_(logItem.detail),
      V2_LogRepository_toText_(logItem.detailType),
      V2_LogRepository_toText_(logItem.detailJson),
      V2_LogRepository_toText_(logItem.detailPreview),
      V2_LogRepository_toText_(logItem.userEmail),
      V2_LogRepository_toText_(logItem.createdAt)
    ];
  } catch (error) {
    return [
      '',
      'ERROR',
      'V2_LogRepository_buildRow_',
      error.message || '로그 행 생성 실패',
      '',
      'TEXT',
      '',
      '',
      '',
      V2_LogRepository_nowText_()
    ];
  }
}

function V2_LogRepository_getSpreadsheet_() {
  try {
    if (typeof V2_getSpreadsheet_ === 'function') {
      return V2_getSpreadsheet_();
    }

    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (error) {
    throw new Error('스프레드시트를 가져올 수 없습니다. ' + error.message);
  }
}

function V2_LogRepository_nowText_() {
  try {
    if (typeof V2_nowText_ === 'function') {
      return V2_nowText_();
    }

    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return '';
  }
}

function V2_LogRepository_toNumber_(value) {
  try {
    var numberValue = Number(value);
    return isNaN(numberValue) ? 0 : numberValue;
  } catch (error) {
    return 0;
  }
}

function V2_LogRepository_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
