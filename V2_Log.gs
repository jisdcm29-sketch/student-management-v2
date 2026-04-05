/**
 * V2_Log.gs
 * 공통 로그 진입점
 * - 기존 파일에서 호출 중인 V2_log_ 함수 제공
 * - 신규/기존 기능 모두 공통 사용
 * - 문자열 / 객체 detail 모두 안정적으로 처리
 * - V2_Log_Repository 의 구조화 로그 저장과 연동
 */

function V2_log_(level, source, message, detail) {
  try {
    var logItem = V2_Log_buildLogItem_(level, source, message, detail);
    return V2_LogRepository_saveLog(logItem);
  } catch (error) {
    try {
      Logger.log('[V2_log_ ERROR] ' + error.message);
    } catch (innerError) {}

    return {
      success: false,
      message: error.message || '로그 저장 중 오류가 발생했습니다.',
      data: null
    };
  }
}

function V2_Log_buildLogItem_(level, source, message, detail) {
  try {
    return {
      logId: V2_Log_createLogId_(),
      level: V2_Log_normalizeLevel_(level),
      source: V2_Log_toText_(source),
      message: V2_Log_toText_(message),
      detail: V2_Log_normalizeDetail_(detail),
      userEmail: V2_Log_getCurrentUserEmailSafe_(),
      createdAt: V2_Log_nowText_()
    };
  } catch (error) {
    return {
      logId: V2_Log_createLogId_(),
      level: 'ERROR',
      source: 'V2_Log_buildLogItem_',
      message: error.message || '로그 항목 생성 실패',
      detail: V2_Log_normalizeDetail_({
        errorSource: 'V2_Log_buildLogItem_',
        errorMessage: error.message || '로그 항목 생성 실패'
      }),
      userEmail: '',
      createdAt: V2_Log_nowText_()
    };
  }
}

function V2_Log_createLogId_() {
  try {
    if (typeof V2_createId_ === 'function') {
      return V2_createId_();
    }

    return 'LOG_' + new Date().getTime() + '_' + Math.floor(Math.random() * 100000);
  } catch (error) {
    return 'LOG_' + new Date().getTime();
  }
}

function V2_Log_normalizeLevel_(level) {
  try {
    var text = V2_Log_toText_(level).toUpperCase();

    if (!text) {
      return 'INFO';
    }

    if (text !== 'INFO' && text !== 'WARN' && text !== 'ERROR' && text !== 'DEBUG') {
      return 'INFO';
    }

    return text;
  } catch (error) {
    return 'INFO';
  }
}

function V2_Log_getCurrentUserEmailSafe_() {
  try {
    if (typeof V2_getCurrentUserEmail_ === 'function') {
      return V2_Log_toText_(V2_getCurrentUserEmail_());
    }

    var email = Session.getActiveUser().getEmail();
    return V2_Log_toText_(email);
  } catch (error) {
    return '';
  }
}

function V2_Log_nowText_() {
  try {
    if (typeof V2_nowText_ === 'function') {
      return V2_nowText_();
    }

    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return '';
  }
}

function V2_Log_normalizeDetail_(detail) {
  try {
    if (detail === null || detail === undefined || detail === '') {
      return '';
    }

    if (typeof detail === 'string') {
      return detail.trim();
    }

    if (Object.prototype.toString.call(detail) === '[object Date]' && !isNaN(detail.getTime())) {
      return Utilities.formatDate(detail, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }

    if (typeof detail === 'number' || typeof detail === 'boolean') {
      return String(detail);
    }

    return V2_Log_safeStringify_(detail);
  } catch (error) {
    return V2_Log_toText_(detail);
  }
}

function V2_Log_safeStringify_(value) {
  try {
    var cache = [];
    var result = JSON.stringify(value, function(key, currentValue) {
      if (typeof currentValue === 'object' && currentValue !== null) {
        if (cache.indexOf(currentValue) > -1) {
          return '[Circular]';
        }
        cache.push(currentValue);
      }

      if (Object.prototype.toString.call(currentValue) === '[object Date]' && !isNaN(currentValue.getTime())) {
        return Utilities.formatDate(currentValue, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      }

      if (typeof currentValue === 'function') {
        return '[Function]';
      }

      if (currentValue === undefined) {
        return '[Undefined]';
      }

      return currentValue;
    });

    cache = null;
    return result || '';
  } catch (error) {
    return V2_Log_toText_(value);
  }
}

function V2_Log_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
