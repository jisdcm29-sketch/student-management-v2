/**
 * V2_Consult_Report_Recent_Service.gs
 * 상담 리포트 최근 조회 학생 Service
 * - UI 호출 전용
 * - 최근 조회 저장
 * - 최근 조회 목록 반환
 * - 반 이름 / 조회 시작일 / 조회 종료일 전달 지원
 * - 최근 조회 요청값 표준화
 * - 최근 조회 최대 5건 정책 고정
 */

function V2_ConsultReportRecentService_saveRecentView(request) {
  try {
    request = request || {};

    var normalizedRequest = V2_CRS_normalizeRecentViewRequest_(request);
    var studentId = normalizedRequest.studentId;
    var studentName = normalizedRequest.studentName;
    var classId = normalizedRequest.classId;
    var className = normalizedRequest.className;
    var startDate = normalizedRequest.startDate;
    var endDate = normalizedRequest.endDate;

    if (!studentId) {
      throw new Error('studentId가 필요합니다.');
    }

    var savedItem = V2_ConsultReportRecentRepository_saveRecentView({
      studentId: studentId,
      studentName: studentName,
      classId: classId,
      className: className,
      startDate: startDate,
      endDate: endDate
    });

    return {
      success: true,
      message: '최근 조회 학생 저장이 완료되었습니다.',
      data: savedItem
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentService_saveRecentView', error.message, error.stack || '');

    return {
      success: false,
      message: error.message,
      data: null
    };
  }
}

function V2_ConsultReportRecentService_getRecentViews(request) {
  try {
    request = request || {};

    var teacherEmail = V2_CRS_toText_(request.teacherEmail);
    var limit = V2_CRS_toRecentViewLimit_(request.limit);

    var items = V2_ConsultReportRecentRepository_getRecentViews({
      teacherEmail: teacherEmail,
      limit: limit
    });

    return {
      success: true,
      message: items.length > 0 ? '최근 조회 학생 목록을 불러왔습니다.' : '최근 조회 학생이 없습니다.',
      data: {
        items: items,
        totalCount: items.length,
        limit: limit
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentService_getRecentViews', error.message, error.stack || '');

    return {
      success: false,
      message: error.message,
      data: {
        items: [],
        totalCount: 0,
        limit: 5
      }
    };
  }
}

function V2_CRS_normalizeRecentViewRequest_(request) {
  try {
    request = request || {};

    return {
      studentId: V2_CRS_toText_(request.studentId),
      studentName: V2_CRS_toText_(request.studentName),
      classId: V2_CRS_toText_(request.classId),
      className: V2_CRS_toText_(request.className),
      startDate: V2_CRS_normalizeDateText_(request.startDate),
      endDate: V2_CRS_normalizeDateText_(request.endDate)
    };
  } catch (error) {
    return {
      studentId: '',
      studentName: '',
      classId: '',
      className: '',
      startDate: '',
      endDate: ''
    };
  }
}

function V2_CRS_normalizeDateText_(value) {
  try {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    var textValue = String(value).trim();

    if (!textValue) {
      return '';
    }

    textValue = textValue.replace(/\./g, '-').replace(/\//g, '-');
    textValue = textValue.replace(/\s+/g, '');

    var directMatch = textValue.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (directMatch) {
      return [
        directMatch[1],
        ('0' + directMatch[2]).slice(-2),
        ('0' + directMatch[3]).slice(-2)
      ].join('-');
    }

    var parsedDate = new Date(textValue);
    if (!isNaN(parsedDate.getTime())) {
      return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    return textValue;
  } catch (error) {
    return V2_CRS_toText_(value);
  }
}

function V2_CRS_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}

function V2_CRS_toLimit_(value) {
  try {
    var numberValue = Number(value);

    if (isNaN(numberValue) || numberValue < 1) {
      return 5;
    }

    if (numberValue > 20) {
      return 20;
    }

    return Math.floor(numberValue);
  } catch (error) {
    return 5;
  }
}

function V2_CRS_toRecentViewLimit_(value) {
  try {
    var numberValue = Number(value);

    if (isNaN(numberValue) || numberValue < 1) {
      return 5;
    }

    if (numberValue > 5) {
      return 5;
    }

    return Math.floor(numberValue);
  } catch (error) {
    return 5;
  }
}
