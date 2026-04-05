/**
 * V2_Consult_Report_Recent_Delete_Service.gs
 * 상담 리포트 최근 조회 삭제 Service
 * - UI 호출 전용
 * - 현재 교사 기준 최근 조회 1건 삭제
 * - 현재 교사 기준 최근 조회 전체 삭제
 * - 삭제 전/후 최근 조회 정리 상태와 충돌 없이 동작
 */

function V2_ConsultReportRecentDeleteService_deleteOne(request) {
  try {
    request = request || {};

    var recentViewId = V2_CRDS_toText_(request.recentViewId);
    var studentId = V2_CRDS_toText_(request.studentId);
    var startDate = V2_CRDS_normalizeDateText_(request.startDate);
    var endDate = V2_CRDS_normalizeDateText_(request.endDate);

    if (!recentViewId && !studentId) {
      throw new Error('recentViewId 또는 studentId가 필요합니다.');
    }

    var result = V2_ConsultReportRecentDeleteRepository_deleteOne({
      recentViewId: recentViewId,
      studentId: studentId,
      startDate: startDate,
      endDate: endDate
    });

    if (!result.deleted) {
      return {
        success: false,
        message: '삭제할 최근 조회 항목을 찾지 못했습니다.',
        data: {
          deletedCount: 0,
          item: null
        }
      };
    }

    return {
      success: true,
      message: '최근 조회 1건이 삭제되었습니다.',
      data: {
        deletedCount: result.deletedCount,
        item: result.item
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentDeleteService_deleteOne', error.message, error.stack || '');

    return {
      success: false,
      message: error.message,
      data: {
        deletedCount: 0,
        item: null
      }
    };
  }
}

function V2_ConsultReportRecentDeleteService_deleteAll() {
  try {
    var result = V2_ConsultReportRecentDeleteRepository_deleteAllForCurrentTeacher();

    if (!result.deleted) {
      return {
        success: true,
        message: '삭제할 최근 조회 항목이 없습니다.',
        data: {
          deletedCount: 0,
          items: []
        }
      };
    }

    return {
      success: true,
      message: '최근 조회 전체 삭제가 완료되었습니다.',
      data: {
        deletedCount: result.deletedCount,
        items: result.items
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentDeleteService_deleteAll', error.message, error.stack || '');

    return {
      success: false,
      message: error.message,
      data: {
        deletedCount: 0,
        items: []
      }
    };
  }
}

function V2_CRDS_normalizeDateText_(value) {
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
    return V2_CRDS_toText_(value);
  }
}

function V2_CRDS_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
