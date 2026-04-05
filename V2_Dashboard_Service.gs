/**
 * V2_Dashboard_Service.gs
 * 홈 대시보드 Service
 * - UI 호출 전용
 * - 대시보드 스냅샷 반환
 */

function V2_DashboardService_getDashboardSnapshot(request) {
  try {
    request = request || {};

    var days = V2_DBS_toDays_(request.days);
    var snapshot = V2_DashboardRepository_getDashboardSnapshot({
      days: days
    });

    return {
      success: true,
      message: snapshot.items && snapshot.items.length > 0
        ? '홈 대시보드 데이터를 불러왔습니다.'
        : '표시할 홈 대시보드 데이터가 없습니다.',
      data: snapshot
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_DashboardService_getDashboardSnapshot', error.message, error.stack || '');

    return {
      success: false,
      message: error.message || '홈 대시보드 데이터 조회 중 오류가 발생했습니다.',
      data: {
        generatedAt: '',
        periodDays: 7,
        totalClassCount: 0,
        totalStudentCount: 0,
        items: []
      }
    };
  }
}

function V2_DBS_toDays_(value) {
  try {
    var numberValue = Number(value);

    if (isNaN(numberValue) || numberValue < 1) {
      return 7;
    }

    if (numberValue > 30) {
      return 30;
    }

    return Math.floor(numberValue);
  } catch (error) {
    return 7;
  }
}
