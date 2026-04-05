/**
 * V2_Teacher_Context_Service.gs
 * 교사 컨텍스트 Service
 * - UI 호출 전용
 * - 현재 로그인 사용자 기준 교사 정보 반환
 * - 자동 입력용 기본값 반환
 */

function V2_TeacherContextService_getCurrentTeacherContext(request) {
  try {
    request = request || {};

    var repositoryResult = V2_TeacherContextRepository_getCurrentTeacherContext({
      teacherEmail: V2_TCS_toText_(request.teacherEmail)
    });

    return {
      success: true,
      message: repositoryResult.message || '현재 교사 정보를 불러왔습니다.',
      data: {
        found: !!repositoryResult.found,
        autoDetected: !!repositoryResult.autoDetected,
        teacherId: V2_TCS_toText_(repositoryResult.teacherId),
        teacherName: V2_TCS_toText_(repositoryResult.teacherName),
        teacherEmail: V2_TCS_toText_(repositoryResult.teacherEmail),
        teacherPhone: V2_TCS_toText_(repositoryResult.teacherPhone),
        status: V2_TCS_toText_(repositoryResult.status),
        displayTeacherName: V2_TCS_buildDisplayTeacherName_(repositoryResult),
        displayLabel: V2_TCS_buildDisplayLabel_(repositoryResult)
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_TeacherContextService_getCurrentTeacherContext', error.message, error.stack || '');

    return {
      success: false,
      message: error.message || '현재 교사 정보를 불러오는 중 오류가 발생했습니다.',
      data: {
        found: false,
        autoDetected: false,
        teacherId: '',
        teacherName: '',
        teacherEmail: '',
        teacherPhone: '',
        status: '',
        displayTeacherName: '',
        displayLabel: ''
      }
    };
  }
}

function V2_TeacherContextService_getTeacherAutoFillValue(request) {
  try {
    request = request || {};

    var response = V2_TeacherContextService_getCurrentTeacherContext({
      teacherEmail: V2_TCS_toText_(request.teacherEmail)
    });

    if (!response.success) {
      return response;
    }

    return {
      success: true,
      message: response.message,
      data: {
        teacherName: V2_TCS_toText_(response.data.displayTeacherName),
        teacherEmail: V2_TCS_toText_(response.data.teacherEmail),
        teacherId: V2_TCS_toText_(response.data.teacherId),
        found: !!response.data.found,
        autoDetected: !!response.data.autoDetected
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_TeacherContextService_getTeacherAutoFillValue', error.message, error.stack || '');

    return {
      success: false,
      message: error.message || '교사 자동 입력값 생성 중 오류가 발생했습니다.',
      data: {
        teacherName: '',
        teacherEmail: '',
        teacherId: '',
        found: false,
        autoDetected: false
      }
    };
  }
}

function V2_TCS_buildDisplayTeacherName_(teacherContext) {
  try {
    teacherContext = teacherContext || {};

    if (V2_TCS_toText_(teacherContext.teacherName)) {
      return V2_TCS_toText_(teacherContext.teacherName);
    }

    if (V2_TCS_toText_(teacherContext.teacherEmail)) {
      return V2_TCS_toText_(teacherContext.teacherEmail).split('@')[0];
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_TCS_buildDisplayLabel_(teacherContext) {
  try {
    teacherContext = teacherContext || {};

    var lines = [];
    lines.push('교사 정보');

    if (teacherContext.found) {
      lines.push('상태: 교사 시트 연결 완료');
    } else if (teacherContext.autoDetected) {
      lines.push('상태: 로그인 정보만 확인');
    } else {
      lines.push('상태: 자동 확인 불가');
    }

    lines.push('교사명: ' + V2_TCS_buildDisplayTeacherName_(teacherContext));
    lines.push('이메일: ' + V2_TCS_toText_(teacherContext.teacherEmail));
    lines.push('교사 ID: ' + V2_TCS_toText_(teacherContext.teacherId));
    lines.push('연락처: ' + V2_TCS_toText_(teacherContext.teacherPhone));
    lines.push('재직 상태: ' + V2_TCS_toText_(teacherContext.status));

    return lines.join('\n');
  } catch (error) {
    return '';
  }
}

function V2_TCS_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
