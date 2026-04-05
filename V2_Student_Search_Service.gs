/**
 * V2_Student_Search_Service.gs
 * 학생 검색 전용 Service
 * - UI 호출 전용
 * - 검색 요청 검증
 * - Repository 결과 반환
 */

function V2_StudentSearchService_searchStudents(request) {
  try {
    request = request || {};

    var keyword = V2_SSS_toText_(request.keyword);
    var status = V2_SSS_toText_(request.status);
    var classId = V2_SSS_toText_(request.classId);
    var limit = V2_SSS_toLimit_(request.limit);

    if (!keyword && !status && !classId) {
      return {
        success: true,
        message: '검색 조건이 없습니다.',
        data: {
          items: [],
          totalCount: 0,
          keyword: '',
          status: status,
          classId: classId,
          limit: limit
        }
      };
    }

    var items = V2_StudentSearchRepository_searchStudents({
      keyword: keyword,
      status: status,
      classId: classId,
      limit: limit
    });

    return {
      success: true,
      message: items.length > 0 ? '학생 검색이 완료되었습니다.' : '검색 결과가 없습니다.',
      data: {
        items: items,
        totalCount: items.length,
        keyword: keyword,
        status: status,
        classId: classId,
        limit: limit
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_StudentSearchService_searchStudents', error.message, error.stack || '');

    return {
      success: false,
      message: error.message,
      data: {
        items: [],
        totalCount: 0,
        keyword: '',
        status: '',
        classId: '',
        limit: 20
      }
    };
  }
}

function V2_StudentSearchService_getStudentByStudentId(studentId) {
  try {
    var normalizedStudentId = V2_SSS_toText_(studentId);

    if (!normalizedStudentId) {
      throw new Error('studentId가 필요합니다.');
    }

    var student = V2_SSR_getStudentByStudentId(normalizedStudentId);

    if (!student) {
      return {
        success: false,
        message: '학생 정보를 찾을 수 없습니다.',
        data: null
      };
    }

    return {
      success: true,
      message: '학생 정보를 찾았습니다.',
      data: student
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_StudentSearchService_getStudentByStudentId', error.message, error.stack || '');

    return {
      success: false,
      message: error.message,
      data: null
    };
  }
}

function V2_SSS_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  } catch (error) {
    return '';
  }
}

function V2_SSS_toLimit_(value) {
  try {
    var numberValue = Number(value);
    if (isNaN(numberValue) || numberValue < 1) {
      return 20;
    }
    if (numberValue > 100) {
      return 100;
    }
    return Math.floor(numberValue);
  } catch (error) {
    return 20;
  }
}

function V2_testStudentSearch_() {
  try {
    var result = V2_StudentSearchService_searchStudents({
      keyword: '홍길',
      limit: 10
    });

    Logger.log(JSON.stringify(result, null, 2));
  } catch (error) {
    Logger.log(error.message);
  }
}
