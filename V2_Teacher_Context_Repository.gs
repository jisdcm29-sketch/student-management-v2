/**
 * V2_Teacher_Context_Repository.gs
 * 교사 컨텍스트 Repository
 * - 시트 접근 전용
 * - 현재 로그인 사용자 기준 교사 정보 조회
 * - 이메일 기준 교사 정보 조회
 * - 기본 교사 컨텍스트 반환
 */

function V2_TeacherContextRepository_getCurrentTeacherContext(criteria) {
  try {
    criteria = criteria || {};

    var email = V2_TCR_toText_(criteria.teacherEmail) || V2_TCR_getCurrentUserEmail_();
    var teacher = V2_TeacherContextRepository_getTeacherByEmail({
      teacherEmail: email
    });

    if (teacher) {
      return {
        found: true,
        teacherId: V2_TCR_toText_(teacher.teacherId),
        teacherName: V2_TCR_toText_(teacher.teacherName),
        teacherEmail: V2_TCR_toText_(teacher.teacherEmail),
        teacherPhone: V2_TCR_toText_(teacher.teacherPhone),
        status: V2_TCR_toText_(teacher.status) || '재직',
        autoDetected: true,
        message: '로그인 사용자 기준 교사 정보를 찾았습니다.'
      };
    }

    return {
      found: false,
      teacherId: '',
      teacherName: V2_TCR_buildTeacherNameFromEmail_(email),
      teacherEmail: email,
      teacherPhone: '',
      status: '',
      autoDetected: !!email,
      message: email
        ? '교사 시트에서 일치하는 교사 정보를 찾지 못했습니다.'
        : '로그인 사용자 이메일을 확인하지 못했습니다.'
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_TeacherContextRepository_getCurrentTeacherContext', error.message, error.stack || '');
    throw error;
  }
}

function V2_TeacherContextRepository_getTeacherByEmail(criteria) {
  try {
    criteria = criteria || {};

    var teacherEmail = V2_TCR_toText_(criteria.teacherEmail).toLowerCase();
    if (!teacherEmail) {
      return null;
    }

    var rows = V2_TCR_getTeacherRows_();

    for (var i = 0; i < rows.length; i++) {
      var item = rows[i];
      var itemEmail = V2_TCR_getTeacherEmail_(item).toLowerCase();

      if (!itemEmail) {
        continue;
      }

      if (itemEmail === teacherEmail) {
        return {
          teacherId: V2_TCR_getTeacherId_(item),
          teacherName: V2_TCR_getTeacherName_(item),
          teacherEmail: V2_TCR_getTeacherEmail_(item),
          teacherPhone: V2_TCR_getTeacherPhone_(item),
          status: V2_TCR_getTeacherStatus_(item)
        };
      }
    }

    return null;
  } catch (error) {
    V2_log_('ERROR', 'V2_TeacherContextRepository_getTeacherByEmail', error.message, error.stack || '');
    throw error;
  }
}

function V2_TCR_getTeacherRows_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName('V2_Teachers');

    if (!sheet) {
      return [];
    }

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var header = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    var values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    var rows = [];

    for (var i = 0; i < values.length; i++) {
      var item = {};
      for (var j = 0; j < header.length; j++) {
        item[V2_TCR_toText_(header[j])] = values[i][j];
      }
      rows.push(item);
    }

    return rows;
  } catch (error) {
    V2_log_('ERROR', 'V2_TCR_getTeacherRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_TCR_getTeacherId_(row) {
  try {
    return V2_TCR_findValueByKeys_(row, ['teacherId', 'teacher_id', '교사ID', '교사Id']);
  } catch (error) {
    return '';
  }
}

function V2_TCR_getTeacherName_(row) {
  try {
    return V2_TCR_findValueByKeys_(row, ['teacherName', 'teacher_name', '교사명', '이름', 'name']);
  } catch (error) {
    return '';
  }
}

function V2_TCR_getTeacherEmail_(row) {
  try {
    return V2_TCR_findValueByKeys_(row, ['teacherEmail', 'teacher_email', 'email', '이메일']);
  } catch (error) {
    return '';
  }
}

function V2_TCR_getTeacherPhone_(row) {
  try {
    return V2_TCR_findValueByKeys_(row, ['teacherPhone', 'teacher_phone', 'phone', '연락처', '전화번호']);
  } catch (error) {
    return '';
  }
}

function V2_TCR_getTeacherStatus_(row) {
  try {
    return V2_TCR_findValueByKeys_(row, ['status', 'teacherStatus', '상태']);
  } catch (error) {
    return '';
  }
}

function V2_TCR_findValueByKeys_(row, keys) {
  try {
    row = row || {};
    keys = keys || [];

    for (var i = 0; i < keys.length; i++) {
      if (Object.prototype.hasOwnProperty.call(row, keys[i])) {
        return V2_TCR_toText_(row[keys[i]]);
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_TCR_getCurrentUserEmail_() {
  try {
    if (typeof V2_getCurrentUserEmail_ === 'function') {
      return V2_TCR_toText_(V2_getCurrentUserEmail_());
    }

    return V2_TCR_toText_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function V2_TCR_buildTeacherNameFromEmail_(email) {
  try {
    var safeEmail = V2_TCR_toText_(email);
    if (!safeEmail) {
      return '';
    }

    var localPart = safeEmail.split('@')[0] || '';
    return localPart.replace(/[._-]+/g, ' ').trim();
  } catch (error) {
    return '';
  }
}

function V2_TCR_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
