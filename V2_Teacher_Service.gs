/**
 * V2_Teacher_Service.gs
 * [신규 전체 파일]
 * - 로그인 교사 정보 조회
 * - 교사 자동 매칭
 * - 직접 입력 저장
 * - 추후 권한 확장 고려
 */

function V2_TeacherService_getCurrentTeacherInfo(request) {
  try {
    request = request || {};

    var teacherSheet = V2_TS_getSheetIfExists_('V2_Teachers');
    var teacherRows = V2_TS_getSheetObjects_(teacherSheet);
    var activeEmail = V2_TS_getActiveUserEmail_();
    var matchedTeacher = V2_TS_findTeacherByEmail_(teacherRows, activeEmail);

    return {
      success: true,
      message: matchedTeacher ? '로그인 교사 정보를 불러왔습니다.' : '자동 매칭 가능한 교사 정보가 없습니다.',
      data: {
        matched: !!matchedTeacher,
        teacherId: matchedTeacher ? matchedTeacher.teacherId : '',
        teacherName: matchedTeacher ? matchedTeacher.teacherName : '',
        teacherEmail: activeEmail || '',
        role: matchedTeacher ? matchedTeacher.role : '',
        manualInputAllowed: true
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '교사 정보 조회 중 오류가 발생했습니다.',
      data: {
        matched: false,
        teacherId: '',
        teacherName: '',
        teacherEmail: '',
        role: '',
        manualInputAllowed: true
      }
    };
  }
}

function V2_TeacherService_saveManualTeacherInfo(request) {
  try {
    request = request || {};

    var normalized = V2_TS_normalizeTeacherRequest_(request);
    V2_TS_validateTeacherRequest_(normalized);

    var teacherSheet = V2_TS_getOrCreateTeachersSheet_();
    var teacherRows = V2_TS_getSheetObjects_(teacherSheet);
    var existingTeacher = V2_TS_findTeacherByEmail_(teacherRows, normalized.teacherEmail);

    var savedTeacher = null;

    if (existingTeacher && existingTeacher.rowIndex > 1) {
      V2_TS_updateTeacherRow_(teacherSheet, existingTeacher.rowIndex, normalized);
      savedTeacher = {
        teacherId: existingTeacher.teacherId || normalized.teacherId,
        teacherName: normalized.teacherName,
        teacherEmail: normalized.teacherEmail,
        role: normalized.role
      };
    } else {
      var newTeacherId = normalized.teacherId || V2_TS_generateTeacherId_(teacherRows);
      V2_TS_appendTeacherRow_(teacherSheet, {
        teacherId: newTeacherId,
        teacherName: normalized.teacherName,
        teacherEmail: normalized.teacherEmail,
        role: normalized.role,
        status: '사용'
      });

      savedTeacher = {
        teacherId: newTeacherId,
        teacherName: normalized.teacherName,
        teacherEmail: normalized.teacherEmail,
        role: normalized.role
      };
    }

    return {
      success: true,
      message: '교사 정보를 저장했습니다.',
      data: savedTeacher
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '교사 정보 저장 중 오류가 발생했습니다.',
      data: {
        teacherId: '',
        teacherName: '',
        teacherEmail: '',
        role: ''
      }
    };
  }
}

function V2_TeacherService_getTeacherList() {
  try {
    var teacherSheet = V2_TS_getSheetIfExists_('V2_Teachers');
    var teacherRows = V2_TS_getSheetObjects_(teacherSheet);
    var items = [];

    for (var i = 0; i < teacherRows.length; i++) {
      var row = teacherRows[i] || {};
      items.push({
        teacherId: String(V2_TS_pickValue_(row, ['teacherid', '교사id']) || '').trim(),
        teacherName: String(V2_TS_pickValue_(row, ['teachername', '교사명']) || '').trim(),
        teacherEmail: String(V2_TS_pickValue_(row, ['teacheremail', '이메일']) || '').trim(),
        role: String(V2_TS_pickValue_(row, ['role', '권한']) || '').trim(),
        status: String(V2_TS_pickValue_(row, ['status', '상태']) || '').trim()
      });
    }

    items.sort(function(a, b) {
      return String(a.teacherName || '').localeCompare(String(b.teacherName || ''), 'ko');
    });

    return {
      success: true,
      message: '교사 목록을 불러왔습니다.',
      data: {
        items: items
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '교사 목록 조회 중 오류가 발생했습니다.',
      data: {
        items: []
      }
    };
  }
}

function V2_TS_getActiveUserEmail_() {
  try {
    return String(Session.getActiveUser().getEmail() || '').trim();
  } catch (error) {
    return '';
  }
}

function V2_TS_findTeacherByEmail_(teacherRows, email) {
  try {
    teacherRows = Array.isArray(teacherRows) ? teacherRows : [];
    email = String(email || '').trim().toLowerCase();

    if (!email) {
      return null;
    }

    for (var i = 0; i < teacherRows.length; i++) {
      var row = teacherRows[i] || {};
      var teacherEmail = String(V2_TS_pickValue_(row, ['teacheremail', 'email', '이메일']) || '').trim().toLowerCase();

      if (teacherEmail && teacherEmail === email) {
        return {
          rowIndex: Number(row.__rowIndex || 0),
          teacherId: String(V2_TS_pickValue_(row, ['teacherid', 'id', '교사id']) || '').trim(),
          teacherName: String(V2_TS_pickValue_(row, ['teachername', 'name', '교사명']) || '').trim(),
          teacherEmail: String(V2_TS_pickValue_(row, ['teacheremail', 'email', '이메일']) || '').trim(),
          role: String(V2_TS_pickValue_(row, ['role', '권한']) || '').trim(),
          status: String(V2_TS_pickValue_(row, ['status', '상태']) || '').trim()
        };
      }
    }

    return null;
  } catch (error) {
    return null;
  }
}

function V2_TS_getOrCreateTeachersSheet_() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheetByName('V2_Teachers');

    if (!sheet) {
      sheet = spreadsheet.insertSheet('V2_Teachers');
      sheet.getRange(1, 1, 1, 6).setValues([[
        'teacherId',
        'teacherName',
        'teacherEmail',
        'role',
        'status',
        'updatedAt'
      ]]);
    }

    return sheet;
  } catch (error) {
    throw new Error('V2_Teachers 시트를 준비하지 못했습니다.');
  }
}

function V2_TS_appendTeacherRow_(sheet, data) {
  try {
    var timezone = Session.getScriptTimeZone();
    var nowText = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      data.teacherId || '',
      data.teacherName || '',
      data.teacherEmail || '',
      data.role || '',
      data.status || '사용',
      nowText
    ]);
  } catch (error) {
    throw new Error('교사 정보 추가 중 오류가 발생했습니다.');
  }
}

function V2_TS_updateTeacherRow_(sheet, rowIndex, data) {
  try {
    var timezone = Session.getScriptTimeZone();
    var nowText = Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd HH:mm:ss');

    sheet.getRange(rowIndex, 1, 1, 6).setValues([[
      data.teacherId || '',
      data.teacherName || '',
      data.teacherEmail || '',
      data.role || '',
      data.status || '사용',
      nowText
    ]]);
  } catch (error) {
    throw new Error('교사 정보 수정 중 오류가 발생했습니다.');
  }
}

function V2_TS_generateTeacherId_(teacherRows) {
  try {
    teacherRows = Array.isArray(teacherRows) ? teacherRows : [];

    var maxNumber = 0;

    for (var i = 0; i < teacherRows.length; i++) {
      var row = teacherRows[i] || {};
      var teacherId = String(V2_TS_pickValue_(row, ['teacherid', 'id', '교사id']) || '').trim();
      var match = teacherId.match(/(\d+)$/);

      if (match) {
        maxNumber = Math.max(maxNumber, Number(match[1] || 0));
      }
    }

    return 'T' + ('0000' + (maxNumber + 1)).slice(-4);
  } catch (error) {
    return 'T0001';
  }
}

function V2_TS_normalizeTeacherRequest_(request) {
  request = request || {};

  return {
    teacherId: String(request.teacherId || '').trim(),
    teacherName: String(request.teacherName || '').trim(),
    teacherEmail: String(request.teacherEmail || '').trim(),
    role: String(request.role || '').trim(),
    status: String(request.status || '사용').trim()
  };
}

function V2_TS_validateTeacherRequest_(request) {
  try {
    if (!request.teacherName) {
      throw new Error('교사명을 입력해라.');
    }

    if (!request.teacherEmail) {
      throw new Error('교사 이메일을 입력해라.');
    }

    if (request.teacherEmail.indexOf('@') === -1) {
      throw new Error('교사 이메일 형식이 올바르지 않다.');
    }
  } catch (error) {
    throw error;
  }
}

function V2_TS_getSheetObjects_(sheet) {
  try {
    if (!sheet) {
      return [];
    }

    var values = sheet.getDataRange().getValues();
    if (!values || values.length < 2) {
      return [];
    }

    var headers = values[0];
    var objects = [];

    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var item = {
        __rowIndex: i + 1
      };

      for (var j = 0; j < headers.length; j++) {
        var normalizedHeader = V2_TS_normalizeHeader_(headers[j]);
        if (!normalizedHeader) {
          continue;
        }
        item[normalizedHeader] = row[j];
      }

      objects.push(item);
    }

    return objects;
  } catch (error) {
    return [];
  }
}

function V2_TS_getSheetIfExists_(sheetName) {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  } catch (error) {
    return null;
  }
}

function V2_TS_pickValue_(row, aliases) {
  try {
    row = row || {};
    aliases = Array.isArray(aliases) ? aliases : [];

    for (var i = 0; i < aliases.length; i++) {
      var key = V2_TS_normalizeHeader_(aliases[i]);
      if (Object.prototype.hasOwnProperty.call(row, key)) {
        return row[key];
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_TS_normalizeHeader_(value) {
  try {
    return String(value || '')
      .trim()
      .toLowerCase()
      .replace(/\s+/g, '')
      .replace(/_/g, '')
      .replace(/-/g, '');
  } catch (error) {
    return '';
  }
}
