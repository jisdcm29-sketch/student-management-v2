/**
 * 수정된 전체 파일: V2_Absence_Contact_Repository.gs
 * 결석 연락 기록 데이터 접근 전용 Repository
 */

function V2_AbsenceContactRepository_getSheet_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheet = ss.getSheetByName(V2_CONFIG.SHEETS.ABSENCE_CONTACTS);

    if (!sheet) {
      throw new Error('V2_Absence_Contacts 시트를 찾을 수 없습니다.');
    }

    return sheet;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_getSheet_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_getHeaders_() {
  try {
    return V2_getSheetSchema_(V2_CONFIG.SHEETS.ABSENCE_CONTACTS);
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_getHeaders_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_getHeaderMap_() {
  try {
    return V2_createHeaderMap_(V2_AbsenceContactRepository_getHeaders_());
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_getHeaderMap_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_formatDateValue_(value, format) {
  try {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(
        value,
        V2_CONFIG.TIMEZONE,
        format || V2_CONFIG.DATETIME_FORMAT
      );
    }

    return String(value).trim();
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_formatDateValue_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_normalizeCellValue_(header, value) {
  try {
    if (header === 'contactDate') {
      return V2_AbsenceContactRepository_formatDateValue_(value, V2_CONFIG.DATE_FORMAT);
    }

    if (header === 'createdAt' || header === 'updatedAt') {
      return V2_AbsenceContactRepository_formatDateValue_(value, V2_CONFIG.DATETIME_FORMAT);
    }

    if (value === null || value === undefined) {
      return '';
    }

    return String(value);
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_normalizeCellValue_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_sanitizeEntity_(obj) {
  try {
    var headers = V2_AbsenceContactRepository_getHeaders_();
    var sanitized = {};

    headers.forEach(function(header) {
      sanitized[header] = V2_AbsenceContactRepository_normalizeCellValue_(
        header,
        obj && obj[header] !== undefined ? obj[header] : ''
      );
    });

    return sanitized;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_sanitizeEntity_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_rowToObject_(row) {
  try {
    var headers = V2_AbsenceContactRepository_getHeaders_();
    var obj = {};

    headers.forEach(function(header, index) {
      obj[header] = V2_AbsenceContactRepository_normalizeCellValue_(header, row[index]);
    });

    return obj;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_rowToObject_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_objectToRow_(obj) {
  try {
    var headers = V2_AbsenceContactRepository_getHeaders_();
    var safeObj = V2_AbsenceContactRepository_sanitizeEntity_(obj || {});

    return headers.map(function(header) {
      return safeObj[header] !== undefined ? safeObj[header] : '';
    });
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_objectToRow_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_getAll() {
  try {
    var sheet = V2_AbsenceContactRepository_getSheet_();
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return [];
    }

    var rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    return rows.map(function(row) {
      return V2_AbsenceContactRepository_rowToObject_(row);
    });
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_getAll 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_findByAttendanceId(attendanceId) {
  try {
    attendanceId = String(attendanceId || '').trim();

    if (!attendanceId) {
      return null;
    }

    var all = V2_AbsenceContactRepository_getAll();
    var found = all.filter(function(item) {
      return String(item.attendanceId).trim() === attendanceId;
    });

    return found.length > 0 ? found[0] : null;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_findByAttendanceId 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_findByAttendanceIds(attendanceIds) {
  try {
    attendanceIds = Array.isArray(attendanceIds) ? attendanceIds : [];

    var targetMap = {};
    attendanceIds.forEach(function(attendanceId) {
      var key = String(attendanceId || '').trim();
      if (key) {
        targetMap[key] = true;
      }
    });

    var keys = Object.keys(targetMap);
    if (keys.length === 0) {
      return [];
    }

    var all = V2_AbsenceContactRepository_getAll();

    return all.filter(function(item) {
      return !!targetMap[String(item.attendanceId || '').trim()];
    });
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_findByAttendanceIds 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_findMapByAttendanceIds(attendanceIds) {
  try {
    var records = V2_AbsenceContactRepository_findByAttendanceIds(attendanceIds);
    var map = {};

    records.forEach(function(record) {
      var attendanceId = String(record.attendanceId || '').trim();
      if (attendanceId && !map[attendanceId]) {
        map[attendanceId] = record;
      }
    });

    return map;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_findMapByAttendanceIds 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_findRowNumberByAttendanceId_(attendanceId) {
  try {
    attendanceId = String(attendanceId || '').trim();

    if (!attendanceId) {
      return -1;
    }

    var sheet = V2_AbsenceContactRepository_getSheet_();
    var lastRow = sheet.getLastRow();
    var map = V2_AbsenceContactRepository_getHeaderMap_();

    if (lastRow < 2) {
      return -1;
    }

    var values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    for (var i = 0; i < values.length; i++) {
      if (String(values[i][map.attendanceId]).trim() === attendanceId) {
        return i + 2;
      }
    }

    return -1;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_findRowNumberByAttendanceId_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_insert(entity) {
  try {
    var sheet = V2_AbsenceContactRepository_getSheet_();
    var safeEntity = V2_AbsenceContactRepository_sanitizeEntity_(entity || {});
    var row = V2_AbsenceContactRepository_objectToRow_(safeEntity);

    sheet.appendRow(row);

    return safeEntity;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_insert 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_updateByAttendanceId(attendanceId, entity) {
  try {
    var sheet = V2_AbsenceContactRepository_getSheet_();
    var rowNumber = V2_AbsenceContactRepository_findRowNumberByAttendanceId_(attendanceId);

    if (rowNumber < 0) {
      throw new Error('수정 대상 연락 기록을 찾을 수 없습니다. attendanceId=' + attendanceId);
    }

    var safeEntity = V2_AbsenceContactRepository_sanitizeEntity_(entity || {});
    var row = V2_AbsenceContactRepository_objectToRow_(safeEntity);
    sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

    return safeEntity;
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_updateByAttendanceId 오류: ' + error.message);
  }
}

function V2_AbsenceContactRepository_upsertByAttendanceId(entity) {
  try {
    var attendanceId = String(entity.attendanceId || '').trim();

    if (!attendanceId) {
      throw new Error('attendanceId는 필수입니다.');
    }

    var existing = V2_AbsenceContactRepository_findByAttendanceId(attendanceId);

    if (!existing) {
      return {
        mode: 'insert',
        data: V2_AbsenceContactRepository_insert(entity)
      };
    }

    var updated = Object.assign({}, existing, entity);
    return {
      mode: 'update',
      data: V2_AbsenceContactRepository_updateByAttendanceId(attendanceId, updated)
    };
  } catch (error) {
    throw new Error('V2_AbsenceContactRepository_upsertByAttendanceId 오류: ' + error.message);
  }
}
