/**
 * V2_Consult_Report_Recent_Delete_Repository.gs
 * 상담 리포트 최근 조회 삭제 Repository
 * - 시트 접근 전용
 * - 현재 교사 기준 최근 조회 1건 삭제
 * - 현재 교사 기준 최근 조회 전체 삭제
 * - 기존 최근 조회 시트 구조 재사용
 * - 삭제 전 정규화 / 삭제 후 재정렬 / 헤더 유지 안정화
 */

function V2_ConsultReportRecentDeleteRepository_deleteOne(criteria) {
  try {
    criteria = criteria || {};

    var recentViewId = V2_CRD_toText_(criteria.recentViewId);
    var studentId = V2_CRD_toText_(criteria.studentId);
    var startDate = V2_CRD_normalizeDateText_(criteria.startDate);
    var endDate = V2_CRD_normalizeDateText_(criteria.endDate);
    var teacherEmail = V2_CRD_getCurrentTeacherEmail_();

    if (!recentViewId && !studentId) {
      throw new Error('recentViewId 또는 studentId가 필요합니다.');
    }

    var sheet = V2_CRD_getRecentViewSheet_();
    var normalizedRows = V2_CRD_getNormalizedRows_(sheet, teacherEmail);
    var deletedItem = null;
    var remainingRows = [];

    for (var i = 0; i < normalizedRows.length; i++) {
      var row = V2_CRD_buildRecentViewRow_(normalizedRows[i]);

      if (deletedItem) {
        remainingRows.push(row);
        continue;
      }

      if (!V2_CRD_isSameText_(row.teacherEmail, teacherEmail)) {
        remainingRows.push(row);
        continue;
      }

      if (V2_CRD_isDeleteTargetRow_(row, {
        recentViewId: recentViewId,
        studentId: studentId,
        startDate: startDate,
        endDate: endDate
      })) {
        deletedItem = row;
        continue;
      }

      remainingRows.push(row);
    }

    if (!deletedItem) {
      if (V2_CRD_shouldRewriteNormalizedRows_(sheet, normalizedRows)) {
        V2_CRD_writeAllRows_(sheet, normalizedRows);
      }

      return {
        deleted: false,
        deletedCount: 0,
        item: null
      };
    }

    V2_CRD_writeAllRows_(sheet, remainingRows);

    return {
      deleted: true,
      deletedCount: 1,
      item: deletedItem
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentDeleteRepository_deleteOne', error.message, error.stack || '');
    throw error;
  }
}

function V2_ConsultReportRecentDeleteRepository_deleteAllForCurrentTeacher() {
  try {
    var teacherEmail = V2_CRD_getCurrentTeacherEmail_();
    var sheet = V2_CRD_getRecentViewSheet_();
    var normalizedRows = V2_CRD_getNormalizedRows_(sheet, teacherEmail);
    var remainingRows = [];
    var deletedItems = [];

    for (var i = 0; i < normalizedRows.length; i++) {
      var row = V2_CRD_buildRecentViewRow_(normalizedRows[i]);

      if (V2_CRD_isSameText_(row.teacherEmail, teacherEmail)) {
        deletedItems.push(row);
        continue;
      }

      remainingRows.push(row);
    }

    if (deletedItems.length === 0) {
      if (V2_CRD_shouldRewriteNormalizedRows_(sheet, normalizedRows)) {
        V2_CRD_writeAllRows_(sheet, normalizedRows);
      }

      return {
        deleted: false,
        deletedCount: 0,
        items: []
      };
    }

    V2_CRD_writeAllRows_(sheet, remainingRows);

    return {
      deleted: true,
      deletedCount: deletedItems.length,
      items: deletedItems
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentDeleteRepository_deleteAllForCurrentTeacher', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRD_getRecentViewSheet_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheetName = 'V2_Consult_Report_Recent_Views';
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(sheetName + ' 시트를 찾을 수 없습니다.');
    }

    V2_CRD_ensureRecentViewHeader_(sheet);
    return sheet;
  } catch (error) {
    throw new Error('최근 조회 시트 접근 중 오류가 발생했습니다. ' + error.message);
  }
}

function V2_CRD_ensureRecentViewHeader_(sheet) {
  try {
    if (!sheet) {
      throw new Error('시트가 없습니다.');
    }

    var header = V2_CRD_getHeader_();
    var lastColumn = sheet.getLastColumn();
    var currentHeader = lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      : [];

    if (currentHeader.join('||') !== header.join('||')) {
      sheet.getRange(1, 1, 1, header.length).setValues([header]);

      if (lastColumn > header.length) {
        sheet.getRange(1, header.length + 1, 1, lastColumn - header.length).clearContent();
      }
    }

    sheet.setFrozenRows(1);
  } catch (error) {
    throw new Error('최근 조회 시트 헤더 확인 중 오류가 발생했습니다. ' + error.message);
  }
}

function V2_CRD_getHeader_() {
  return [
    'recentViewId',
    'teacherEmail',
    'teacherName',
    'studentId',
    'studentName',
    'classId',
    'className',
    'startDate',
    'endDate',
    'viewedAt'
  ];
}

function V2_CRD_getSheetObjects_(sheet) {
  try {
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
        item[V2_CRD_toText_(header[j])] = values[i][j];
      }
      rows.push(V2_CRD_buildRecentViewRow_(item));
    }

    return rows;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRD_getSheetObjects_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRD_getNormalizedRows_(sheet, teacherEmail) {
  try {
    var rows = V2_CRD_getSheetObjects_(sheet);
    return V2_CRD_normalizeRows_(rows, teacherEmail);
  } catch (error) {
    V2_log_('ERROR', 'V2_CRD_getNormalizedRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRD_normalizeRows_(rows, teacherEmail) {
  try {
    rows = rows || [];
    var normalizedTeacherEmail = V2_CRD_toText_(teacherEmail);
    var keepLimit = 5;
    var teacherRows = [];
    var otherRows = [];
    var uniqueMap = {};
    var normalizedTeacherRows = [];
    var mergedRows = [];

    for (var i = 0; i < rows.length; i++) {
      var row = V2_CRD_buildRecentViewRow_(rows[i]);

      if (normalizedTeacherEmail && V2_CRD_isSameText_(row.teacherEmail, normalizedTeacherEmail)) {
        teacherRows.push(row);
      } else {
        otherRows.push(row);
      }
    }

    teacherRows = V2_CRD_sortRows_(teacherRows);

    for (var j = 0; j < teacherRows.length; j++) {
      var teacherRow = teacherRows[j];
      var uniqueKey = V2_CRD_buildUniqueKey_(teacherRow);

      if (!uniqueKey) {
        continue;
      }

      if (uniqueMap[uniqueKey]) {
        continue;
      }

      uniqueMap[uniqueKey] = true;
      normalizedTeacherRows.push(teacherRow);

      if (normalizedTeacherRows.length >= keepLimit) {
        break;
      }
    }

    mergedRows = otherRows.concat(normalizedTeacherRows);
    mergedRows = V2_CRD_sortRows_(mergedRows);

    return mergedRows;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRD_normalizeRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRD_sortRows_(rows) {
  try {
    rows = rows || [];

    var copiedRows = [];
    for (var i = 0; i < rows.length; i++) {
      copiedRows.push(V2_CRD_buildRecentViewRow_(rows[i]));
    }

    copiedRows.sort(function(a, b) {
      var timeDiff = V2_CRD_toSortTime_(b.viewedAt) - V2_CRD_toSortTime_(a.viewedAt);
      if (timeDiff !== 0) {
        return timeDiff;
      }

      return V2_CRD_toText_(b.recentViewId).localeCompare(V2_CRD_toText_(a.recentViewId));
    });

    return copiedRows;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRD_sortRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRD_shouldRewriteNormalizedRows_(sheet, normalizedRows) {
  try {
    var currentRows = V2_CRD_getSheetObjects_(sheet);

    if (currentRows.length !== normalizedRows.length) {
      return true;
    }

    for (var i = 0; i < currentRows.length; i++) {
      if (V2_CRD_rowSignature_(currentRows[i]) !== V2_CRD_rowSignature_(normalizedRows[i])) {
        return true;
      }
    }

    return false;
  } catch (error) {
    return true;
  }
}

function V2_CRD_writeAllRows_(sheet, rows) {
  try {
    if (!sheet) {
      throw new Error('시트가 없습니다.');
    }

    var header = V2_CRD_getHeader_();
    var normalizedRows = rows || [];

    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);

    if (normalizedRows.length < 1) {
      return;
    }

    var values = [];
    for (var i = 0; i < normalizedRows.length; i++) {
      var row = V2_CRD_buildRecentViewRow_(normalizedRows[i]);
      values.push([
        row.recentViewId,
        row.teacherEmail,
        row.teacherName,
        row.studentId,
        row.studentName,
        row.classId,
        row.className,
        row.startDate,
        row.endDate,
        row.viewedAt
      ]);
    }

    sheet.getRange(2, 1, values.length, header.length).setValues(values);
  } catch (error) {
    V2_log_('ERROR', 'V2_CRD_writeAllRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRD_isDeleteTargetRow_(row, criteria) {
  try {
    row = V2_CRD_buildRecentViewRow_(row);
    criteria = criteria || {};

    var recentViewId = V2_CRD_toText_(criteria.recentViewId);
    var studentId = V2_CRD_toText_(criteria.studentId);
    var startDate = V2_CRD_normalizeDateText_(criteria.startDate);
    var endDate = V2_CRD_normalizeDateText_(criteria.endDate);

    if (recentViewId) {
      return V2_CRD_isSameText_(row.recentViewId, recentViewId);
    }

    if (!studentId) {
      return false;
    }

    if (!V2_CRD_isSameText_(row.studentId, studentId)) {
      return false;
    }

    if (startDate || endDate) {
      return V2_CRD_isSameText_(row.startDate, startDate) &&
        V2_CRD_isSameText_(row.endDate, endDate);
    }

    return true;
  } catch (error) {
    return false;
  }
}

function V2_CRD_rowSignature_(row) {
  try {
    row = V2_CRD_buildRecentViewRow_(row);

    return [
      row.recentViewId,
      row.teacherEmail,
      row.teacherName,
      row.studentId,
      row.studentName,
      row.classId,
      row.className,
      row.startDate,
      row.endDate,
      row.viewedAt
    ].join('||');
  } catch (error) {
    return '';
  }
}

function V2_CRD_buildRecentViewRow_(row) {
  try {
    row = row || {};

    return {
      recentViewId: V2_CRD_toText_(row.recentViewId),
      teacherEmail: V2_CRD_toText_(row.teacherEmail),
      teacherName: V2_CRD_toText_(row.teacherName),
      studentId: V2_CRD_toText_(row.studentId),
      studentName: V2_CRD_toText_(row.studentName),
      classId: V2_CRD_toText_(row.classId),
      className: V2_CRD_toText_(row.className),
      startDate: V2_CRD_normalizeDateText_(row.startDate),
      endDate: V2_CRD_normalizeDateText_(row.endDate),
      viewedAt: V2_CRD_normalizeDateTimeText_(row.viewedAt)
    };
  } catch (error) {
    return {
      recentViewId: '',
      teacherEmail: '',
      teacherName: '',
      studentId: '',
      studentName: '',
      classId: '',
      className: '',
      startDate: '',
      endDate: '',
      viewedAt: ''
    };
  }
}

function V2_CRD_buildUniqueKey_(row) {
  try {
    row = V2_CRD_buildRecentViewRow_(row);

    return [
      V2_CRD_toText_(row.teacherEmail).toLowerCase(),
      V2_CRD_toText_(row.studentId).toLowerCase(),
      V2_CRD_normalizeDateText_(row.startDate),
      V2_CRD_normalizeDateText_(row.endDate)
    ].join('||');
  } catch (error) {
    return '';
  }
}

function V2_CRD_getCurrentTeacherEmail_() {
  try {
    if (typeof V2_getCurrentUserEmail_ === 'function') {
      return V2_CRD_toText_(V2_getCurrentUserEmail_());
    }

    return V2_CRD_toText_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function V2_CRD_normalizeDateText_(value) {
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
    return V2_CRD_toText_(value);
  }
}

function V2_CRD_normalizeDateTimeText_(value) {
  try {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }

    var textValue = String(value).trim();

    if (!textValue) {
      return '';
    }

    var normalizedText = textValue.replace(/\./g, '-').replace(/\//g, '-');
    var parsedDate = new Date(normalizedText);

    if (!isNaN(parsedDate.getTime())) {
      return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }

    return textValue;
  } catch (error) {
    return V2_CRD_toText_(value);
  }
}

function V2_CRD_toSortTime_(value) {
  try {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return value.getTime();
    }

    var normalizedText = V2_CRD_normalizeDateTimeText_(value);
    var parsedDate = new Date(normalizedText);

    if (!isNaN(parsedDate.getTime())) {
      return parsedDate.getTime();
    }

    return 0;
  } catch (error) {
    return 0;
  }
}

function V2_CRD_isSameText_(a, b) {
  try {
    return V2_CRD_toText_(a).toLowerCase() === V2_CRD_toText_(b).toLowerCase();
  } catch (error) {
    return false;
  }
}

function V2_CRD_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
