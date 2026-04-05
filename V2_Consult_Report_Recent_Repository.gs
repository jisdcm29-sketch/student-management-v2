/**
 * V2_Consult_Report_Recent_Repository.gs
 * 상담 리포트 최근 조회 학생 Repository
 * - 시트 접근 전용
 * - 최근 조회 기록 저장
 * - 교사별 최근 조회 목록 반환
 * - V2_Consult_Report_Recent_Views 시트 자동 생성
 * - 반 이름 / 조회 시작일 / 조회 종료일 저장 지원
 * - 최근 조회 중복 정책 고정
 * - 교사별 최근 조회 최대 보관 개수 유지
 * - 저장 시 정렬 / 중복 제거 / 최대 5건 유지 통합 처리
 */

function V2_ConsultReportRecentRepository_saveRecentView(data) {
  try {
    data = data || {};

    var normalizedData = V2_CRR_normalizeRecentViewData_(data);
    var studentId = normalizedData.studentId;
    var studentName = normalizedData.studentName;
    var classId = normalizedData.classId;
    var className = normalizedData.className;
    var startDate = normalizedData.startDate;
    var endDate = normalizedData.endDate;
    var teacherEmail = V2_CRR_getCurrentTeacherEmail_();
    var teacherName = V2_CRR_getCurrentTeacherName_();
    var viewedAt = V2_CRR_nowText_();
    var keepLimit = V2_CRR_getRecentViewKeepLimit_();

    if (!studentId) {
      throw new Error('studentId가 필요합니다.');
    }

    var sheet = V2_CRR_getOrCreateSheet_();
    var allRows = V2_CRR_getSheetObjects_(sheet);
    var uniqueKey = V2_CRR_buildUniqueKey_({
      teacherEmail: teacherEmail,
      studentId: studentId,
      startDate: startDate,
      endDate: endDate
    });

    var reuseRecentViewId = '';

    for (var i = 0; i < allRows.length; i++) {
      var row = allRows[i];

      if (V2_CRR_buildUniqueKey_(row) === uniqueKey) {
        reuseRecentViewId = V2_CRR_toText_(row.recentViewId);
        break;
      }
    }

    var newRow = V2_CRR_buildRowObject_({
      recentViewId: reuseRecentViewId || V2_CRR_createId_(),
      teacherEmail: teacherEmail,
      teacherName: teacherName,
      studentId: studentId,
      studentName: studentName,
      classId: classId,
      className: className,
      startDate: startDate,
      endDate: endDate,
      viewedAt: viewedAt
    });

    var mergedRows = [];
    for (var j = 0; j < allRows.length; j++) {
      mergedRows.push(V2_CRR_buildRowObject_(allRows[j]));
    }
    mergedRows.push(newRow);

    var syncResult = V2_CRR_syncAllRows_(mergedRows, {
      targetTeacherEmail: teacherEmail,
      keepLimit: keepLimit
    });

    V2_CRR_writeAllRows_(sheet, syncResult.rows);

    return newRow;
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentRepository_saveRecentView', error.message, error.stack || '');
    throw error;
  }
}

function V2_ConsultReportRecentRepository_getRecentViews(criteria) {
  try {
    criteria = criteria || {};

    var teacherEmail = V2_CRR_toText_(criteria.teacherEmail) || V2_CRR_getCurrentTeacherEmail_();
    var limit = V2_CRR_toRecentViewLimit_(criteria.limit);
    var sheet = V2_CRR_getOrCreateSheet_();
    var allRows = V2_CRR_getSheetObjects_(sheet);

    var syncResult = V2_CRR_syncAllRows_(allRows, {
      targetTeacherEmail: teacherEmail,
      keepLimit: V2_CRR_getRecentViewKeepLimit_()
    });

    if (syncResult.changed) {
      V2_CRR_writeAllRows_(sheet, syncResult.rows);
    }

    var results = [];

    for (var i = 0; i < syncResult.rows.length; i++) {
      var row = syncResult.rows[i];

      if (teacherEmail && !V2_CRR_isSameText_(row.teacherEmail, teacherEmail)) {
        continue;
      }

      results.push({
        recentViewId: V2_CRR_toText_(row.recentViewId),
        teacherEmail: V2_CRR_toText_(row.teacherEmail),
        teacherName: V2_CRR_toText_(row.teacherName),
        studentId: V2_CRR_toText_(row.studentId),
        studentName: V2_CRR_toText_(row.studentName),
        classId: V2_CRR_toText_(row.classId),
        className: V2_CRR_toText_(row.className),
        startDate: V2_CRR_normalizeDateText_(row.startDate),
        endDate: V2_CRR_normalizeDateText_(row.endDate),
        viewedAt: V2_CRR_toText_(row.viewedAt)
      });

      if (results.length >= limit) {
        break;
      }
    }

    return results;
  } catch (error) {
    V2_log_('ERROR', 'V2_ConsultReportRecentRepository_getRecentViews', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_getOrCreateSheet_() {
  try {
    var ss = V2_getSpreadsheet_();
    var sheetName = 'V2_Consult_Report_Recent_Views';
    var header = V2_CRR_getHeader_();

    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

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
    return sheet;
  } catch (error) {
    throw new Error('최근 조회 시트 준비 중 오류가 발생했습니다. ' + error.message);
  }
}

function V2_CRR_getHeader_() {
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

function V2_CRR_getSheetObjects_(sheet) {
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
        item[V2_CRR_toText_(header[j])] = values[i][j];
      }
      rows.push(V2_CRR_buildRowObject_(item));
    }

    return rows;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_getSheetObjects_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_writeAllRows_(sheet, rows) {
  try {
    if (!sheet) {
      throw new Error('시트가 없습니다.');
    }

    var header = V2_CRR_getHeader_();
    var normalizedRows = rows || [];

    sheet.clearContents();
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    sheet.setFrozenRows(1);

    if (normalizedRows.length < 1) {
      return;
    }

    var values = [];
    for (var i = 0; i < normalizedRows.length; i++) {
      var row = V2_CRR_buildRowObject_(normalizedRows[i]);
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
    V2_log_('ERROR', 'V2_CRR_writeAllRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_normalizeRecentViewData_(data) {
  try {
    data = data || {};

    return {
      studentId: V2_CRR_toText_(data.studentId),
      studentName: V2_CRR_toText_(data.studentName),
      classId: V2_CRR_toText_(data.classId),
      className: V2_CRR_toText_(data.className),
      startDate: V2_CRR_normalizeDateText_(data.startDate),
      endDate: V2_CRR_normalizeDateText_(data.endDate)
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

function V2_CRR_syncAllRows_(rows, options) {
  try {
    rows = rows || [];
    options = options || {};

    var targetTeacherEmail = V2_CRR_toText_(options.targetTeacherEmail);
    var keepLimit = V2_CRR_toKeepLimit_(options.keepLimit);
    var allRows = [];
    var teacherRows = [];
    var otherRows = [];

    for (var i = 0; i < rows.length; i++) {
      var row = V2_CRR_buildRowObject_(rows[i]);

      if (targetTeacherEmail && V2_CRR_isSameText_(row.teacherEmail, targetTeacherEmail)) {
        teacherRows.push(row);
      } else {
        otherRows.push(row);
      }
    }

    var syncedTeacherRows = V2_CRR_normalizeTeacherRows_(teacherRows, keepLimit);

    allRows = otherRows.concat(syncedTeacherRows);
    allRows = V2_CRR_sortRecentRows_(allRows);

    return {
      changed: V2_CRR_hasRowSetChanged_(rows, allRows),
      rows: allRows
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_syncAllRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_normalizeTeacherRows_(teacherRows, keepLimit) {
  try {
    teacherRows = teacherRows || [];
    keepLimit = V2_CRR_toKeepLimit_(keepLimit);

    var sortedRows = V2_CRR_sortRecentRows_(teacherRows);
    var uniqueRows = [];
    var uniqueMap = {};

    for (var i = 0; i < sortedRows.length; i++) {
      var row = V2_CRR_buildRowObject_(sortedRows[i]);
      var uniqueKey = V2_CRR_buildUniqueKey_(row);

      if (!uniqueKey) {
        continue;
      }

      if (uniqueMap[uniqueKey]) {
        continue;
      }

      uniqueMap[uniqueKey] = true;
      uniqueRows.push(row);

      if (uniqueRows.length >= keepLimit) {
        break;
      }
    }

    return uniqueRows;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_normalizeTeacherRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_sortRecentRows_(rows) {
  try {
    rows = rows || [];

    var copiedRows = [];
    for (var i = 0; i < rows.length; i++) {
      copiedRows.push(V2_CRR_buildRowObject_(rows[i]));
    }

    copiedRows.sort(function(a, b) {
      var timeDiff = V2_CRR_toSortTime_(b.viewedAt) - V2_CRR_toSortTime_(a.viewedAt);
      if (timeDiff !== 0) {
        return timeDiff;
      }

      return V2_CRR_toText_(b.recentViewId).localeCompare(V2_CRR_toText_(a.recentViewId));
    });

    return copiedRows;
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_sortRecentRows_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_hasRowSetChanged_(beforeRows, afterRows) {
  try {
    beforeRows = beforeRows || [];
    afterRows = afterRows || [];

    if (beforeRows.length !== afterRows.length) {
      return true;
    }

    for (var i = 0; i < beforeRows.length; i++) {
      var beforeRow = V2_CRR_buildRowObject_(beforeRows[i]);
      var afterRow = V2_CRR_buildRowObject_(afterRows[i]);

      if (V2_CRR_rowSignature_(beforeRow) !== V2_CRR_rowSignature_(afterRow)) {
        return true;
      }
    }

    return false;
  } catch (error) {
    return true;
  }
}

function V2_CRR_rowSignature_(row) {
  try {
    row = V2_CRR_buildRowObject_(row);

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

function V2_CRR_buildRowObject_(data) {
  try {
    data = data || {};

    return {
      recentViewId: V2_CRR_toText_(data.recentViewId),
      teacherEmail: V2_CRR_toText_(data.teacherEmail),
      teacherName: V2_CRR_toText_(data.teacherName),
      studentId: V2_CRR_toText_(data.studentId),
      studentName: V2_CRR_toText_(data.studentName),
      classId: V2_CRR_toText_(data.classId),
      className: V2_CRR_toText_(data.className),
      startDate: V2_CRR_normalizeDateText_(data.startDate),
      endDate: V2_CRR_normalizeDateText_(data.endDate),
      viewedAt: V2_CRR_normalizeDateTimeText_(data.viewedAt)
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

function V2_CRR_cleanupTeacherRecentViews_(sheet, teacherEmail, keepLimit) {
  try {
    if (!sheet) {
      return;
    }

    var rows = V2_CRR_getSheetObjects_(sheet);
    var syncResult = V2_CRR_syncAllRows_(rows, {
      targetTeacherEmail: teacherEmail,
      keepLimit: keepLimit
    });

    if (syncResult.changed) {
      V2_CRR_writeAllRows_(sheet, syncResult.rows);
    }
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_cleanupTeacherRecentViews_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_enforceTeacherRecentLimit_(sheet, teacherEmail, keepLimit) {
  try {
    V2_CRR_cleanupTeacherRecentViews_(sheet, teacherEmail, keepLimit);
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_enforceTeacherRecentLimit_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_deleteRowsDescending_(sheet, rowIndexes) {
  try {
    if (!sheet || !rowIndexes || rowIndexes.length < 1) {
      return;
    }

    rowIndexes.sort(function(a, b) {
      return b - a;
    });

    for (var i = 0; i < rowIndexes.length; i++) {
      var rowIndex = Number(rowIndexes[i]);

      if (!isNaN(rowIndex) && rowIndex >= 2 && rowIndex <= sheet.getLastRow()) {
        sheet.deleteRow(rowIndex);
      }
    }
  } catch (error) {
    V2_log_('ERROR', 'V2_CRR_deleteRowsDescending_', error.message, error.stack || '');
    throw error;
  }
}

function V2_CRR_isSameRecentView_(row, target) {
  try {
    row = row || {};
    target = target || {};

    return V2_CRR_buildUniqueKey_(row) === V2_CRR_buildUniqueKey_(target);
  } catch (error) {
    return false;
  }
}

function V2_CRR_buildUniqueKey_(data) {
  try {
    data = data || {};

    return [
      V2_CRR_toText_(data.teacherEmail).toLowerCase(),
      V2_CRR_toText_(data.studentId).toLowerCase(),
      V2_CRR_normalizeDateText_(data.startDate),
      V2_CRR_normalizeDateText_(data.endDate)
    ].join('||');
  } catch (error) {
    return '';
  }
}

function V2_CRR_normalizeDateText_(value) {
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
    return V2_CRR_toText_(value);
  }
}

function V2_CRR_normalizeDateTimeText_(value) {
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
    return V2_CRR_toText_(value);
  }
}

function V2_CRR_toKeepLimit_(value) {
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

function V2_CRR_toRecentViewLimit_(value) {
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

function V2_CRR_getRecentViewKeepLimit_() {
  return 5;
}

function V2_CRR_getCurrentTeacherEmail_() {
  try {
    if (typeof V2_getCurrentUserEmail_ === 'function') {
      return V2_CRR_toText_(V2_getCurrentUserEmail_());
    }

    return V2_CRR_toText_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function V2_CRR_getCurrentTeacherName_() {
  try {
    if (typeof V2_getCurrentUserName_ === 'function') {
      return V2_CRR_toText_(V2_getCurrentUserName_());
    }

    var email = V2_CRR_getCurrentTeacherEmail_();
    return email ? email.split('@')[0] : '';
  } catch (error) {
    return '';
  }
}

function V2_CRR_nowText_() {
  try {
    if (typeof V2_nowText_ === 'function') {
      return V2_nowText_();
    }

    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return '';
  }
}

function V2_CRR_createId_() {
  try {
    if (typeof V2_createId_ === 'function') {
      return V2_createId_();
    }

    return 'CRRV_' + new Date().getTime();
  } catch (error) {
    return 'CRRV_' + new Date().getTime();
  }
}

function V2_CRR_toSortTime_(value) {
  try {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return value.getTime();
    }

    var normalizedText = V2_CRR_normalizeDateTimeText_(value);
    var parsedDate = new Date(normalizedText);

    if (!isNaN(parsedDate.getTime())) {
      return parsedDate.getTime();
    }

    return 0;
  } catch (error) {
    return 0;
  }
}

function V2_CRR_isSameText_(a, b) {
  try {
    return V2_CRR_toText_(a).toLowerCase() === V2_CRR_toText_(b).toLowerCase();
  } catch (error) {
    return false;
  }
}

function V2_CRR_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}

function V2_CRR_toLimit_(value) {
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
