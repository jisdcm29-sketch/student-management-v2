/**
 * V2_Attendance_Sheet_Repository.gs
 * [수정된 전체 파일]
 * - 기존 조회 / 저장 구조 유지
 * - 반 검색/선택 UI용 반 목록 조회 함수 추가
 */

function V2_AttendanceSheetRepository_getClassAttendanceSheet(criteria) {
  try {
    criteria = criteria || {};

    var classId = V2_ASR_toText_(criteria.classId);
    var startDateText = V2_ASR_normalizeDateText_(criteria.startDate);
    var endDateText = V2_ASR_normalizeDateText_(criteria.endDate);
    var includeInactive = V2_ASR_toBoolean_(criteria.includeInactive);

    if (!classId) {
      throw new Error('classId가 필요합니다.');
    }

    if (!startDateText || !endDateText) {
      throw new Error('시작일과 종료일이 필요합니다.');
    }

    var ss = V2_getSpreadsheet_();
    var classes = V2_ASR_getSheetObjectsByName_(ss, 'V2_Classes');
    var students = V2_ASR_getSheetObjectsByName_(ss, 'V2_Students');
    var attendanceRows = V2_ASR_getSheetObjectsByName_(ss, 'V2_Attendance');

    var classInfo = V2_ASR_findClassInfo_(classes, classId);
    if (!classInfo.classId) {
      throw new Error('반 정보를 찾을 수 없습니다. classId=' + classId);
    }

    var targetStudents = V2_ASR_filterStudentsByClass_(students, classId, includeInactive);
    var dateColumns = V2_ASR_buildDateColumns_(startDateText, endDateText);
    var attendanceMap = V2_ASR_buildAttendanceMap_(attendanceRows, startDateText, endDateText);
    var studentRows = V2_ASR_buildStudentRows_(targetStudents, dateColumns, attendanceMap);
    var summary = V2_ASR_buildSheetSummary_(studentRows);

    return {
      generatedAt: V2_ASR_nowText_(),
      classInfo: classInfo,
      period: {
        startDate: startDateText,
        endDate: endDateText,
        totalDays: dateColumns.length
      },
      summary: summary,
      dateColumns: dateColumns,
      studentRows: studentRows
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceSheetRepository_getClassAttendanceSheet', error.message, error.stack || '');
    throw error;
  }
}

function V2_AttendanceSheetRepository_getClassSearchList(criteria) {
  try {
    criteria = criteria || {};

    var searchText = V2_ASR_toText_(criteria.searchText).toLowerCase();
    var ss = V2_getSpreadsheet_();
    var classes = V2_ASR_getSheetObjectsByName_(ss, 'V2_Classes');
    var items = [];

    for (var i = 0; i < classes.length; i++) {
      var row = classes[i] || {};
      var classId = V2_ASR_getClassId_(row);
      var className = V2_ASR_getClassName_(row) || classId;
      var teacherId = V2_ASR_getTeacherId_(row);
      var searchBase = (classId + ' ' + className + ' ' + teacherId).toLowerCase();

      if (searchText && searchBase.indexOf(searchText) === -1) {
        continue;
      }

      items.push({
        classId: classId,
        className: className,
        teacherId: teacherId,
        displayName: className + ' (' + classId + ')',
        sortOrder: V2_ASR_toNumber_(V2_ASR_findValueByKeys_(row, ['sortOrder', 'sort_order', '정렬순서']))
      });
    }

    items.sort(function(a, b) {
      var sortDiff = V2_ASR_toNumber_(a.sortOrder) - V2_ASR_toNumber_(b.sortOrder);
      if (sortDiff !== 0) {
        return sortDiff;
      }
      return V2_ASR_toText_(a.className).localeCompare(V2_ASR_toText_(b.className), 'ko');
    });

    return items;
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceSheetRepository_getClassSearchList', error.message, error.stack || '');
    return [];
  }
}

function V2_AttendanceSheetRepository_saveAttendanceUploadBatch(saveRequest) {
  try {
    saveRequest = saveRequest || {};

    var items = Array.isArray(saveRequest.items) ? saveRequest.items : [];
    if (items.length < 1) {
      return {
        success: true,
        message: '저장할 출석 데이터가 없습니다.',
        data: {
          updateCount: 0,
          insertCount: 0,
          totalSavedCount: 0
        }
      };
    }

    var ss = V2_getSpreadsheet_();
    var sheet = V2_ASR_getOrCreateAttendanceSheet_(ss);
    var headerInfo = V2_ASR_getAttendanceHeaderInfo_(sheet);
    var headers = headerInfo.headers;
    var headerMap = headerInfo.headerMap;
    var values = headerInfo.values;
    var existingMap = V2_ASR_buildExistingAttendanceRowMap_(headers, values);

    var updateCount = 0;
    var insertCount = 0;
    var nowText = V2_ASR_nowText_();

    items.forEach(function(item) {
      var normalizedItem = V2_ASR_normalizeSaveItem_(item, nowText);
      var key = V2_ASR_buildAttendanceKey_(normalizedItem.studentId, normalizedItem.date);

      if (!normalizedItem.studentId || !normalizedItem.date || !normalizedItem.status) {
        return;
      }

      if (existingMap[key]) {
        var rowInfo = existingMap[key];
        var updatedRow = rowInfo.rowValues.slice();
        V2_ASR_applySaveItemToRow_(updatedRow, headerMap, normalizedItem, false);
        values[rowInfo.valueIndex] = updatedRow;
        updateCount += 1;
      } else {
        var newRow = V2_ASR_createEmptyRowByHeaderLength_(headers.length);
        V2_ASR_applySaveItemToRow_(newRow, headerMap, normalizedItem, true);
        values.push(newRow);
        existingMap[key] = {
          valueIndex: values.length - 1,
          rowValues: newRow
        };
        insertCount += 1;
      }
    });

    V2_ASR_writeAttendanceSheetValues_(sheet, headers, values);

    return {
      success: true,
      message: '출석 업로드 데이터가 저장되었습니다.',
      data: {
        updateCount: updateCount,
        insertCount: insertCount,
        totalSavedCount: updateCount + insertCount
      }
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_AttendanceSheetRepository_saveAttendanceUploadBatch', error.message, error.stack || '');
    throw error;
  }
}

function V2_ASR_findClassInfo_(classes, classId) {
  try {
    classes = classes || [];
    classId = V2_ASR_toText_(classId);

    for (var i = 0; i < classes.length; i++) {
      var item = classes[i];
      var currentClassId = V2_ASR_getClassId_(item);

      if (!V2_ASR_isSameText_(currentClassId, classId)) {
        continue;
      }

      return {
        classId: currentClassId,
        className: V2_ASR_getClassName_(item) || currentClassId,
        teacherId: V2_ASR_getTeacherId_(item),
        sortOrder: V2_ASR_toNumber_(V2_ASR_findValueByKeys_(item, ['sortOrder', 'sort_order', '정렬순서']))
      };
    }

    return {
      classId: '',
      className: '',
      teacherId: '',
      sortOrder: 0
    };
  } catch (error) {
    return {
      classId: '',
      className: '',
      teacherId: '',
      sortOrder: 0
    };
  }
}

function V2_ASR_filterStudentsByClass_(students, classId, includeInactive) {
  try {
    students = students || [];
    classId = V2_ASR_toText_(classId);
    includeInactive = V2_ASR_toBoolean_(includeInactive);

    var items = [];

    for (var i = 0; i < students.length; i++) {
      var item = students[i];
      var currentClassId = V2_ASR_getClassId_(item);
      var status = V2_ASR_getStudentStatus_(item);

      if (!V2_ASR_isSameText_(currentClassId, classId)) {
        continue;
      }

      if (!includeInactive && !V2_ASR_isActiveStudentStatus_(status)) {
        continue;
      }

      items.push({
        studentId: V2_ASR_getStudentId_(item),
        studentName: V2_ASR_getStudentName_(item),
        classId: currentClassId,
        status: status,
        parentName: V2_ASR_findValueByKeys_(item, ['parentName', '학부모명']),
        parentPhone: V2_ASR_findValueByKeys_(item, ['parentPhone', '학부모연락처']),
        memo: V2_ASR_findValueByKeys_(item, ['memo', 'notes', '비고'])
      });
    }

    items.sort(function(a, b) {
      return V2_ASR_toText_(a.studentName).localeCompare(V2_ASR_toText_(b.studentName), 'ko');
    });

    return items;
  } catch (error) {
    return [];
  }
}

function V2_ASR_buildDateColumns_(startDateText, endDateText) {
  try {
    var startDate = V2_ASR_toDate_(startDateText);
    var endDate = V2_ASR_toDate_(endDateText);

    if (!startDate || !endDate) {
      return [];
    }

    var columns = [];
    var current = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    var end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

    while (current.getTime() <= end.getTime()) {
      var dateText = Utilities.formatDate(current, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      columns.push({
        date: dateText,
        dayOfWeek: V2_ASR_getDayOfWeekLabel_(current),
        displayLabel: dateText + ' (' + V2_ASR_getDayOfWeekLabel_(current) + ')'
      });
      current.setDate(current.getDate() + 1);
    }

    return columns;
  } catch (error) {
    return [];
  }
}

function V2_ASR_buildAttendanceMap_(attendanceRows, startDateText, endDateText) {
  try {
    attendanceRows = attendanceRows || [];
    var startDate = V2_ASR_toDate_(startDateText);
    var endDate = V2_ASR_toDate_(endDateText);
    var map = {};

    for (var i = 0; i < attendanceRows.length; i++) {
      var item = attendanceRows[i];
      var studentId = V2_ASR_getStudentId_(item);
      var dateText = V2_ASR_getAttendanceDateText_(item);
      var status = V2_ASR_getAttendanceStatus_(item);
      var teacherName = V2_ASR_findValueByKeys_(item, ['teacherName', '교사명']);
      var memo = V2_ASR_findValueByKeys_(item, ['memo', '비고']);
      var updatedAt = V2_ASR_findValueByKeys_(item, ['updatedAt', '수정일']);

      if (!studentId || !dateText) {
        continue;
      }

      if (!V2_ASR_isDateTextInRange_(dateText, startDate, endDate)) {
        continue;
      }

      map[V2_ASR_buildAttendanceKey_(studentId, dateText)] = {
        status: status,
        teacherName: teacherName,
        memo: memo,
        updatedAt: updatedAt
      };
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_ASR_buildStudentRows_(students, dateColumns, attendanceMap) {
  try {
    students = students || [];
    dateColumns = dateColumns || [];
    attendanceMap = attendanceMap || {};

    var rows = [];

    for (var i = 0; i < students.length; i++) {
      var student = students[i];
      var dateStatuses = [];
      var counts = {
        totalMarkedCount: 0,
        presentCount: 0,
        lateCount: 0,
        absentCount: 0,
        excusedCount: 0,
        emptyCount: 0
      };

      for (var j = 0; j < dateColumns.length; j++) {
        var column = dateColumns[j];
        var key = V2_ASR_buildAttendanceKey_(student.studentId, column.date);
        var attendanceItem = attendanceMap[key] || null;
        var status = attendanceItem ? V2_ASR_toText_(attendanceItem.status) : '';

        dateStatuses.push({
          date: column.date,
          dayOfWeek: column.dayOfWeek,
          displayLabel: column.displayLabel,
          status: status,
          teacherName: attendanceItem ? attendanceItem.teacherName : '',
          memo: attendanceItem ? attendanceItem.memo : '',
          updatedAt: attendanceItem ? attendanceItem.updatedAt : ''
        });

        if (!status) {
          counts.emptyCount += 1;
          continue;
        }

        counts.totalMarkedCount += 1;

        if (V2_ASR_isPresentStatus_(status)) {
          counts.presentCount += 1;
        } else if (V2_ASR_isLateStatus_(status)) {
          counts.lateCount += 1;
        } else if (V2_ASR_isExcusedStatus_(status)) {
          counts.excusedCount += 1;
        } else {
          counts.absentCount += 1;
        }
      }

      rows.push({
        studentId: student.studentId,
        studentName: student.studentName,
        classId: student.classId,
        status: student.status,
        parentName: student.parentName,
        parentPhone: student.parentPhone,
        memo: student.memo,
        counts: counts,
        dateStatuses: dateStatuses
      });
    }

    return rows;
  } catch (error) {
    return [];
  }
}

function V2_ASR_buildSheetSummary_(studentRows) {
  try {
    studentRows = studentRows || [];

    var summary = {
      totalStudentCount: studentRows.length,
      totalAttendanceCount: 0,
      presentCount: 0,
      lateCount: 0,
      absentCount: 0,
      excusedCount: 0
    };

    for (var i = 0; i < studentRows.length; i++) {
      var counts = studentRows[i].counts || {};

      summary.totalAttendanceCount += V2_ASR_toNumber_(counts.totalMarkedCount);
      summary.presentCount += V2_ASR_toNumber_(counts.presentCount);
      summary.lateCount += V2_ASR_toNumber_(counts.lateCount);
      summary.absentCount += V2_ASR_toNumber_(counts.absentCount);
      summary.excusedCount += V2_ASR_toNumber_(counts.excusedCount);
    }

    return summary;
  } catch (error) {
    return {
      totalStudentCount: 0,
      totalAttendanceCount: 0,
      presentCount: 0,
      lateCount: 0,
      absentCount: 0,
      excusedCount: 0
    };
  }
}

function V2_ASR_getSheetObjectsByName_(ss, sheetName) {
  try {
    var sheet = ss.getSheetByName(sheetName);

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
        item[V2_ASR_toText_(header[j])] = values[i][j];
      }
      rows.push(item);
    }

    return rows;
  } catch (error) {
    V2_log_('ERROR', 'V2_ASR_getSheetObjectsByName_', error.message, error.stack || '');
    throw error;
  }
}

function V2_ASR_getOrCreateAttendanceSheet_(ss) {
  try {
    var sheet = ss.getSheetByName('V2_Attendance');
    var requiredHeaders = [
      'studentId',
      'date',
      'status',
      'teacherName',
      'updatedAt',
      'studentNameSnapshot',
      'classIdSnapshot',
      'classNameSnapshot',
      'studentStatusSnapshot',
      'uploadSource',
      'uploadAppliedAt',
      'memo'
    ];

    if (!sheet) {
      sheet = ss.insertSheet('V2_Attendance');
      sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
      sheet.setFrozenRows(1);
      return sheet;
    }

    V2_ASR_ensureRequiredHeaders_(sheet, requiredHeaders);
    sheet.setFrozenRows(1);
    return sheet;
  } catch (error) {
    throw new Error('V2_ASR_getOrCreateAttendanceSheet_ 오류: ' + error.message);
  }
}

function V2_ASR_ensureRequiredHeaders_(sheet, requiredHeaders) {
  try {
    requiredHeaders = Array.isArray(requiredHeaders) ? requiredHeaders : [];

    var lastColumn = sheet.getLastColumn();
    var existingHeaders = lastColumn > 0
      ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
      : [];

    if (existingHeaders.length < 1) {
      sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
      return;
    }

    var normalizedExisting = existingHeaders.map(function(item) {
      return V2_ASR_toText_(item);
    });

    var appendHeaders = [];
    requiredHeaders.forEach(function(requiredHeader) {
      if (normalizedExisting.indexOf(requiredHeader) === -1) {
        appendHeaders.push(requiredHeader);
      }
    });

    if (appendHeaders.length > 0) {
      sheet.getRange(1, normalizedExisting.length + 1, 1, appendHeaders.length).setValues([appendHeaders]);
    }
  } catch (error) {
    throw new Error('V2_ASR_ensureRequiredHeaders_ 오류: ' + error.message);
  }
}

function V2_ASR_getAttendanceHeaderInfo_(sheet) {
  try {
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var headers = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];
    var values = (lastRow >= 2 && lastColumn > 0)
      ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
      : [];

    return {
      headers: headers,
      headerMap: V2_ASR_buildHeaderMap_(headers),
      values: values
    };
  } catch (error) {
    throw new Error('V2_ASR_getAttendanceHeaderInfo_ 오류: ' + error.message);
  }
}

function V2_ASR_buildHeaderMap_(headers) {
  try {
    headers = Array.isArray(headers) ? headers : [];
    var map = {};

    headers.forEach(function(header, index) {
      map[V2_ASR_toText_(header)] = index;
    });

    return map;
  } catch (error) {
    return {};
  }
}

function V2_ASR_buildExistingAttendanceRowMap_(headers, values) {
  try {
    headers = Array.isArray(headers) ? headers : [];
    values = Array.isArray(values) ? values : [];

    var headerMap = V2_ASR_buildHeaderMap_(headers);
    var map = {};

    for (var i = 0; i < values.length; i++) {
      var rowValues = values[i] || [];
      var studentId = V2_ASR_getRowValueByAliases_(rowValues, headerMap, ['studentId', 'student_id', '학생ID', '학생Id']);
      var dateText = V2_ASR_normalizeDateText_(V2_ASR_getRowValueByAliases_(rowValues, headerMap, ['date', 'attendanceDate', '출석일', '날짜']));

      if (!studentId || !dateText) {
        continue;
      }

      map[V2_ASR_buildAttendanceKey_(studentId, dateText)] = {
        valueIndex: i,
        rowValues: rowValues
      };
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_ASR_getRowValueByAliases_(rowValues, headerMap, aliases) {
  try {
    rowValues = Array.isArray(rowValues) ? rowValues : [];
    headerMap = headerMap || {};
    aliases = Array.isArray(aliases) ? aliases : [];

    for (var i = 0; i < aliases.length; i++) {
      var alias = aliases[i];
      if (headerMap[alias] !== undefined) {
        return rowValues[headerMap[alias]];
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_ASR_setRowValueByAliases_(rowValues, headerMap, aliases, value) {
  try {
    rowValues = Array.isArray(rowValues) ? rowValues : [];
    headerMap = headerMap || {};
    aliases = Array.isArray(aliases) ? aliases : [];

    for (var i = 0; i < aliases.length; i++) {
      var alias = aliases[i];
      if (headerMap[alias] !== undefined) {
        rowValues[headerMap[alias]] = value;
        return;
      }
    }
  } catch (error) {}
}

function V2_ASR_applySaveItemToRow_(rowValues, headerMap, item, isInsert) {
  try {
    rowValues = Array.isArray(rowValues) ? rowValues : [];
    headerMap = headerMap || {};
    item = item || {};

    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['studentId', 'student_id', '학생ID', '학생Id'], item.studentId);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['date', 'attendanceDate', '출석일', '날짜'], item.date);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['status', 'attendanceStatus', '출석상태'], item.status);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['teacherName', '교사명'], item.teacherName);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['updatedAt', '수정일'], item.updatedAt);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['studentNameSnapshot'], item.studentNameSnapshot);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['classIdSnapshot'], item.classIdSnapshot);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['classNameSnapshot'], item.classNameSnapshot);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['studentStatusSnapshot'], item.studentStatusSnapshot);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['uploadSource'], item.uploadSource);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['uploadAppliedAt'], item.uploadAppliedAt);
    V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['memo', '비고'], item.memo);

    if (isInsert) {
      V2_ASR_setRowValueByAliases_(rowValues, headerMap, ['createdAt', '생성일'], item.createdAt);
    }
  } catch (error) {
    throw new Error('V2_ASR_applySaveItemToRow_ 오류: ' + error.message);
  }
}

function V2_ASR_normalizeSaveItem_(item, nowText) {
  try {
    item = item || {};
    nowText = V2_ASR_toText_(nowText) || V2_ASR_nowText_();

    return {
      studentId: V2_ASR_toText_(item.studentId),
      date: V2_ASR_normalizeDateText_(item.date),
      status: V2_ASR_toText_(item.status),
      teacherName: V2_ASR_toText_(item.teacherName),
      updatedAt: V2_ASR_toText_(item.updatedAt) || nowText,
      createdAt: V2_ASR_toText_(item.createdAt) || nowText,
      studentNameSnapshot: V2_ASR_toText_(item.studentNameSnapshot),
      classIdSnapshot: V2_ASR_toText_(item.classIdSnapshot),
      classNameSnapshot: V2_ASR_toText_(item.classNameSnapshot),
      studentStatusSnapshot: V2_ASR_toText_(item.studentStatusSnapshot),
      uploadSource: V2_ASR_toText_(item.uploadSource),
      uploadAppliedAt: V2_ASR_toText_(item.uploadAppliedAt) || nowText,
      memo: V2_ASR_toText_(item.memo)
    };
  } catch (error) {
    return {
      studentId: '',
      date: '',
      status: '',
      teacherName: '',
      updatedAt: nowText || '',
      createdAt: nowText || '',
      studentNameSnapshot: '',
      classIdSnapshot: '',
      classNameSnapshot: '',
      studentStatusSnapshot: '',
      uploadSource: '',
      uploadAppliedAt: nowText || '',
      memo: ''
    };
  }
}

function V2_ASR_createEmptyRowByHeaderLength_(length) {
  try {
    var size = V2_ASR_toNumber_(length);
    if (size < 1) {
      size = 1;
    }

    var row = [];
    for (var i = 0; i < size; i++) {
      row.push('');
    }
    return row;
  } catch (error) {
    return [''];
  }
}

function V2_ASR_writeAttendanceSheetValues_(sheet, headers, values) {
  try {
    headers = Array.isArray(headers) ? headers : [];
    values = Array.isArray(values) ? values : [];

    if (headers.length < 1) {
      return;
    }

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    var maxRows = sheet.getMaxRows();
    if (maxRows > 1) {
      sheet.getRange(2, 1, maxRows - 1, headers.length).clearContent();
    }

    if (values.length > 0) {
      sheet.getRange(2, 1, values.length, headers.length).setValues(values);
    }
  } catch (error) {
    throw new Error('V2_ASR_writeAttendanceSheetValues_ 오류: ' + error.message);
  }
}

function V2_ASR_getStudentId_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['studentId', 'student_id', '학생ID', '학생Id']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getStudentName_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['studentName', 'student_name', '학생명', 'name']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getClassId_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['classId', 'class_id', '반ID', '반Id']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getClassName_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['className', 'class_name', '반명', '반이름']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getTeacherId_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['teacherId', 'teacher_id', '교사ID']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getStudentStatus_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['status', 'studentStatus', '학생상태']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getAttendanceStatus_(row) {
  try {
    return V2_ASR_findValueByKeys_(row, ['status', 'attendanceStatus', '출석상태']);
  } catch (error) {
    return '';
  }
}

function V2_ASR_getAttendanceDateText_(row) {
  try {
    var dateValue = V2_ASR_findDateByKeys_(row, ['date', 'attendanceDate', '출석일', '날짜']);

    if (!dateValue) {
      return '';
    }

    return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    return '';
  }
}

function V2_ASR_findValueByKeys_(row, keys) {
  try {
    row = row || {};
    keys = keys || [];

    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];

      if (Object.prototype.hasOwnProperty.call(row, key)) {
        return V2_ASR_toText_(row[key]);
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_ASR_findDateByKeys_(row, keys) {
  try {
    row = row || {};
    keys = keys || [];

    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];

      if (!Object.prototype.hasOwnProperty.call(row, key)) {
        continue;
      }

      var dateValue = V2_ASR_toDate_(row[key]);
      if (dateValue) {
        return dateValue;
      }
    }

    return null;
  } catch (error) {
    return null;
  }
}

function V2_ASR_buildAttendanceKey_(studentId, dateText) {
  try {
    return [
      V2_ASR_toText_(studentId),
      V2_ASR_toText_(dateText)
    ].join('||');
  } catch (error) {
    return '';
  }
}

function V2_ASR_isDateTextInRange_(dateText, startDate, endDate) {
  try {
    var target = V2_ASR_toDate_(dateText);

    if (!target || !startDate || !endDate) {
      return false;
    }

    var targetTime = new Date(target.getFullYear(), target.getMonth(), target.getDate()).getTime();
    var startTime = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate()).getTime();
    var endTime = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate()).getTime();

    return targetTime >= startTime && targetTime <= endTime;
  } catch (error) {
    return false;
  }
}

function V2_ASR_getDayOfWeekLabel_(dateValue) {
  try {
    var labels = ['일', '월', '화', '수', '목', '금', '토'];
    return labels[dateValue.getDay()] || '';
  } catch (error) {
    return '';
  }
}

function V2_ASR_isActiveStudentStatus_(status) {
  try {
    return V2_ASR_toText_(status) === '재학';
  } catch (error) {
    return false;
  }
}

function V2_ASR_isPresentStatus_(status) {
  try {
    var text = V2_ASR_toText_(status).toLowerCase();
    return text === '출석' || text === '재석' || text === 'present';
  } catch (error) {
    return false;
  }
}

function V2_ASR_isLateStatus_(status) {
  try {
    var text = V2_ASR_toText_(status).toLowerCase();
    return text === '지각' || text === 'late';
  } catch (error) {
    return false;
  }
}

function V2_ASR_isExcusedStatus_(status) {
  try {
    var text = V2_ASR_toText_(status).toLowerCase();
    return text === '공결' || text === '인정결석' || text === 'excused';
  } catch (error) {
    return false;
  }
}

function V2_ASR_normalizeDateText_(value) {
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
    return V2_ASR_toText_(value);
  }
}

function V2_ASR_toDate_(value) {
  try {
    if (!value && value !== 0) {
      return null;
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return value;
    }

    var normalizedText = V2_ASR_normalizeDateText_(value);
    var date = new Date(normalizedText);

    if (isNaN(date.getTime())) {
      return null;
    }

    return date;
  } catch (error) {
    return null;
  }
}

function V2_ASR_nowText_() {
  try {
    if (typeof V2_nowText_ === 'function') {
      return V2_nowText_();
    }

    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return '';
  }
}

function V2_ASR_toBoolean_(value) {
  try {
    if (value === true || value === 'true' || value === 'TRUE' || value === 1 || value === '1') {
      return true;
    }

    return false;
  } catch (error) {
    return false;
  }
}

function V2_ASR_toNumber_(value) {
  try {
    var numberValue = Number(value);

    if (isNaN(numberValue)) {
      return 0;
    }

    return numberValue;
  } catch (error) {
    return 0;
  }
}

function V2_ASR_isSameText_(a, b) {
  try {
    return V2_ASR_toText_(a).toLowerCase() === V2_ASR_toText_(b).toLowerCase();
  } catch (error) {
    return false;
  }
}

function V2_ASR_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
