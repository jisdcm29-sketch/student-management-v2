/**
 * V2_Student_Search_Repository.gs
 * 학생 검색 전용 Repository
 * - 시트 접근 전용
 * - V2_Students / V2_Classes 조회
 * - studentId / studentName / classId / className / parentName / parentPhone 기준 검색
 * - 완전삭제 상태 기본 제외
 */

function V2_StudentSearchRepository_searchStudents(criteria) {
  try {
    criteria = criteria || {};

    var keyword = V2_SSR_toText_(criteria.keyword).toLowerCase();
    var status = V2_SSR_toText_(criteria.status);
    var classId = V2_SSR_toText_(criteria.classId);
    var limit = V2_SSR_toLimit_(criteria.limit);

    var studentsSheetName = V2_SSR_getSheetName_('STUDENTS', 'V2_Students');
    var classesSheetName = V2_SSR_getSheetName_('CLASSES', 'V2_Classes');

    var students = V2_SSR_getSheetObjects_(studentsSheetName);
    var classMap = V2_SSR_getClassMap_(classesSheetName);

    var results = [];

    for (var i = 0; i < students.length; i++) {
      var row = students[i];

      var currentStudentId = V2_SSR_pickValue_(row, ['studentId']);
      var currentStudentName = V2_SSR_pickValue_(row, ['studentName', 'name']);
      var currentClassId = V2_SSR_pickValue_(row, ['classId']);
      var currentStatus = V2_SSR_pickValue_(row, ['status']);
      var currentParentName = V2_SSR_pickValue_(row, ['parentName']);
      var currentParentPhone = V2_SSR_pickValue_(row, ['parentPhone']);
      var currentMemo = V2_SSR_pickValue_(row, ['memo', 'notes', 'note']);
      var currentCreatedAt = V2_SSR_pickValue_(row, ['createdAt']);
      var currentUpdatedAt = V2_SSR_pickValue_(row, ['updatedAt']);

      var currentClassInfo = classMap[currentClassId] || null;
      var currentClassName = currentClassInfo ? V2_SSR_toText_(currentClassInfo.className) : '';
      var currentSortOrder = currentClassInfo ? V2_SSR_toNumber_(currentClassInfo.sortOrder, 999999) : 999999;

      if (!currentStudentId) {
        continue;
      }

      if (V2_SSR_isHardDeletedStatus_(currentStatus)) {
        continue;
      }

      if (status && !V2_SSR_isSameText_(currentStatus, status)) {
        continue;
      }

      if (classId && !V2_SSR_isSameText_(currentClassId, classId)) {
        continue;
      }

      if (keyword) {
        var searchTarget = (
          currentStudentId + ' ' +
          currentStudentName + ' ' +
          currentClassId + ' ' +
          currentClassName + ' ' +
          currentParentName + ' ' +
          currentParentPhone + ' ' +
          currentMemo
        ).toLowerCase();

        if (searchTarget.indexOf(keyword) === -1) {
          continue;
        }
      }

      results.push({
        studentId: currentStudentId,
        studentName: currentStudentName,
        classId: currentClassId,
        className: currentClassName,
        status: currentStatus,
        parentName: currentParentName,
        parentPhone: currentParentPhone,
        memo: currentMemo,
        createdAt: currentCreatedAt,
        updatedAt: currentUpdatedAt,
        classSortOrder: currentSortOrder
      });
    }

    results.sort(function(a, b) {
      return V2_SSR_compareStudents_(a, b);
    });

    if (results.length > limit) {
      results = results.slice(0, limit);
    }

    return results.map(function(item) {
      return {
        studentId: item.studentId,
        studentName: item.studentName,
        classId: item.classId,
        className: item.className,
        status: item.status,
        parentName: item.parentName,
        parentPhone: item.parentPhone,
        memo: item.memo,
        createdAt: item.createdAt,
        updatedAt: item.updatedAt
      };
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_StudentSearchRepository_searchStudents', error.message, error.stack || '');
    throw error;
  }
}

function V2_SSR_getStudentByStudentId(studentId) {
  try {
    var targetStudentId = V2_SSR_toText_(studentId);
    if (!targetStudentId) {
      throw new Error('studentId가 필요합니다.');
    }

    var studentsSheetName = V2_SSR_getSheetName_('STUDENTS', 'V2_Students');
    var classesSheetName = V2_SSR_getSheetName_('CLASSES', 'V2_Classes');

    var students = V2_SSR_getSheetObjects_(studentsSheetName);
    var classMap = V2_SSR_getClassMap_(classesSheetName);

    for (var i = 0; i < students.length; i++) {
      var row = students[i];
      var currentStudentId = V2_SSR_pickValue_(row, ['studentId']);

      if (!V2_SSR_isSameText_(currentStudentId, targetStudentId)) {
        continue;
      }

      var currentClassId = V2_SSR_pickValue_(row, ['classId']);
      var currentClassInfo = classMap[currentClassId] || null;
      var currentClassName = currentClassInfo ? V2_SSR_toText_(currentClassInfo.className) : '';

      return {
        studentId: currentStudentId,
        studentName: V2_SSR_pickValue_(row, ['studentName', 'name']),
        classId: currentClassId,
        className: currentClassName,
        status: V2_SSR_pickValue_(row, ['status']),
        parentName: V2_SSR_pickValue_(row, ['parentName']),
        parentPhone: V2_SSR_pickValue_(row, ['parentPhone']),
        memo: V2_SSR_pickValue_(row, ['memo', 'notes', 'note']),
        createdAt: V2_SSR_pickValue_(row, ['createdAt']),
        updatedAt: V2_SSR_pickValue_(row, ['updatedAt'])
      };
    }

    return null;
  } catch (error) {
    V2_log_('ERROR', 'V2_SSR_getStudentByStudentId', error.message, error.stack || '');
    throw error;
  }
}

function V2_SSR_getClassMap_(sheetName) {
  try {
    var rows = V2_SSR_getSheetObjects_(sheetName);
    var map = {};

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var currentClassId = V2_SSR_pickValue_(row, ['classId']);

      if (!currentClassId) {
        continue;
      }

      map[currentClassId] = {
        classId: currentClassId,
        className: V2_SSR_pickValue_(row, ['className']),
        teacherId: V2_SSR_pickValue_(row, ['teacherId']),
        sortOrder: V2_SSR_toNumber_(row.sortOrder, 999999),
        isActive: row.isActive,
        createdAt: V2_SSR_pickValue_(row, ['createdAt']),
        updatedAt: V2_SSR_pickValue_(row, ['updatedAt'])
      };
    }

    return map;
  } catch (error) {
    V2_log_('ERROR', 'V2_SSR_getClassMap_', error.message, error.stack || '');
    throw error;
  }
}

function V2_SSR_getSheetObjects_(sheetName) {
  try {
    var ss = V2_getSpreadsheet_();
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
      var obj = {};
      for (var j = 0; j < header.length; j++) {
        obj[String(header[j]).trim()] = values[i][j];
      }
      rows.push(obj);
    }

    return rows;
  } catch (error) {
    V2_log_('ERROR', 'V2_SSR_getSheetObjects_', error.message, error.stack || '');
    throw error;
  }
}

function V2_SSR_compareStudents_(a, b) {
  try {
    var activeA = V2_SSR_isActiveStatus_(a.status) ? 0 : 1;
    var activeB = V2_SSR_isActiveStatus_(b.status) ? 0 : 1;

    if (activeA !== activeB) {
      return activeA - activeB;
    }

    var sortOrderA = V2_SSR_toNumber_(a.classSortOrder, 999999);
    var sortOrderB = V2_SSR_toNumber_(b.classSortOrder, 999999);

    if (sortOrderA !== sortOrderB) {
      return sortOrderA - sortOrderB;
    }

    var classCompare = V2_SSR_toText_(a.classId).localeCompare(V2_SSR_toText_(b.classId));
    if (classCompare !== 0) {
      return classCompare;
    }

    var nameCompare = V2_SSR_toText_(a.studentName).localeCompare(V2_SSR_toText_(b.studentName));
    if (nameCompare !== 0) {
      return nameCompare;
    }

    return V2_SSR_toText_(a.studentId).localeCompare(V2_SSR_toText_(b.studentId));
  } catch (error) {
    return 0;
  }
}

function V2_SSR_isActiveStatus_(status) {
  try {
    var text = V2_SSR_toText_(status);
    return text === '재학' || text.toUpperCase() === 'ACTIVE';
  } catch (error) {
    return false;
  }
}

function V2_SSR_isHardDeletedStatus_(status) {
  try {
    var text = V2_SSR_toText_(status);
    return text === '완전삭제' || text.toUpperCase() === 'DELETED';
  } catch (error) {
    return false;
  }
}

function V2_SSR_isSameText_(a, b) {
  try {
    return V2_SSR_toText_(a).toLowerCase() === V2_SSR_toText_(b).toLowerCase();
  } catch (error) {
    return false;
  }
}

function V2_SSR_pickValue_(row, keys) {
  try {
    if (!row || !keys || !keys.length) {
      return '';
    }

    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (Object.prototype.hasOwnProperty.call(row, key)) {
        return V2_SSR_toText_(row[key]);
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_SSR_getSheetName_(configKey, fallbackName) {
  try {
    if (
      typeof V2_CONFIG !== 'undefined' &&
      V2_CONFIG &&
      V2_CONFIG.SHEETS &&
      V2_CONFIG.SHEETS[configKey]
    ) {
      return V2_CONFIG.SHEETS[configKey];
    }

    return fallbackName;
  } catch (error) {
    return fallbackName;
  }
}

function V2_SSR_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  } catch (error) {
    return '';
  }
}

function V2_SSR_toLimit_(value) {
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

function V2_SSR_toNumber_(value, defaultValue) {
  try {
    var numberValue = Number(value);
    if (isNaN(numberValue)) {
      return defaultValue;
    }
    return numberValue;
  } catch (error) {
    return defaultValue;
  }
}
