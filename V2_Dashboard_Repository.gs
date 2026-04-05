/**
 * V2_Dashboard_Repository.gs
 * 홈 대시보드 Repository
 * - 시트 조회 전용
 * - 반별 학생 수
 * - 최근 7일 출석률
 * - 반 평균 점수
 * - 로그인 교사 기준 반 필터링
 * - 출석 시트에 classId가 없어도 studentId 기준으로 반 연결
 */

function V2_DashboardRepository_getDashboardSnapshot(criteria) {
  try {
    criteria = criteria || {};

    var days = V2_DBR_toDays_(criteria.days);
    var ss = V2_getSpreadsheet_();
    var teacherEmail = V2_DBR_getCurrentTeacherEmail_();
    var teacherName = V2_DBR_getCurrentTeacherName_();

    var classes = V2_DBR_getSheetObjectsByName_(ss, 'V2_Classes');
    var students = V2_DBR_getSheetObjectsByName_(ss, 'V2_Students');
    var attendance = V2_DBR_getSheetObjectsByName_(ss, 'V2_Attendance');
    var scores = V2_DBR_getSheetObjectsByName_(ss, 'V2_Scores');

    var filteredClasses = V2_DBR_filterClassesByTeacher_(classes, teacherEmail);
    var allowedClassIdMap = V2_DBR_buildAllowedClassIdMap_(filteredClasses);
    var studentClassMap = V2_DBR_buildStudentClassMap_(students);

    var filteredStudents = V2_DBR_filterRowsByAllowedClassIds_(students, allowedClassIdMap, studentClassMap);
    var filteredAttendance = V2_DBR_filterRowsByAllowedClassIds_(attendance, allowedClassIdMap, studentClassMap);
    var filteredScores = V2_DBR_filterRowsByAllowedClassIds_(scores, allowedClassIdMap, studentClassMap);

    var classMap = V2_DBR_buildClassMap_(filteredClasses);
    var studentCountMap = V2_DBR_buildStudentCountMap_(filteredStudents);
    var attendanceSummaryMap = V2_DBR_buildAttendanceSummaryMap_(filteredAttendance, days, studentClassMap);
    var scoreSummaryMap = V2_DBR_buildScoreSummaryMap_(filteredScores, studentClassMap);

    var items = V2_DBR_buildDashboardItems_(
      filteredClasses,
      classMap,
      studentCountMap,
      attendanceSummaryMap,
      scoreSummaryMap
    );

    items.sort(function(a, b) {
      var sortDiff = V2_DBR_toNumber_(a.sortOrder) - V2_DBR_toNumber_(b.sortOrder);
      if (sortDiff !== 0) {
        return sortDiff;
      }

      return V2_DBR_toText_(a.className).localeCompare(V2_DBR_toText_(b.className), 'ko');
    });

    return {
      generatedAt: V2_DBR_nowText_(),
      periodDays: days,
      teacherEmail: teacherEmail,
      teacherName: teacherName,
      totalClassCount: items.length,
      totalStudentCount: V2_DBR_sumNumberByKey_(items, 'studentCount'),
      items: items
    };
  } catch (error) {
    V2_log_('ERROR', 'V2_DashboardRepository_getDashboardSnapshot', error.message, error.stack || '');
    throw error;
  }
}

function V2_DBR_buildDashboardItems_(classes, classMap, studentCountMap, attendanceSummaryMap, scoreSummaryMap) {
  try {
    var items = [];
    var addedClassIdMap = {};

    for (var i = 0; i < classes.length; i++) {
      var classItem = classes[i];
      var classId = V2_DBR_getClassId_(classItem);
      var className = V2_DBR_getClassName_(classItem);
      var sortOrder = V2_DBR_getSortOrder_(classItem);

      if (!classId) {
        continue;
      }

      addedClassIdMap[classId] = true;
      items.push(V2_DBR_buildSingleDashboardItem_(
        classId,
        className,
        sortOrder,
        studentCountMap[classId],
        attendanceSummaryMap[classId],
        scoreSummaryMap[classId]
      ));
    }

    for (var classId in studentCountMap) {
      if (!Object.prototype.hasOwnProperty.call(studentCountMap, classId)) {
        continue;
      }

      if (addedClassIdMap[classId]) {
        continue;
      }

      items.push(V2_DBR_buildSingleDashboardItem_(
        classId,
        V2_DBR_getClassNameFromMap_(classMap, classId),
        V2_DBR_getClassSortOrderFromMap_(classMap, classId),
        studentCountMap[classId],
        attendanceSummaryMap[classId],
        scoreSummaryMap[classId]
      ));
    }

    return items;
  } catch (error) {
    V2_log_('ERROR', 'V2_DBR_buildDashboardItems_', error.message, error.stack || '');
    throw error;
  }
}

function V2_DBR_buildSingleDashboardItem_(classId, className, sortOrder, studentCount, attendanceSummary, scoreSummary) {
  try {
    var attendanceTotal = attendanceSummary ? attendanceSummary.totalCount : 0;
    var presentCount = attendanceSummary ? attendanceSummary.presentCount : 0;
    var absentCount = attendanceSummary ? attendanceSummary.absentCount : 0;
    var lateCount = attendanceSummary ? attendanceSummary.lateCount : 0;
    var excusedCount = attendanceSummary ? attendanceSummary.excusedCount : 0;
    var averageScore = scoreSummary ? scoreSummary.averageScore : '';

    return {
      classId: V2_DBR_toText_(classId),
      className: V2_DBR_toText_(className) || V2_DBR_toText_(classId),
      sortOrder: V2_DBR_toNumber_(sortOrder),
      studentCount: V2_DBR_toNumber_(studentCount),
      attendanceRate7Days: V2_DBR_calculateAttendanceRateText_(attendanceSummary),
      attendanceCount7Days: attendanceTotal,
      presentCount7Days: presentCount,
      absentCount7Days: absentCount,
      lateCount7Days: lateCount,
      excusedCount7Days: excusedCount,
      averageScore: averageScore,
      averageScoreText: averageScore === '' ? '' : String(averageScore)
    };
  } catch (error) {
    return {
      classId: V2_DBR_toText_(classId),
      className: V2_DBR_toText_(className) || V2_DBR_toText_(classId),
      sortOrder: V2_DBR_toNumber_(sortOrder),
      studentCount: 0,
      attendanceRate7Days: '0%',
      attendanceCount7Days: 0,
      presentCount7Days: 0,
      absentCount7Days: 0,
      lateCount7Days: 0,
      excusedCount7Days: 0,
      averageScore: '',
      averageScoreText: ''
    };
  }
}

function V2_DBR_buildClassMap_(classes) {
  try {
    var map = {};

    for (var i = 0; i < classes.length; i++) {
      var item = classes[i];
      var classId = V2_DBR_getClassId_(item);

      if (!classId) {
        continue;
      }

      map[classId] = {
        classId: classId,
        className: V2_DBR_getClassName_(item),
        sortOrder: V2_DBR_getSortOrder_(item)
      };
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_DBR_buildAllowedClassIdMap_(classes) {
  try {
    var map = {};

    for (var i = 0; i < classes.length; i++) {
      var classId = V2_DBR_getClassId_(classes[i]);
      if (!classId) {
        continue;
      }

      map[classId] = true;
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_DBR_buildStudentClassMap_(students) {
  try {
    var map = {};

    for (var i = 0; i < students.length; i++) {
      var item = students[i];
      var studentId = V2_DBR_getStudentId_(item);
      var classId = V2_DBR_getClassId_(item);

      if (!studentId || !classId) {
        continue;
      }

      map[studentId] = classId;
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_DBR_filterClassesByTeacher_(classes, teacherEmail) {
  try {
    classes = classes || [];
    teacherEmail = V2_DBR_toText_(teacherEmail);

    if (!teacherEmail) {
      return [];
    }

    var result = [];

    for (var i = 0; i < classes.length; i++) {
      var item = classes[i];
      var classTeacherEmail = V2_DBR_getClassTeacherEmail_(item);

      if (classTeacherEmail && V2_DBR_isSameText_(classTeacherEmail, teacherEmail)) {
        result.push(item);
      }
    }

    return result;
  } catch (error) {
    return [];
  }
}

function V2_DBR_filterRowsByAllowedClassIds_(rows, allowedClassIdMap, studentClassMap) {
  try {
    rows = rows || [];
    allowedClassIdMap = allowedClassIdMap || {};
    studentClassMap = studentClassMap || {};

    var result = [];

    for (var i = 0; i < rows.length; i++) {
      var item = rows[i];
      var classId = V2_DBR_resolveClassIdFromRow_(item, studentClassMap);

      if (!classId) {
        continue;
      }

      if (!allowedClassIdMap[classId]) {
        continue;
      }

      result.push(item);
    }

    return result;
  } catch (error) {
    return [];
  }
}

function V2_DBR_buildStudentCountMap_(students) {
  try {
    var map = {};

    for (var i = 0; i < students.length; i++) {
      var item = students[i];
      var classId = V2_DBR_getClassId_(item);
      var status = V2_DBR_getStudentStatus_(item);

      if (!classId) {
        continue;
      }

      if (!V2_DBR_isActiveStudentStatus_(status)) {
        continue;
      }

      map[classId] = (map[classId] || 0) + 1;
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_DBR_buildAttendanceSummaryMap_(attendanceRows, days, studentClassMap) {
  try {
    var map = {};
    var startDate = V2_DBR_getDateDaysAgo_(days - 1);
    var endDate = new Date();
    studentClassMap = studentClassMap || {};

    for (var i = 0; i < attendanceRows.length; i++) {
      var item = attendanceRows[i];
      var classId = V2_DBR_resolveClassIdFromRow_(item, studentClassMap);
      var status = V2_DBR_getAttendanceStatus_(item);
      var dateValue = V2_DBR_getAttendanceDate_(item);

      if (!classId || !dateValue) {
        continue;
      }

      if (!V2_DBR_isDateInRange_(dateValue, startDate, endDate)) {
        continue;
      }

      if (!map[classId]) {
        map[classId] = {
          totalCount: 0,
          presentCount: 0,
          absentCount: 0,
          lateCount: 0,
          excusedCount: 0
        };
      }

      map[classId].totalCount += 1;

      if (V2_DBR_isPresentStatus_(status)) {
        map[classId].presentCount += 1;
      } else if (V2_DBR_isLateStatus_(status)) {
        map[classId].lateCount += 1;
      } else if (V2_DBR_isExcusedStatus_(status)) {
        map[classId].excusedCount += 1;
      } else {
        map[classId].absentCount += 1;
      }
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_DBR_buildScoreSummaryMap_(scoreRows, studentClassMap) {
  try {
    var map = {};
    studentClassMap = studentClassMap || {};

    for (var i = 0; i < scoreRows.length; i++) {
      var item = scoreRows[i];
      var classId = V2_DBR_resolveClassIdFromRow_(item, studentClassMap);
      var scoreValue = V2_DBR_getScoreValue_(item);

      if (!classId) {
        continue;
      }

      if (scoreValue === null) {
        continue;
      }

      if (!map[classId]) {
        map[classId] = {
          totalScore: 0,
          scoreCount: 0,
          averageScore: ''
        };
      }

      map[classId].totalScore += scoreValue;
      map[classId].scoreCount += 1;
    }

    for (var classId in map) {
      if (!Object.prototype.hasOwnProperty.call(map, classId)) {
        continue;
      }

      if (map[classId].scoreCount < 1) {
        map[classId].averageScore = '';
        continue;
      }

      map[classId].averageScore = V2_DBR_roundToOneDecimal_(
        map[classId].totalScore / map[classId].scoreCount
      );
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_DBR_resolveClassIdFromRow_(row, studentClassMap) {
  try {
    row = row || {};
    studentClassMap = studentClassMap || {};

    var directClassId = V2_DBR_getClassId_(row);
    if (directClassId) {
      return directClassId;
    }

    var studentId = V2_DBR_getStudentId_(row);
    if (studentId && studentClassMap[studentId]) {
      return studentClassMap[studentId];
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_DBR_getSheetObjectsByName_(ss, sheetName) {
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
        item[V2_DBR_toText_(header[j])] = values[i][j];
      }
      rows.push(item);
    }

    return rows;
  } catch (error) {
    V2_log_('ERROR', 'V2_DBR_getSheetObjectsByName_', error.message, error.stack || '');
    throw error;
  }
}

function V2_DBR_getStudentId_(row) {
  try {
    return V2_DBR_findValueByKeys_(row, ['studentId', 'student_id', '학생ID', '학생Id']);
  } catch (error) {
    return '';
  }
}

function V2_DBR_getClassId_(row) {
  try {
    return V2_DBR_findValueByKeys_(row, ['classId', 'class_id', '반ID', '반Id']);
  } catch (error) {
    return '';
  }
}

function V2_DBR_getClassName_(row) {
  try {
    return V2_DBR_findValueByKeys_(row, ['className', 'class_name', '반명', '반이름', 'name']);
  } catch (error) {
    return '';
  }
}

function V2_DBR_getSortOrder_(row) {
  try {
    var value = V2_DBR_findValueByKeys_(row, ['sortOrder', 'sort_order', '정렬순서']);
    return V2_DBR_toNumber_(value);
  } catch (error) {
    return 0;
  }
}

function V2_DBR_getClassTeacherEmail_(row) {
  try {
    row = row || {};

    var teacherEmail = V2_DBR_findValueByKeys_(row, ['teacherEmail', 'teacher_email', '담당교사이메일']);
    if (teacherEmail) {
      return teacherEmail;
    }

    var teacherId = V2_DBR_findValueByKeys_(row, ['teacherId', 'teacher_id', '담당교사ID']);
    if (teacherId) {
      var teacher = V2_DBR_getTeacherById_(teacherId);
      if (teacher && teacher.teacherEmail) {
        return teacher.teacherEmail;
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_DBR_getTeacherById_(teacherId) {
  try {
    var ss = V2_getSpreadsheet_();
    var teachers = V2_DBR_getSheetObjectsByName_(ss, 'V2_Teachers');
    var normalizedTeacherId = V2_DBR_toText_(teacherId);

    for (var i = 0; i < teachers.length; i++) {
      var item = teachers[i];
      var currentTeacherId = V2_DBR_findValueByKeys_(item, ['teacherId', 'teacher_id', '교사ID']);

      if (V2_DBR_isSameText_(currentTeacherId, normalizedTeacherId)) {
        return {
          teacherId: currentTeacherId,
          teacherName: V2_DBR_findValueByKeys_(item, ['teacherName', 'teacher_name', '교사명']),
          teacherEmail: V2_DBR_findValueByKeys_(item, ['teacherEmail', 'teacher_email', '교사이메일'])
        };
      }
    }

    return null;
  } catch (error) {
    return null;
  }
}

function V2_DBR_getStudentStatus_(row) {
  try {
    return V2_DBR_findValueByKeys_(row, ['status', 'studentStatus', '학생상태']);
  } catch (error) {
    return '';
  }
}

function V2_DBR_getAttendanceStatus_(row) {
  try {
    return V2_DBR_findValueByKeys_(row, ['status', 'attendanceStatus', '출석상태']);
  } catch (error) {
    return '';
  }
}

function V2_DBR_getAttendanceDate_(row) {
  try {
    return V2_DBR_findDateByKeys_(row, ['date', 'attendanceDate', '출석일', '날짜']);
  } catch (error) {
    return null;
  }
}

function V2_DBR_getScoreValue_(row) {
  try {
    var rawValue = V2_DBR_findValueByKeys_(row, ['score', '점수']);
    var numberValue = Number(rawValue);

    if (rawValue === '' || rawValue === null || rawValue === undefined || isNaN(numberValue)) {
      return null;
    }

    return numberValue;
  } catch (error) {
    return null;
  }
}

function V2_DBR_findValueByKeys_(row, keys) {
  try {
    row = row || {};
    keys = keys || [];

    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (Object.prototype.hasOwnProperty.call(row, key)) {
        return V2_DBR_toText_(row[key]);
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_DBR_findDateByKeys_(row, keys) {
  try {
    row = row || {};
    keys = keys || [];

    for (var i = 0; i < keys.length; i++) {
      var key = keys[i];
      if (!Object.prototype.hasOwnProperty.call(row, key)) {
        continue;
      }

      var dateValue = V2_DBR_toDate_(row[key]);
      if (dateValue) {
        return dateValue;
      }
    }

    return null;
  } catch (error) {
    return null;
  }
}

function V2_DBR_isActiveStudentStatus_(status) {
  try {
    return V2_DBR_toText_(status) === '재학';
  } catch (error) {
    return false;
  }
}

function V2_DBR_isPresentStatus_(status) {
  try {
    var text = V2_DBR_toText_(status);
    return text === '출석' || text === '재석' || text === 'present';
  } catch (error) {
    return false;
  }
}

function V2_DBR_isLateStatus_(status) {
  try {
    var text = V2_DBR_toText_(status);
    return text === '지각' || text === 'late';
  } catch (error) {
    return false;
  }
}

function V2_DBR_isExcusedStatus_(status) {
  try {
    var text = V2_DBR_toText_(status);
    return text === '공결' || text === '인정결석' || text === 'excused';
  } catch (error) {
    return false;
  }
}

function V2_DBR_calculateAttendanceRateText_(attendanceSummary) {
  try {
    if (!attendanceSummary || attendanceSummary.totalCount < 1) {
      return '0%';
    }

    var effectivePresentCount =
      V2_DBR_toNumber_(attendanceSummary.presentCount) +
      V2_DBR_toNumber_(attendanceSummary.lateCount) +
      V2_DBR_toNumber_(attendanceSummary.excusedCount);

    var rate = (effectivePresentCount / attendanceSummary.totalCount) * 100;
    return V2_DBR_roundToOneDecimal_(rate) + '%';
  } catch (error) {
    return '0%';
  }
}

function V2_DBR_getDateDaysAgo_(daysAgo) {
  try {
    var date = new Date();
    date.setHours(0, 0, 0, 0);
    date.setDate(date.getDate() - V2_DBR_toNumber_(daysAgo));
    return date;
  } catch (error) {
    return new Date();
  }
}

function V2_DBR_isDateInRange_(dateValue, startDate, endDate) {
  try {
    if (!dateValue || !startDate || !endDate) {
      return false;
    }

    var target = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
    var start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    var end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

    return target.getTime() >= start.getTime() && target.getTime() <= end.getTime();
  } catch (error) {
    return false;
  }
}

function V2_DBR_getClassNameFromMap_(classMap, classId) {
  try {
    if (classMap && classMap[classId] && classMap[classId].className) {
      return classMap[classId].className;
    }

    return classId;
  } catch (error) {
    return classId;
  }
}

function V2_DBR_getClassSortOrderFromMap_(classMap, classId) {
  try {
    if (classMap && classMap[classId]) {
      return V2_DBR_toNumber_(classMap[classId].sortOrder);
    }

    return 0;
  } catch (error) {
    return 0;
  }
}

function V2_DBR_sumNumberByKey_(items, key) {
  try {
    var sum = 0;

    for (var i = 0; i < items.length; i++) {
      sum += V2_DBR_toNumber_(items[i][key]);
    }

    return sum;
  } catch (error) {
    return 0;
  }
}

function V2_DBR_roundToOneDecimal_(value) {
  try {
    return Math.round(Number(value) * 10) / 10;
  } catch (error) {
    return 0;
  }
}

function V2_DBR_toDate_(value) {
  try {
    if (!value && value !== 0) {
      return null;
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return value;
    }

    var date = new Date(value);
    if (isNaN(date.getTime())) {
      return null;
    }

    return date;
  } catch (error) {
    return null;
  }
}

function V2_DBR_nowText_() {
  try {
    if (typeof V2_nowText_ === 'function') {
      return V2_nowText_();
    }

    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return '';
  }
}

function V2_DBR_getCurrentTeacherEmail_() {
  try {
    if (typeof V2_getCurrentUserEmail_ === 'function') {
      return V2_DBR_toText_(V2_getCurrentUserEmail_());
    }

    return V2_DBR_toText_(Session.getActiveUser().getEmail());
  } catch (error) {
    return '';
  }
}

function V2_DBR_getCurrentTeacherName_() {
  try {
    if (typeof V2_getCurrentUserName_ === 'function') {
      return V2_DBR_toText_(V2_getCurrentUserName_());
    }

    var email = V2_DBR_getCurrentTeacherEmail_();
    return email ? email.split('@')[0] : '';
  } catch (error) {
    return '';
  }
}

function V2_DBR_toDays_(value) {
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

function V2_DBR_toNumber_(value) {
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

function V2_DBR_isSameText_(a, b) {
  try {
    return V2_DBR_toText_(a).toLowerCase() === V2_DBR_toText_(b).toLowerCase();
  } catch (error) {
    return false;
  }
}

function V2_DBR_toText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    return String(value).trim();
  } catch (error) {
    return '';
  }
}
