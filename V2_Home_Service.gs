/**
 * V2_Home_Service.gs
 * [신규 전체 파일]
 * - 홈 대시보드 서비스
 * - 반별 현재 학생 수
 * - 최근 7일 출석률
 * - 반 평균 점수
 * - 로그인 교사 정보 조회
 */

function V2_HomeService_getDashboardData(request) {
  try {
    request = request || {};

    var timezone = Session.getScriptTimeZone();
    var today = new Date();
    var sevenDaysAgo = new Date(today);
    sevenDaysAgo.setDate(today.getDate() - 6);

    var classSheet = V2_HOME_getSheetIfExists_('V2_Classes');
    var studentSheet = V2_HOME_getSheetIfExists_('V2_Students');
    var attendanceSheet = V2_HOME_getSheetIfExists_('V2_Attendance');
    var scoreSheet = V2_HOME_getSheetIfExists_('V2_Scores');
    var teacherSheet = V2_HOME_getSheetIfExists_('V2_Teachers');

    var classRows = V2_HOME_getSheetObjects_(classSheet);
    var studentRows = V2_HOME_getSheetObjects_(studentSheet);
    var attendanceRows = V2_HOME_getSheetObjects_(attendanceSheet);
    var scoreRows = V2_HOME_getSheetObjects_(scoreSheet);
    var teacherRows = V2_HOME_getSheetObjects_(teacherSheet);

    var classMap = V2_HOME_buildClassMap_(classRows);
    var activeStudents = V2_HOME_filterActiveStudents_(studentRows);
    var classStudentCountMap = V2_HOME_buildClassStudentCountMap_(activeStudents);
    var classAttendanceMap = V2_HOME_buildRecentAttendanceRateMap_(attendanceRows, classMap, activeStudents, sevenDaysAgo, today, timezone);
    var classScoreMap = V2_HOME_buildClassAverageScoreMap_(scoreRows, classMap);

    var dashboardItems = V2_HOME_buildDashboardItems_({
      classMap: classMap,
      classStudentCountMap: classStudentCountMap,
      classAttendanceMap: classAttendanceMap,
      classScoreMap: classScoreMap
    });

    return {
      success: true,
      message: '홈 대시보드 데이터를 불러왔습니다.',
      data: {
        teacherInfo: V2_HOME_getCurrentTeacherInfo_(teacherRows),
        summary: {
          totalClassCount: Object.keys(classMap).length,
          totalActiveStudentCount: activeStudents.length,
          generatedAt: Utilities.formatDate(today, timezone, 'yyyy-MM-dd HH:mm:ss')
        },
        items: dashboardItems
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '홈 대시보드 조회 중 오류가 발생했습니다.',
      data: {
        teacherInfo: {
          teacherId: '',
          teacherName: '',
          teacherEmail: ''
        },
        summary: {
          totalClassCount: 0,
          totalActiveStudentCount: 0,
          generatedAt: ''
        },
        items: []
      }
    };
  }
}

function V2_HomeService_getTeacherPanelInfo() {
  try {
    var teacherSheet = V2_HOME_getSheetIfExists_('V2_Teachers');
    var teacherRows = V2_HOME_getSheetObjects_(teacherSheet);

    return {
      success: true,
      message: '교사 정보를 불러왔습니다.',
      data: V2_HOME_getCurrentTeacherInfo_(teacherRows)
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '교사 정보 조회 중 오류가 발생했습니다.',
      data: {
        teacherId: '',
        teacherName: '',
        teacherEmail: ''
      }
    };
  }
}

function V2_HOME_getCurrentTeacherInfo_(teacherRows) {
  try {
    teacherRows = Array.isArray(teacherRows) ? teacherRows : [];

    var email = '';
    try {
      email = Session.getActiveUser().getEmail() || '';
    } catch (innerError) {
      email = '';
    }

    var teacher = null;
    for (var i = 0; i < teacherRows.length; i++) {
      var row = teacherRows[i] || {};
      var teacherEmail = V2_HOME_pickValue_(row, ['teacheremail', 'email', '교사이메일', '이메일']);
      if (String(teacherEmail || '').trim().toLowerCase() === String(email || '').trim().toLowerCase()) {
        teacher = row;
        break;
      }
    }

    return {
      teacherId: teacher ? String(V2_HOME_pickValue_(teacher, ['teacherid', 'id', '교사id', '교사번호']) || '') : '',
      teacherName: teacher ? String(V2_HOME_pickValue_(teacher, ['teachername', 'name', '교사명']) || '') : '',
      teacherEmail: email || ''
    };
  } catch (error) {
    return {
      teacherId: '',
      teacherName: '',
      teacherEmail: ''
    };
  }
}

function V2_HOME_buildDashboardItems_(context) {
  try {
    context = context || {};

    var classMap = context.classMap || {};
    var classStudentCountMap = context.classStudentCountMap || {};
    var classAttendanceMap = context.classAttendanceMap || {};
    var classScoreMap = context.classScoreMap || {};
    var classIds = Object.keys(classMap);
    var items = [];

    for (var i = 0; i < classIds.length; i++) {
      var classId = classIds[i];
      var classInfo = classMap[classId] || {};

      items.push({
        classId: classId,
        className: classInfo.className || classId,
        currentStudentCount: Number(classStudentCountMap[classId] || 0),
        recent7DayAttendanceRate: Number(classAttendanceMap[classId] || 0),
        averageScore: Number(classScoreMap[classId] || 0)
      });
    }

    items.sort(function(a, b) {
      return String(a.className || '').localeCompare(String(b.className || ''), 'ko');
    });

    return items;
  } catch (error) {
    return [];
  }
}

function V2_HOME_buildClassMap_(classRows) {
  try {
    classRows = Array.isArray(classRows) ? classRows : [];

    var map = {};

    for (var i = 0; i < classRows.length; i++) {
      var row = classRows[i] || {};
      var classId = String(V2_HOME_pickValue_(row, ['classid', 'id', '반id', '반코드']) || '').trim();
      var className = String(V2_HOME_pickValue_(row, ['classname', 'name', '반명']) || '').trim();

      if (!classId && !className) {
        continue;
      }

      if (!classId) {
        classId = className;
      }

      map[classId] = {
        classId: classId,
        className: className || classId
      };
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_HOME_filterActiveStudents_(studentRows) {
  try {
    studentRows = Array.isArray(studentRows) ? studentRows : [];

    var activeStudents = [];

    for (var i = 0; i < studentRows.length; i++) {
      var row = studentRows[i] || {};
      var studentId = String(V2_HOME_pickValue_(row, ['studentid', 'id', '학생id', '학생번호']) || '').trim();
      var status = String(V2_HOME_pickValue_(row, ['studentstatus', 'status', '학생상태']) || '').trim();

      if (!studentId) {
        continue;
      }

      if (!status || status === '재학') {
        activeStudents.push(row);
      }
    }

    return activeStudents;
  } catch (error) {
    return [];
  }
}

function V2_HOME_buildClassStudentCountMap_(activeStudents) {
  try {
    activeStudents = Array.isArray(activeStudents) ? activeStudents : [];

    var map = {};

    for (var i = 0; i < activeStudents.length; i++) {
      var row = activeStudents[i] || {};
      var classId = String(V2_HOME_pickValue_(row, ['classid', '반id', 'currentclassid', '소속반id']) || '').trim();

      if (!classId) {
        continue;
      }

      if (!map[classId]) {
        map[classId] = 0;
      }

      map[classId] += 1;
    }

    return map;
  } catch (error) {
    return {};
  }
}

function V2_HOME_buildRecentAttendanceRateMap_(attendanceRows, classMap, activeStudents, fromDate, toDate, timezone) {
  try {
    attendanceRows = Array.isArray(attendanceRows) ? attendanceRows : [];
    classMap = classMap || {};
    activeStudents = Array.isArray(activeStudents) ? activeStudents : [];

    var studentClassMap = {};
    var activeStudentIdSet = {};

    for (var i = 0; i < activeStudents.length; i++) {
      var student = activeStudents[i] || {};
      var studentId = String(V2_HOME_pickValue_(student, ['studentid', 'id', '학생id', '학생번호']) || '').trim();
      var classId = String(V2_HOME_pickValue_(student, ['classid', '반id', 'currentclassid', '소속반id']) || '').trim();

      if (studentId) {
        activeStudentIdSet[studentId] = true;
        studentClassMap[studentId] = classId;
      }
    }

    var countedMap = {};
    var attendedMap = {};
    var classIds = Object.keys(classMap);

    for (var c = 0; c < classIds.length; c++) {
      countedMap[classIds[c]] = 0;
      attendedMap[classIds[c]] = 0;
    }

    for (var j = 0; j < attendanceRows.length; j++) {
      var row = attendanceRows[j] || {};
      var studentIdInAttendance = String(V2_HOME_pickValue_(row, ['studentid', '학생id', '학생번호']) || '').trim();
      var classIdInAttendance = String(V2_HOME_pickValue_(row, ['classid', '반id']) || '').trim();
      var status = String(V2_HOME_pickValue_(row, ['status', 'attendancestatus', '출석상태']) || '').trim();
      var dateValue = V2_HOME_pickValue_(row, ['date', 'attendancedate', '출석일', '날짜']);
      var normalizedDate = V2_HOME_normalizeDateValue_(dateValue, timezone);

      if (!studentIdInAttendance || !activeStudentIdSet[studentIdInAttendance]) {
        continue;
      }

      if (!normalizedDate) {
        continue;
      }

      if (normalizedDate < V2_HOME_formatDateKey_(fromDate, timezone) || normalizedDate > V2_HOME_formatDateKey_(toDate, timezone)) {
        continue;
      }

      if (!classIdInAttendance) {
        classIdInAttendance = studentClassMap[studentIdInAttendance] || '';
      }

      if (!classIdInAttendance) {
        continue;
      }

      countedMap[classIdInAttendance] = Number(countedMap[classIdInAttendance] || 0) + 1;

      if (status === '출석' || status === '재석' || status === '지각' || status === '공결' || status === '인정결석') {
        attendedMap[classIdInAttendance] = Number(attendedMap[classIdInAttendance] || 0) + 1;
      }
    }

    var rateMap = {};
    for (var k = 0; k < classIds.length; k++) {
      var classId = classIds[k];
      var totalCount = Number(countedMap[classId] || 0);
      var attendedCount = Number(attendedMap[classId] || 0);

      if (totalCount < 1) {
        rateMap[classId] = 0;
      } else {
        rateMap[classId] = Math.round((attendedCount / totalCount) * 1000) / 10;
      }
    }

    return rateMap;
  } catch (error) {
    return {};
  }
}

function V2_HOME_buildClassAverageScoreMap_(scoreRows, classMap) {
  try {
    scoreRows = Array.isArray(scoreRows) ? scoreRows : [];
    classMap = classMap || {};

    var totalMap = {};
    var countMap = {};
    var classIds = Object.keys(classMap);

    for (var i = 0; i < classIds.length; i++) {
      totalMap[classIds[i]] = 0;
      countMap[classIds[i]] = 0;
    }

    for (var j = 0; j < scoreRows.length; j++) {
      var row = scoreRows[j] || {};
      var classId = String(V2_HOME_pickValue_(row, ['classid', '반id']) || '').trim();
      var scoreValue = V2_HOME_pickValue_(row, ['score', 'totalscore', '점수', '총점']);

      if (!classId) {
        continue;
      }

      var numericScore = Number(scoreValue);
      if (isNaN(numericScore)) {
        continue;
      }

      totalMap[classId] = Number(totalMap[classId] || 0) + numericScore;
      countMap[classId] = Number(countMap[classId] || 0) + 1;
    }

    var averageMap = {};
    for (var k = 0; k < classIds.length; k++) {
      var targetClassId = classIds[k];
      var count = Number(countMap[targetClassId] || 0);

      if (count < 1) {
        averageMap[targetClassId] = 0;
      } else {
        averageMap[targetClassId] = Math.round((Number(totalMap[targetClassId] || 0) / count) * 10) / 10;
      }
    }

    return averageMap;
  } catch (error) {
    return {};
  }
}

function V2_HOME_getSheetObjects_(sheet) {
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
      var item = {};

      for (var j = 0; j < headers.length; j++) {
        var rawHeader = headers[j];
        var normalizedHeader = V2_HOME_normalizeHeader_(rawHeader);

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

function V2_HOME_getSheetIfExists_(sheetName) {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    return spreadsheet.getSheetByName(sheetName);
  } catch (error) {
    return null;
  }
}

function V2_HOME_pickValue_(row, aliases) {
  try {
    row = row || {};
    aliases = Array.isArray(aliases) ? aliases : [];

    for (var i = 0; i < aliases.length; i++) {
      var key = V2_HOME_normalizeHeader_(aliases[i]);
      if (Object.prototype.hasOwnProperty.call(row, key)) {
        return row[key];
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_HOME_normalizeHeader_(value) {
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

function V2_HOME_normalizeDateValue_(value, timezone) {
  try {
    if (!value) {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
    }

    var text = String(value).trim();
    if (!text) {
      return '';
    }

    text = text.replace(/\./g, '-').replace(/\//g, '-').replace(/\s+/g, '');
    var match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);

    if (match) {
      return [
        match[1],
        ('0' + match[2]).slice(-2),
        ('0' + match[3]).slice(-2)
      ].join('-');
    }

    var parsed = new Date(text);
    if (!isNaN(parsed.getTime())) {
      return Utilities.formatDate(parsed, timezone, 'yyyy-MM-dd');
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_HOME_formatDateKey_(date, timezone) {
  try {
    if (!date || isNaN(date.getTime())) {
      return '';
    }

    return Utilities.formatDate(date, timezone, 'yyyy-MM-dd');
  } catch (error) {
    return '';
  }
}
