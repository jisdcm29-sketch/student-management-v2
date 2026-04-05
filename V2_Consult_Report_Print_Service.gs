/**
 * V2_Consult_Report_Print_Service.gs
 * 상담 리포트 출력/PDF용 데이터 가공 Service
 * 역할:
 * - Repository 원본 데이터 수집
 * - 화면 출력/인쇄용 구조 생성
 * - 요약 정보 생성
 * - 브라우저로 안전하게 전달 가능한 값만 반환
 */

function V2_getConsultReportPrintData(request) {
  try {
    request = request || {};

    var studentId = V2_CRP_safeText_(request.studentId);
    if (!studentId) {
      throw new Error('studentId가 필요합니다.');
    }

    var options = {
      startDate: V2_CRP_safeText_(request.startDate),
      endDate: V2_CRP_safeText_(request.endDate),
      includeAttendance: request.includeAttendance !== false,
      includeScores: request.includeScores !== false,
      includeLessons: request.includeLessons !== false,
      includeComments: request.includeComments !== false,
      includeAbsenceContacts: request.includeAbsenceContacts !== false
    };

    var source = V2_CRP_getPrintSourceDataByStudentId_(studentId, options);
    var report = V2_CRP_buildPrintReportModel_(source);

    return {
      success: true,
      message: '상담 리포트 출력 데이터 생성 완료',
      data: report
    };
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_getConsultReportPrintData', errorMessage, error && error.stack ? error.stack : '');

    return {
      success: false,
      message: errorMessage,
      data: null
    };
  }
}

function V2_CRP_buildPrintReportModel_(source) {
  try {
    var student = source.student || {};
    var classInfo = source.classInfo || {};
    var teacherMap = source.teacherMap || {};

    var attendanceItems = source.options.includeAttendance
      ? V2_CRP_buildAttendancePrintItems_(source.attendanceRows || [])
      : [];

    var scoreItems = source.options.includeScores
      ? V2_CRP_buildScorePrintItems_(source.scoreRows || [])
      : [];

    var lessonItems = source.options.includeLessons
      ? V2_CRP_buildLessonPrintItems_(source.lessonRows || [], teacherMap)
      : [];

    var commentItems = source.options.includeComments
      ? V2_CRP_buildCommentPrintItems_(source.commentRows || [], teacherMap)
      : [];

    var absenceContactItems = source.options.includeAbsenceContacts
      ? V2_CRP_buildAbsenceContactPrintItems_(source.absenceContactRows || [], teacherMap)
      : [];

    var homeroomTeacherName = classInfo.teacherName || '';
    if (!homeroomTeacherName && classInfo.teacherId && teacherMap[classInfo.teacherId]) {
      homeroomTeacherName = teacherMap[classInfo.teacherId].teacherName || '';
    }

    return {
      reportTitle: '학부모 상담 리포트',
      generatedAt: V2_CRP_nowText_(),
      period: {
        startDate: source.options.startDate || '',
        endDate: source.options.endDate || '',
        label: V2_CRP_buildPeriodLabel_(source.options.startDate, source.options.endDate)
      },
      student: {
        studentId: V2_CRP_safeText_(student.studentId),
        studentName: V2_CRP_safeText_(student.studentName),
        classId: V2_CRP_safeText_(student.classId),
        className: V2_CRP_safeText_(classInfo.className || student.classId),
        status: V2_CRP_safeText_(student.status),
        parentName: V2_CRP_safeText_(student.parentName),
        parentPhone: V2_CRP_safeText_(student.parentPhone),
        memo: V2_CRP_safeText_(student.memo),
        homeroomTeacherName: V2_CRP_safeText_(homeroomTeacherName)
      },
      summary: V2_CRP_buildSummary_(attendanceItems, scoreItems, lessonItems, commentItems, absenceContactItems),
      sections: {
        attendance: attendanceItems,
        scores: scoreItems,
        lessons: lessonItems,
        comments: commentItems,
        absenceContacts: absenceContactItems
      }
    };
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildPrintReportModel_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_buildSummary_(attendanceItems, scoreItems, lessonItems, commentItems, absenceContactItems) {
  try {
    var attendanceSummary = V2_CRP_summarizeAttendance_(attendanceItems);
    var scoreSummary = V2_CRP_summarizeScores_(scoreItems);

    return {
      attendanceCount: attendanceItems.length,
      scoreCount: scoreItems.length,
      lessonCount: lessonItems.length,
      commentCount: commentItems.length,
      absenceContactCount: absenceContactItems.length,
      attendanceStatusCount: attendanceSummary,
      averageScore: scoreSummary.averageScore,
      scoreSummaryText: scoreSummary.summaryText
    };
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildSummary_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_buildAttendancePrintItems_(rows) {
  try {
    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];

      items.push({
        date: V2_CRP_pickDateText_(row, ['attendanceDate', 'date', 'lessonDate', 'createdAt']),
        status: V2_CRP_safeText_(row.status || row.attendanceStatus),
        note: V2_CRP_safeText_(row.note || row.memo),
        teacherId: V2_CRP_safeText_(row.teacherId),
        teacherName: V2_CRP_safeText_(row.teacherName)
      });
    }

    return items;
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildAttendancePrintItems_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_buildScorePrintItems_(rows) {
  try {
    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var scoreValue = V2_CRP_toNumber_(row.score || row.value || row.point || row.points);

      items.push({
        date: V2_CRP_pickDateText_(row, ['scoreDate', 'examDate', 'date', 'createdAt']),
        subject: V2_CRP_safeText_(row.subject || row.category),
        score: scoreValue,
        scoreText: scoreValue === null ? '' : String(scoreValue),
        note: V2_CRP_safeText_(row.note || row.memo),
        teacherId: V2_CRP_safeText_(row.teacherId),
        teacherName: V2_CRP_safeText_(row.teacherName)
      });
    }

    return items;
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildScorePrintItems_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_buildLessonPrintItems_(rows, teacherMap) {
  try {
    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var teacherName = V2_CRP_safeText_(row.teacherName);
      var teacherId = V2_CRP_safeText_(row.teacherId);

      if (!teacherName && teacherId && teacherMap[teacherId]) {
        teacherName = V2_CRP_safeText_(teacherMap[teacherId].teacherName);
      }

      items.push({
        date: V2_CRP_pickDateText_(row, ['lessonDate', 'date', 'createdAt']),
        subject: V2_CRP_safeText_(row.subject || row.lessonType || row.topic),
        content: V2_CRP_safeText_(row.content || row.lessonContent || row.topic),
        note: V2_CRP_safeText_(row.note || row.memo),
        teacherId: teacherId,
        teacherName: teacherName
      });
    }

    return items;
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildLessonPrintItems_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_buildCommentPrintItems_(rows, teacherMap) {
  try {
    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var teacherName = V2_CRP_safeText_(row.teacherName);
      var teacherId = V2_CRP_safeText_(row.teacherId);

      if (!teacherName && teacherId && teacherMap[teacherId]) {
        teacherName = V2_CRP_safeText_(teacherMap[teacherId].teacherName);
      }

      items.push({
        date: V2_CRP_pickDateText_(row, ['commentDate', 'date', 'createdAt']),
        category: V2_CRP_safeText_(row.category || row.commentType),
        comment: V2_CRP_safeText_(row.comment || row.content || row.memo),
        teacherId: teacherId,
        teacherName: teacherName
      });
    }

    return items;
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildCommentPrintItems_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_buildAbsenceContactPrintItems_(rows, teacherMap) {
  try {
    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      var teacherName = V2_CRP_safeText_(row.teacherName);
      var teacherId = V2_CRP_safeText_(row.teacherId);

      if (!teacherName && teacherId && teacherMap[teacherId]) {
        teacherName = V2_CRP_safeText_(teacherMap[teacherId].teacherName);
      }

      items.push({
        absenceDate: V2_CRP_pickDateText_(row, ['absenceDate', 'date', 'createdAt']),
        contactDate: V2_CRP_pickDateText_(row, ['contactDate', 'updatedAt', 'createdAt']),
        contactType: V2_CRP_safeText_(row.contactType || row.method || row.contactStatus),
        result: V2_CRP_safeText_(row.result || row.contactResult || row.contactStatus),
        memo: V2_CRP_safeText_(row.memo || row.note),
        teacherId: teacherId,
        teacherName: teacherName
      });
    }

    return items;
  } catch (error) {
    var errorMessage = V2_CRP_getErrorMessage_(error);
    V2_log_('ERROR', 'V2_CRP_buildAbsenceContactPrintItems_', errorMessage, error && error.stack ? error.stack : '');
    throw new Error(errorMessage);
  }
}

function V2_CRP_summarizeAttendance_(items) {
  try {
    var summary = {};

    for (var i = 0; i < items.length; i++) {
      var status = items[i].status || '미분류';
      if (!summary[status]) {
        summary[status] = 0;
      }
      summary[status]++;
    }

    return summary;
  } catch (error) {
    return {};
  }
}

function V2_CRP_summarizeScores_(items) {
  try {
    var total = 0;
    var count = 0;

    for (var i = 0; i < items.length; i++) {
      if (typeof items[i].score === 'number' && !isNaN(items[i].score)) {
        total += items[i].score;
        count++;
      }
    }

    var average = count > 0 ? Math.round((total / count) * 100) / 100 : null;

    return {
      averageScore: average,
      summaryText: count > 0 ? ('평균 ' + average + '점 / 총 ' + count + '건') : '성적 데이터 없음'
    };
  } catch (error) {
    return {
      averageScore: null,
      summaryText: '성적 데이터 없음'
    };
  }
}

function V2_CRP_buildPeriodLabel_(startDate, endDate) {
  try {
    if (startDate && endDate) {
      return startDate + ' ~ ' + endDate;
    }
    if (startDate) {
      return startDate + ' 이후';
    }
    if (endDate) {
      return endDate + ' 이전';
    }
    return '전체 기간';
  } catch (error) {
    return '전체 기간';
  }
}

function V2_CRP_pickDateText_(row, keys) {
  try {
    for (var i = 0; i < keys.length; i++) {
      var value = row[keys[i]];
      var formatted = V2_CRP_formatDate_(value);
      if (formatted) {
        return formatted;
      }
    }
    return '';
  } catch (error) {
    return '';
  }
}

function V2_CRP_formatDate_(value) {
  try {
    if (!value) {
      return '';
    }

    var dateObj;
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      dateObj = value;
    } else {
      var text = String(value).trim().replace(/\./g, '-').replace(/\//g, '-');
      dateObj = new Date(text);
    }

    if (isNaN(dateObj.getTime())) {
      return String(value);
    }

    return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch (error) {
    return String(value || '');
  }
}

function V2_CRP_toNumber_(value) {
  try {
    if (value === null || value === undefined || value === '') {
      return null;
    }

    var numberValue = Number(value);
    if (isNaN(numberValue)) {
      return null;
    }

    return numberValue;
  } catch (error) {
    return null;
  }
}

function V2_CRP_nowText_() {
  try {
    if (typeof V2_nowText_ === 'function') {
      return V2_nowText_();
    }

    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  } catch (error) {
    return '';
  }
}

function V2_CRP_safeText_(value) {
  try {
    if (value === null || value === undefined) {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    }

    return String(value);
  } catch (error) {
    return '';
  }
}

function V2_CRP_getErrorMessage_(error) {
  try {
    if (!error) {
      return '알 수 없는 오류가 발생했습니다.';
    }

    if (error.message) {
      return String(error.message);
    }

    return String(error);
  } catch (innerError) {
    return '오류 메시지를 읽을 수 없습니다.';
  }
}
