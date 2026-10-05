/**
 * Student Management V2 - Home Dashboard service
 *
 * Goal:
 * - Reproduce the teacher-friendly legacy Google dashboard using V2-local data.
 * - One API call returns operating classes, active students, cumulative attendance,
 *   latest lesson summaries and attention students.
 * - Mobile learning and legacy attendance use the already-synced V2 local sheets.
 * - The original legacy/mobile source spreadsheets are NOT read by this API.
 */

const HOME_DASHBOARD_V2_ = Object.freeze({
  CACHE_SECONDS: 120,
  ATTENDANCE_WARNING_RATE: 80,
  ATTENDANCE_MIN_RECORDS: 3,
  APP_INACTIVE_DAYS: 7,
  PROGRESS_DELAY_LESSONS: 2,
  LEGACY_ATTENDANCE_SHEET: 'LegacyAttendanceHistory',
  LEARNING_PROGRESS_SHEET: 'StudentLearningProgress'
});

function homeTextV2_(value) {
  return String(value == null ? '' : value).trim();
}

function homePhoneV2_(value) {
  return String(value || '').replace(/\D/g, '');
}

function homeDateV2_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  }
  const text = homeTextV2_(value);
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return '';
  return match[1] + '-' + String(match[2]).padStart(2, '0') + '-' + String(match[3]).padStart(2, '0');
}

function homeNumberV2_(value, fallback) {
  if (value === '' || value === null || value === undefined) return fallback;
  const n = Number(value);
  return Number.isFinite(n) ? n : fallback;
}

function homeLessonNumberV2_(value) {
  const text = homeTextV2_(value);
  const match = text.match(/\d+/);
  return match ? Number(match[0]) : null;
}


const HOME_SNU_PROGRESS_BOOKS_V2_ = Object.freeze([
  { id:'SNU-1A', lessonStart:1,  lessonEnd:8,  ordinalOffset:0 },
  { id:'SNU-1B', lessonStart:9,  lessonEnd:16, ordinalOffset:0 },
  { id:'SNU-2A', lessonStart:1,  lessonEnd:9,  ordinalOffset:16 },
  { id:'SNU-2B', lessonStart:10, lessonEnd:18, ordinalOffset:16 },
  { id:'SNU-3A', lessonStart:1,  lessonEnd:9,  ordinalOffset:34 },
  { id:'SNU-3B', lessonStart:10, lessonEnd:18, ordinalOffset:34 },
  { id:'SNU-4A', lessonStart:1,  lessonEnd:9,  ordinalOffset:52 },
  { id:'SNU-4B', lessonStart:10, lessonEnd:18, ordinalOffset:52 }
]);

function homeNormalizeSnuBookV2_(value) {
  const text = homeTextV2_(value).toUpperCase().replace(/\s+/g, '');
  if (!text) return '';
  const match = text.match(/(?:SNU[-_]?)?([1-4][AB])/);
  return match ? 'SNU-' + match[1] : '';
}

function homeSnuProgressOrdinalV2_(book, lesson) {
  const bookId = homeNormalizeSnuBookV2_(book);
  const lessonNo = homeLessonNumberV2_(lesson);
  if (!bookId || lessonNo === null) return null;
  const info = HOME_SNU_PROGRESS_BOOKS_V2_.find(function(item) { return item.id === bookId; });
  if (!info || lessonNo < info.lessonStart || lessonNo > info.lessonEnd) return null;
  return info.ordinalOffset + lessonNo;
}

function homeExplicitClassBookV2_(lesson) {
  if (!lesson) return '';
  // Only the current lesson topic is used as an explicit book hint.
  // This avoids accidentally treating a next-plan reference as the current book.
  return homeNormalizeSnuBookV2_(lesson.topic);
}

function homeClassProgressOrdinalV2_(lesson) {
  if (!lesson) return null;
  const raw = homeNumberV2_(lesson.progressCurrentCount, null);
  if (raw === null || raw <= 0) return null;

  // Legacy advanced-class records can store the lesson number within a book
  // (for example SNU-2A 6과 => 6), while newer records can store the cumulative
  // SNU ordinal. If the topic explicitly names a book, normalize that legacy value.
  const explicitBook = homeExplicitClassBookV2_(lesson);
  if (explicitBook) {
    const normalized = homeSnuProgressOrdinalV2_(explicitBook, raw);
    if (normalized !== null) return normalized;
  }

  // Otherwise the saved class progress is already the cumulative SNU ordinal.
  return raw;
}

function homeDaysSinceV2_(value) {
  const date = homeDateV2_(value);
  if (!date) return null;
  const todayText = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Ulaanbaatar', 'yyyy-MM-dd');
  const a = date.split('-').map(Number);
  const b = todayText.split('-').map(Number);
  const from = Date.UTC(a[0], a[1] - 1, a[2]);
  const to = Date.UTC(b[0], b[1] - 1, b[2]);
  return Math.max(0, Math.floor((to - from) / 86400000));
}

function homeReadOptionalSheetV2_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
  return readSheetObjects_(ss, sheetName);
}

function homeAttendanceStatusV2_(value) {
  const s = homeTextV2_(value);
  if (s === '결석') return 'absent';
  if (s === '지각') return 'late';
  if (s === '조퇴') return 'early';
  if (s === '출석') return 'present';
  return '';
}

function homeBuildAttendanceV2_(ss, activeStudentMap) {
  const byStudentDate = {};

  // Historical data already synchronized from the legacy system.
  homeReadOptionalSheetV2_(ss, HOME_DASHBOARD_V2_.LEGACY_ATTENDANCE_SHEET).forEach(function(row) {
    const studentId = homeTextV2_(row.studentId);
    const date = homeDateV2_(row.date);
    const status = homeAttendanceStatusV2_(row.status);
    if (!studentId || !date || !status || !activeStudentMap[studentId]) return;
    byStudentDate[studentId + '|' + date] = { studentId: studentId, date: date, status: status };
  });

  // Current V2 entries override the same student/date if both exist.
  homeReadOptionalSheetV2_(ss, 'Attendance').forEach(function(row) {
    const studentId = homeTextV2_(row.studentId);
    const date = homeDateV2_(row.date);
    const status = homeAttendanceStatusV2_(row.status);
    if (!studentId || !date || !status || !activeStudentMap[studentId]) return;
    byStudentDate[studentId + '|' + date] = { studentId: studentId, date: date, status: status };
  });

  const studentSummary = {};
  Object.keys(activeStudentMap).forEach(function(studentId) {
    studentSummary[studentId] = { total: 0, attended: 0, present: 0, late: 0, early: 0, absent: 0, rate: null };
  });

  Object.keys(byStudentDate).forEach(function(key) {
    const item = byStudentDate[key];
    const summary = studentSummary[item.studentId];
    if (!summary) return;
    summary.total++;
    if (item.status === 'absent') summary.absent++;
    else {
      summary.attended++;
      if (item.status === 'late') summary.late++;
      else if (item.status === 'early') summary.early++;
      else summary.present++;
    }
  });

  let total = 0;
  let attended = 0;
  Object.keys(studentSummary).forEach(function(studentId) {
    const summary = studentSummary[studentId];
    if (summary.total > 0) summary.rate = Math.round((summary.attended / summary.total) * 100);
    total += summary.total;
    attended += summary.attended;
  });

  return {
    overallRate: total > 0 ? Math.round((attended / total) * 100) : null,
    totalRecords: total,
    attendedRecords: attended,
    byStudent: studentSummary
  };
}

function homeBuildLearningProgressV2_(ss) {
  const byStudentId = {};
  const byPhone = {};
  homeReadOptionalSheetV2_(ss, HOME_DASHBOARD_V2_.LEARNING_PROGRESS_SHEET).forEach(function(row) {
    const studentId = homeTextV2_(row.studentId);
    const phone = homePhoneV2_(row.normalizedPhone);
    if (studentId) byStudentId[studentId] = row;
    if (phone) byPhone[phone] = row;
  });
  return { byStudentId: byStudentId, byPhone: byPhone };
}

function homeLatestLessonByClassV2_(auth, classes) {
  const out = {};
  (classes || []).forEach(function(classInfo) {
    const classId = homeTextV2_(classInfo.classId);
    if (!classId) return;
    try {
      const data = getLessonsListV2_(auth, { classId: classId, limit: 1 });
      out[classId] = data && data.rows && data.rows.length ? data.rows[0] : null;
    } catch (e) {
      out[classId] = null;
    }
  });
  return out;
}

function homeProgressDisplayV2_(lesson) {
  if (!lesson) return '';
  const current = homeNumberV2_(lesson.progressCurrentCount, null);
  if (current !== null && current > 0) {
    return String(current) + homeTextV2_(lesson.progressUnit || '');
  }
  return homeTextV2_(lesson.topic);
}

function homeAttentionItemV2_(student, classInfo, latestLesson, progressRow, attendanceSummary) {
  const labels = [];
  const reasons = [];
  const flags = { attendance: false, appInactive: false, progressDelay: false, linkage: false };

  if (!progressRow) {
    flags.linkage = true;
    labels.push('앱 연동');
    reasons.push(homePhoneV2_(student.phone) ? '모바일 진도 기록 없음' : '전화번호 없음');
  }

  if (attendanceSummary && attendanceSummary.total >= HOME_DASHBOARD_V2_.ATTENDANCE_MIN_RECORDS && attendanceSummary.rate < HOME_DASHBOARD_V2_.ATTENDANCE_WARNING_RATE) {
    flags.attendance = true;
    labels.push('출결');
    reasons.push('출석률 ' + attendanceSummary.rate + '%');
  }

  if (progressRow) {
    const recent = homeTextV2_(progressRow.recentActivityAt) || homeTextV2_(progressRow.recentLoginAt);
    const inactiveDays = homeDaysSinceV2_(recent);
    if (inactiveDays !== null && inactiveDays >= HOME_DASHBOARD_V2_.APP_INACTIVE_DAYS) {
      flags.appInactive = true;
      labels.push('앱 사용');
      reasons.push('최근 ' + inactiveDays + '일 미사용');
    }

    const classProgress = homeClassProgressOrdinalV2_(latestLesson);
    const studentProgress = homeSnuProgressOrdinalV2_(progressRow.currentBook, progressRow.currentLesson);
    if (classProgress !== null && classProgress > 0 && studentProgress !== null) {
      const gap = Math.floor(classProgress - studentProgress);
      if (gap >= HOME_DASHBOARD_V2_.PROGRESS_DELAY_LESSONS) {
        flags.progressDelay = true;
        labels.push('앱 진도');
        reasons.push('수업보다 ' + gap + '과 지연');
      }
    }
  }

  if (!labels.length) return null;
  return {
    studentId: homeTextV2_(student.studentId),
    name: homeTextV2_(student.name) || homeTextV2_(student.studentId),
    classId: homeTextV2_(student.classId),
    className: homeTextV2_(classInfo && classInfo.className) || homeTextV2_(student.classId),
    checkItem: labels.join(' · '),
    reason: reasons.join(' · '),
    flags: flags
  };
}

function homeBuildDashboardV2_(auth) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classes = readSheetObjects_(ss, 'Classes');
  const operatingClasses = classes.filter(function(row) {
    return homeTextV2_(row.status) === '운영중';
  });
  const operatingClassMap = {};
  operatingClasses.forEach(function(row) { operatingClassMap[homeTextV2_(row.classId)] = row; });

  const students = readSheetObjects_(ss, 'Students');
  const activeStudents = students.filter(function(row) {
    return homeTextV2_(row.status) === '재학' && !!operatingClassMap[homeTextV2_(row.classId)];
  });
  const activeStudentMap = {};
  activeStudents.forEach(function(row) { activeStudentMap[homeTextV2_(row.studentId)] = row; });

  const attendance = homeBuildAttendanceV2_(ss, activeStudentMap);
  const learning = homeBuildLearningProgressV2_(ss);
  const latestByClass = homeLatestLessonByClassV2_(auth, operatingClasses);

  const classSummaries = operatingClasses.map(function(classInfo) {
    const classId = homeTextV2_(classInfo.classId);
    const latest = latestByClass[classId] || null;
    return {
      classId: classId,
      className: homeTextV2_(classInfo.className) || classId,
      teacherName: homeTextV2_(classInfo.teacherName),
      activeStudentCount: activeStudents.filter(function(student) {
        return homeTextV2_(student.classId) === classId;
      }).length,
      recentProgress: homeProgressDisplayV2_(latest),
      recentLessonDate: latest ? homeDateV2_(latest.date) : '',
      recentTopic: latest ? homeTextV2_(latest.topic) : ''
    };
  }).sort(function(a, b) {
    return String(a.className).localeCompare(String(b.className), 'ko');
  });

  const attentionStudents = [];
  const breakdown = { attendance: 0, appInactive: 0, progressDelay: 0, linkage: 0 };

  activeStudents.forEach(function(student) {
    const classId = homeTextV2_(student.classId);
    const progressRow = learning.byStudentId[homeTextV2_(student.studentId)] || learning.byPhone[homePhoneV2_(student.phone)] || null;
    const item = homeAttentionItemV2_(
      student,
      operatingClassMap[classId],
      latestByClass[classId] || null,
      progressRow,
      attendance.byStudent[homeTextV2_(student.studentId)] || null
    );
    if (!item) return;
    attentionStudents.push(item);
    if (item.flags.attendance) breakdown.attendance++;
    if (item.flags.appInactive) breakdown.appInactive++;
    if (item.flags.progressDelay) breakdown.progressDelay++;
    if (item.flags.linkage) breakdown.linkage++;
  });

  attentionStudents.sort(function(a, b) {
    const byClass = String(a.className).localeCompare(String(b.className), 'ko');
    if (byClass !== 0) return byClass;
    return String(a.name).localeCompare(String(b.name), 'ko');
  });

  return {
    version: 'home-dashboard-v2-20261006-progress-fix1',
    generatedAt: new Date().toISOString(),
    stats: {
      operatingClassCount: operatingClasses.length,
      activeStudentCount: activeStudents.length,
      attendanceRate: attendance.overallRate,
      attentionCount: attentionStudents.length
    },
    attentionBreakdown: breakdown,
    classSummaries: classSummaries,
    attentionStudents: attentionStudents,
    criteria: {
      attendanceWarningRate: HOME_DASHBOARD_V2_.ATTENDANCE_WARNING_RATE,
      attendanceMinRecords: HOME_DASHBOARD_V2_.ATTENDANCE_MIN_RECORDS,
      appInactiveDays: HOME_DASHBOARD_V2_.APP_INACTIVE_DAYS,
      progressDelayLessons: HOME_DASHBOARD_V2_.PROGRESS_DELAY_LESSONS
    },
    diagnostics: {
      sourceMode: 'V2_LOCAL_SYNC_ONLY',
      activeStudentCount: activeStudents.length,
      attendanceRecords: attendance.totalRecords,
      learningProgressRows: Object.keys(learning.byStudentId).length,
      externalSpreadsheetRead: false
    }
  };
}

function getHomeDashboardV2_(auth, options) {
  options = options || {};
  const forceRefresh = options.forceRefresh === true;
  const teacherId = homeTextV2_(auth && auth.teacher && auth.teacher.teacherId) || 'teacher';
  const cache = CacheService.getScriptCache();
  const cacheKey = 'HOME_DASHBOARD_V2_' + teacherId;

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        const parsed = JSON.parse(cached);
        parsed.cached = true;
        return parsed;
      } catch (e) {}
    }
  }

  const data = homeBuildDashboardV2_(auth);
  data.cached = false;
  try {
    cache.put(cacheKey, JSON.stringify(data), HOME_DASHBOARD_V2_.CACHE_SECONDS);
  } catch (e) {}
  return data;
}
