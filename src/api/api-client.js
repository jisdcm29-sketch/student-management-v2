import { API_BASE_URL } from '../config.js';

async function callApi(action, payload = {}) {
  if (!API_BASE_URL) throw new Error('API_BASE_URL이 설정되지 않았습니다.');
  const response = await fetch(API_BASE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload }),
    redirect: 'follow'
  });
  if (!response.ok) throw new Error('API HTTP 오류: ' + response.status);
  const result = await response.json();
  if (!result || result.ok !== true) {
    const error = new Error(result && result.error && result.error.message ? result.error.message : 'API 요청에 실패했습니다.');
    error.code = result && result.error ? result.error.code : 'API_ERROR';
    throw error;
  }
  return result.data;
}

export async function health() {
  const url = API_BASE_URL + (API_BASE_URL.includes('?') ? '&' : '?') + 'action=health';
  const response = await fetch(url, { method: 'GET', redirect: 'follow' });
  if (!response.ok) throw new Error('API HTTP 오류: ' + response.status);
  const result = await response.json();
  if (!result || result.ok !== true) throw new Error('Health check 실패');
  return result.data;
}

export const login = (teacherId, password) => callApi('login', { teacherId, password });
export const logout = sessionToken => callApi('logout', { sessionToken });
export const validateSession = sessionToken => callApi('session', { sessionToken });
export const bootstrap = sessionToken => callApi('bootstrap', { sessionToken });
export const getHomeDashboard = (sessionToken, options = {}) => callApi('home.dashboard', { sessionToken, options });
export const listClasses = (sessionToken, options = {}) => callApi('classes.list', { sessionToken, options });
export const getClassDetail = (sessionToken, classId) => callApi('classes.get', { sessionToken, classId });
export const saveClass = (sessionToken, classData) => callApi('classes.save', { sessionToken, classData });
export const closeClass = (sessionToken, classId) => callApi('classes.close', { sessionToken, classId });
export const deleteClass = (sessionToken, classId) => callApi('classes.delete', { sessionToken, classId });

export const listStudents = (sessionToken, options = {}) => callApi('students.list', { sessionToken, options });
export const getStudentDetail = (sessionToken, studentId) => callApi('students.get', { sessionToken, studentId });
export const saveStudent = (sessionToken, studentData) => callApi('students.save', { sessionToken, studentData });

export const listArchivedStudents = (sessionToken, options = {}) => callApi('students.archived.list', { sessionToken, options });
export const restoreStudent = (sessionToken, studentId) => callApi('students.restore', { sessionToken, studentId });

export const listDailyRecords = (sessionToken, options = {}) => callApi('dailyRecords.list', { sessionToken, options });
export const getDailyRecord = (sessionToken, recordId) => callApi('dailyRecords.get', { sessionToken, recordId });
export const saveDailyRecord = (sessionToken, recordData) => callApi('dailyRecords.save', { sessionToken, recordData });
export const deleteDailyRecord = (sessionToken, recordId) => callApi('dailyRecords.delete', { sessionToken, recordId });

export const loadAttendance = (sessionToken, options = {}) => callApi('attendance.load', { sessionToken, options });
export const saveAttendance = (sessionToken, attendanceData) => callApi('attendance.save', { sessionToken, attendanceData });
export const saveAttendanceHoliday = (sessionToken, holidayData) => callApi('attendance.holiday.save', { sessionToken, holidayData });
export const clearAttendanceHoliday = (sessionToken, holidayData) => callApi('attendance.holiday.clear', { sessionToken, holidayData });

export const getScoreHubSources = sessionToken => callApi('scoreHub.sources', { sessionToken });
export const getScoreHubStudentSummary = (sessionToken, studentId) => callApi('scoreHub.studentSummary', { sessionToken, studentId });
export const listTeacherManualScores = (sessionToken, studentId) => callApi('scores.manual.list', { sessionToken, studentId });
export const saveTeacherManualScore = (sessionToken, scoreData) => callApi('scores.manual.save', { sessionToken, scoreData });
export const deleteTeacherManualScore = (sessionToken, recordId) => callApi('scores.manual.delete', { sessionToken, recordId });

export const getStudentAttendanceHistory = (sessionToken, studentId, options = {}) => callApi('report.attendance.student', { sessionToken, studentId, options });
export const getStudentSummaryReport = (sessionToken, studentId, options = {}) => callApi('report.student.summary', { sessionToken, studentId, options });
export const getClassSummaryReport = (sessionToken, classId, options = {}) => callApi('report.class.summary', { sessionToken, classId, options });
export const getStudentStatusStatsReport = (sessionToken, options = {}) => callApi('report.student.statusStats', { sessionToken, options });
export const getReportEmailContact = (sessionToken, studentId) => callApi('report.email.contact', { sessionToken, studentId });
export const prepareReportEmail = (sessionToken, emailData = {}) => callApi('report.email.prepare', { sessionToken, emailData });

export const listLessons = (sessionToken, options = {}) => callApi('lessons.list', { sessionToken, options });
export const getPreviousLesson = (sessionToken, classId, beforeDate = '') => callApi('lessons.previous', { sessionToken, classId, beforeDate });
export const getLesson = (sessionToken, lessonId) => callApi('lessons.get', { sessionToken, lessonId });
export const saveLesson = (sessionToken, lessonData) => callApi('lessons.save', { sessionToken, lessonData });
export const listLessonAssignments = (sessionToken, lessonId) => callApi('lessons.assignments.list', { sessionToken, lessonId });
export const saveLessonAssignments = (sessionToken, assignmentData) => callApi('lessons.assignments.save', { sessionToken, assignmentData });
export const listArchivedLessons = (sessionToken, options = {}) => callApi('lessons.archive.list', { sessionToken, options });
export const archiveLesson = (sessionToken, lessonId) => callApi('lessons.archive', { sessionToken, lessonId });
export const restoreArchivedLesson = (sessionToken, lessonId) => callApi('lessons.restore', { sessionToken, lessonId });
export const previewLessonMigrationStep7C = sessionToken => callApi('lessons.migration.preview', { sessionToken });
export const executeLessonMigrationStep7C = (sessionToken, confirmText) => callApi('lessons.migration.execute', { sessionToken, confirmText });
export const previewLessonTestCleanupStep7E = sessionToken => callApi('lessons.migration.cleanup.preview', { sessionToken });
export const executeLessonTestCleanupStep7E = (sessionToken, confirmText) => callApi('lessons.migration.cleanup.execute', { sessionToken, confirmText });

export const getQrAttendanceSetup = (sessionToken, options = {}) => callApi('qrAttendance.setup', { sessionToken, options });
export const startQrAttendanceSession = (sessionToken, sessionData) => callApi('qrAttendance.start', { sessionToken, sessionData });
export const closeQrAttendanceSession = (sessionToken, sessionData) => callApi('qrAttendance.close', { sessionToken, sessionData });
