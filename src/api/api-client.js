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

export const getQrAttendanceSetup = (sessionToken, options = {}) => callApi('qrAttendance.setup', { sessionToken, options });
export const startQrAttendanceSession = (sessionToken, sessionData) => callApi('qrAttendance.start', { sessionToken, sessionData });
export const closeQrAttendanceSession = (sessionToken, sessionData) => callApi('qrAttendance.close', { sessionToken, sessionData });
