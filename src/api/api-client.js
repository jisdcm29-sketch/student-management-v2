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
