import { API_BASE_URL } from '../config.js';

async function callApi(action, payload = {}) {
  if (!API_BASE_URL) {
    throw new Error('API_BASE_URL이 설정되지 않았습니다. src/config.js에 V2 Apps Script 배포 URL을 넣어 주세요.');
  }

  const body = {
    action,
    ...payload
  };

  const response = await fetch(API_BASE_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'text/plain;charset=utf-8'
    },
    body: JSON.stringify(body),
    redirect: 'follow'
  });

  if (!response.ok) {
    throw new Error('API HTTP 오류: ' + response.status);
  }

  const result = await response.json();

  if (!result || result.ok !== true) {
    const message = result && result.error && result.error.message
      ? result.error.message
      : 'API 요청에 실패했습니다.';
    const error = new Error(message);
    error.code = result && result.error ? result.error.code : 'API_ERROR';
    throw error;
  }

  return result.data;
}

export async function health() {
  if (!API_BASE_URL) {
    throw new Error('API_BASE_URL이 설정되지 않았습니다.');
  }

  const url = API_BASE_URL + (API_BASE_URL.includes('?') ? '&' : '?') + 'action=health';
  const response = await fetch(url, { method: 'GET', redirect: 'follow' });

  if (!response.ok) {
    throw new Error('API HTTP 오류: ' + response.status);
  }

  const result = await response.json();
  if (!result || result.ok !== true) {
    throw new Error('Health check 실패');
  }
  return result.data;
}

export function login(teacherId, password) {
  return callApi('login', { teacherId, password });
}

export function logout(sessionToken) {
  return callApi('logout', { sessionToken });
}

export function validateSession(sessionToken) {
  return callApi('session', { sessionToken });
}

export function bootstrap(sessionToken) {
  return callApi('bootstrap', { sessionToken });
}

export function listClasses(sessionToken) {
  return callApi('classes.list', { sessionToken });
}
