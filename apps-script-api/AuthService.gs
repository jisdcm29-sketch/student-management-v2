function bytesToHex_(bytes) {
  return bytes.map(function(b) {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function sha256Hex_(text) {
  return bytesToHex_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(text),
    Utilities.Charset.UTF_8
  ));
}

function hashPassword_(password, salt) {
  let value = String(salt) + '|' + String(password);
  const rounds = Number(V2_CONFIG.PASSWORD_HASH_ROUNDS || 1200);

  for (let i = 0; i < rounds; i++) {
    value = sha256Hex_(value + '|' + salt + '|' + i);
  }
  return value;
}

function setupInitialAdminPassword() {
  return setupInitialAdminPasswordFromProperty_();
}

function setupInitialAdminPasswordFromProperty_() {
  const props = PropertiesService.getScriptProperties();
  const password = String(props.getProperty('SETUP_ADMIN_PASSWORD') || '');

  if (password.length < 10) {
    throw new Error('SETUP_ADMIN_PASSWORD를 10자 이상으로 설정한 뒤 다시 실행해 주세요.');
  }

  const teacher = findTeacherById_(V2_CONFIG.ADMIN_TEACHER_ID);
  if (!teacher) {
    throw new Error('초기 관리자 교사 행을 찾을 수 없습니다.');
  }

  const salt = Utilities.getUuid() + Utilities.getUuid();
  const verifier = hashPassword_(password, salt);
  updateTeacherAuthFields_(V2_CONFIG.ADMIN_TEACHER_ID, salt, verifier);

  props.deleteProperty('SETUP_ADMIN_PASSWORD');
  appendAuditLog_(V2_CONFIG.ADMIN_TEACHER_ID, 'ADMIN_PASSWORD_SETUP', 'Teacher', V2_CONFIG.ADMIN_TEACHER_ID, 'SUCCESS', 'Initial V2 admin password configured.');

  return 'V2 관리자 비밀번호 설정 완료';
}

function sessionStorageKeyV2_(token) {
  return 'V2_SESSION_' + sha256Hex_(String(token || '').trim());
}

function sessionCacheKeyV2_(token) {
  return 'session:' + sha256Hex_(String(token || '').trim());
}

function sessionTtlSecondsV2_() {
  return Math.max(300, Number(V2_CONFIG.SESSION_TTL_SECONDS || 21600));
}

function removePersistentSessionV2_(token) {
  const normalized = String(token || '').trim();
  if (!normalized) return;
  CacheService.getScriptCache().remove(sessionCacheKeyV2_(normalized));
  PropertiesService.getScriptProperties().deleteProperty(sessionStorageKeyV2_(normalized));
}

function issueSession_(teacher) {
  const token = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
  const ttlSeconds = sessionTtlSecondsV2_();
  const issuedAt = new Date();
  const expiresAt = new Date(issuedAt.getTime() + ttlSeconds * 1000);
  const payload = {
    teacherId: teacher.teacherId,
    displayName: teacher.displayName || '',
    role: teacher.role || 'TEACHER',
    issuedAt: issuedAt.toISOString(),
    expiresAt: expiresAt.toISOString()
  };

  const raw = JSON.stringify(payload);
  CacheService.getScriptCache().put(
    sessionCacheKeyV2_(token),
    raw,
    Math.min(21600, ttlSeconds)
  );

  // CacheService는 만료 전에도 제거될 수 있으므로 새로고침 복원을 위해
  // 동일 세션을 Script Properties에도 보존한다. 토큰 원문은 저장하지 않고
  // SHA-256 해시를 키로 사용한다.
  PropertiesService.getScriptProperties().setProperty(
    sessionStorageKeyV2_(token),
    raw
  );

  return {
    sessionToken: token,
    expiresInSeconds: ttlSeconds,
    teacher: {
      teacherId: payload.teacherId,
      displayName: payload.displayName,
      role: payload.role
    }
  };
}

function requireSession_(sessionToken) {
  const token = String(sessionToken || '').trim();
  if (!token) {
    const error = new Error('로그인이 필요합니다.');
    error.code = 'AUTH_REQUIRED';
    throw error;
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = sessionCacheKeyV2_(token);
  const propertyKey = sessionStorageKeyV2_(token);
  const props = PropertiesService.getScriptProperties();

  let raw = cache.get(cacheKey);
  let fromPersistentStore = false;

  if (!raw) {
    raw = props.getProperty(propertyKey);
    fromPersistentStore = !!raw;
  }

  if (!raw) {
    const error = new Error('세션이 만료되었거나 유효하지 않습니다.');
    error.code = 'SESSION_INVALID';
    throw error;
  }

  let session;
  try {
    session = JSON.parse(raw);
  } catch (e) {
    cache.remove(cacheKey);
    props.deleteProperty(propertyKey);
    const error = new Error('세션 정보가 손상되었습니다. 다시 로그인해 주세요.');
    error.code = 'SESSION_INVALID';
    throw error;
  }

  const expiresAtMs = Date.parse(String(session.expiresAt || ''));
  if (!Number.isFinite(expiresAtMs) || expiresAtMs <= Date.now()) {
    cache.remove(cacheKey);
    props.deleteProperty(propertyKey);
    const error = new Error('세션이 만료되었습니다. 다시 로그인해 주세요.');
    error.code = 'SESSION_INVALID';
    throw error;
  }

  if (fromPersistentStore) {
    const remainingSeconds = Math.max(1, Math.floor((expiresAtMs - Date.now()) / 1000));
    cache.put(cacheKey, raw, Math.min(21600, remainingSeconds));
  }

  const teacher = findTeacherById_(session.teacherId);

  if (!teacher || String(teacher.status || '').toUpperCase() !== 'ACTIVE') {
    cache.remove(cacheKey);
    props.deleteProperty(propertyKey);
    const error = new Error('사용할 수 없는 교사 계정입니다.');
    error.code = 'ACCOUNT_INACTIVE';
    throw error;
  }

  return {
    token: token,
    cacheKey: cacheKey,
    session: session,
    teacher: teacher
  };
}

function loginTeacher_(teacherId, password) {
  const teacher = findTeacherById_(teacherId);

  if (!teacher || String(teacher.status || '').toUpperCase() !== 'ACTIVE') {
    const error = new Error('교사 ID 또는 비밀번호를 확인해 주세요.');
    error.code = 'LOGIN_FAILED';
    throw error;
  }

  const salt = String(teacher.passwordSalt || '');
  const verifier = String(teacher.passwordVerifier || '');

  if (!salt || !verifier) {
    const error = new Error('V2 초기 비밀번호 설정이 필요합니다.');
    error.code = 'PASSWORD_SETUP_REQUIRED';
    throw error;
  }

  const incoming = hashPassword_(String(password || ''), salt);
  if (incoming !== verifier) {
    appendAuditLog_(teacher.teacherId, 'LOGIN', 'Teacher', teacher.teacherId, 'FAILED', 'Invalid credentials');
    const error = new Error('교사 ID 또는 비밀번호를 확인해 주세요.');
    error.code = 'LOGIN_FAILED';
    throw error;
  }

  appendAuditLog_(teacher.teacherId, 'LOGIN', 'Teacher', teacher.teacherId, 'SUCCESS', '');
  return issueSession_(teacher);
}

function logoutTeacher_(sessionToken) {
  const token = String(sessionToken || '').trim();
  if (!token) return { loggedOut: true };

  const cacheKey = sessionCacheKeyV2_(token);
  const propertyKey = sessionStorageKeyV2_(token);
  const cache = CacheService.getScriptCache();
  const props = PropertiesService.getScriptProperties();

  const raw = cache.get(cacheKey) || props.getProperty(propertyKey);
  if (raw) {
    try {
      const session = JSON.parse(raw);
      appendAuditLog_(session.teacherId || '', 'LOGOUT', 'Teacher', session.teacherId || '', 'SUCCESS', '');
    } catch (e) {}
  }

  cache.remove(cacheKey);
  props.deleteProperty(propertyKey);
  return { loggedOut: true };
}
