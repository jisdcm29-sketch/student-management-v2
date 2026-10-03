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

function issueSession_(teacher) {
  const token = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
  const cacheKey = 'session:' + sha256Hex_(token);
  const payload = {
    teacherId: teacher.teacherId,
    displayName: teacher.displayName || '',
    role: teacher.role || 'TEACHER',
    issuedAt: new Date().toISOString()
  };

  CacheService.getScriptCache().put(
    cacheKey,
    JSON.stringify(payload),
    Number(V2_CONFIG.SESSION_TTL_SECONDS || 21600)
  );

  return {
    sessionToken: token,
    expiresInSeconds: Number(V2_CONFIG.SESSION_TTL_SECONDS || 21600),
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

  const cacheKey = 'session:' + sha256Hex_(token);
  const raw = CacheService.getScriptCache().get(cacheKey);
  if (!raw) {
    const error = new Error('세션이 만료되었거나 유효하지 않습니다.');
    error.code = 'SESSION_INVALID';
    throw error;
  }

  const session = JSON.parse(raw);
  const teacher = findTeacherById_(session.teacherId);

  if (!teacher || String(teacher.status || '').toUpperCase() !== 'ACTIVE') {
    CacheService.getScriptCache().remove(cacheKey);
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

  const cacheKey = 'session:' + sha256Hex_(token);
  const raw = CacheService.getScriptCache().get(cacheKey);
  if (raw) {
    try {
      const session = JSON.parse(raw);
      appendAuditLog_(session.teacherId || '', 'LOGOUT', 'Teacher', session.teacherId || '', 'SUCCESS', '');
    } catch (e) {}
  }
  CacheService.getScriptCache().remove(cacheKey);
  return { loggedOut: true };
}
