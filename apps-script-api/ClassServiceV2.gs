function normalizeClassStatusV2_(status) {
  const s = String(status || '').trim();
  const lower = s.toLowerCase ? s.toLowerCase() : s;
  if (!s) return '운영중';
  if (s === '종료' || s === '종결' || s === '폐강' || s === '비활성' || lower === 'closed' || lower === 'inactive') return '종료';
  if (s === '예정') return '예정';
  return '운영중';
}

function requireSuperAdminV2_(auth) {
  if (!auth || !auth.teacher || String(auth.teacher.role || '').toUpperCase() !== 'SUPER_ADMIN') {
    const error = new Error('관리자 권한이 필요합니다.');
    error.code = 'FORBIDDEN';
    throw error;
  }
}

function getSheetHeaderMapV2_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) return {};
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  const map = {};
  headers.forEach(function(header, index) {
    const key = String(header || '').trim();
    if (key) map[key] = index;
  });
  return map;
}

function getRequiredDataSheetV2_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    const error = new Error(name + ' 시트를 찾을 수 없습니다.');
    error.code = 'SHEET_NOT_FOUND';
    throw error;
  }
  return sheet;
}

function makeNextPrefixedIdV2_(sheet, idHeader, prefix, digits) {
  const headers = getSheetHeaderMapV2_(sheet);
  if (headers[idHeader] == null) throw new Error(sheet.getName() + ' 시트에 ' + idHeader + ' 열이 없습니다.');
  const lastRow = sheet.getLastRow();
  let max = 0;
  if (lastRow >= 2) {
    const values = sheet.getRange(2, headers[idHeader] + 1, lastRow - 1, 1).getDisplayValues();
    values.forEach(function(row) {
      const id = String(row[0] || '').trim();
      if (id.indexOf(prefix) !== 0) return;
      const tail = id.slice(prefix.length);
      if (!/^\d+$/.test(tail)) return;
      max = Math.max(max, Number(tail || 0));
    });
  }
  return prefix + String(max + 1).padStart(digits || 3, '0');
}

function findDataRowByIdV2_(sheet, idHeader, idValue) {
  const headers = getSheetHeaderMapV2_(sheet);
  if (headers[idHeader] == null) return -1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, headers[idHeader] + 1, lastRow - 1, 1).getDisplayValues();
  const target = String(idValue || '').trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) return i + 2;
  }
  return -1;
}

function writeObjectToRowV2_(sheet, rowNumber, obj) {
  const headers = getSheetHeaderMapV2_(sheet);
  Object.keys(obj).forEach(function(key) {
    if (headers[key] != null) sheet.getRange(rowNumber, headers[key] + 1).setValue(obj[key]);
  });
}

function appendObjectRowV2_(sheet, obj) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(function(v) { return String(v || '').trim(); });
  const row = headers.map(function(header) { return Object.prototype.hasOwnProperty.call(obj, header) ? obj[header] : ''; });
  sheet.appendRow(row);
}

function deleteRowsMatchingV2_(sheet, predicate) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return 0;
  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim(); });
  let deleted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const obj = {};
    headers.forEach(function(h, idx) { if (h) obj[h] = values[i][idx]; });
    if (predicate(obj)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}

function summarizeScheduleGroupsV2_(groups) {
  const clean = (Array.isArray(groups) ? groups : []).map(function(group) {
    return {
      days: Array.isArray(group.days) ? group.days.map(function(v) { return String(v || '').trim(); }).filter(Boolean) : [],
      startTime: String(group.startTime || '').trim(),
      endTime: String(group.endTime || '').trim()
    };
  }).filter(function(group) { return group.days.length || group.startTime || group.endTime; });

  return {
    groups: clean,
    days: clean.map(function(g) { return g.days.join('/'); }).join(', '),
    startTime: clean.length ? clean[0].startTime : '',
    endTime: clean.length ? clean[0].endTime : ''
  };
}

