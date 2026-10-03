function getRegistrySpreadsheet_() {
  return SpreadsheetApp.openById(V2_CONFIG.ADMIN_REGISTRY_SPREADSHEET_ID);
}

function getRegistrySheet_(name) {
  const sheet = getRegistrySpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error('Registry sheet not found: ' + name);
  return sheet;
}

function getHeaderMap_(sheet) {
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

function rowToObject_(headers, row) {
  const out = {};
  Object.keys(headers).forEach(function(key) {
    out[key] = row[headers[key]];
  });
  return out;
}

function findTeacherById_(teacherId) {
  const sheet = getRegistrySheet_('Teachers');
  const headers = getHeaderMap_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const target = String(teacherId || '').trim();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[headers.teacherId] || '').trim() === target) {
      const obj = rowToObject_(headers, row);
      obj.__rowNumber = i + 2;
      return obj;
    }
  }
  return null;
}

function updateTeacherAuthFields_(teacherId, salt, verifier) {
  const sheet = getRegistrySheet_('Teachers');
  const headers = getHeaderMap_(sheet);
  const teacher = findTeacherById_(teacherId);
  if (!teacher) throw new Error('Teacher not found: ' + teacherId);

  if (headers.passwordSalt == null || headers.passwordVerifier == null) {
    throw new Error('Teachers headers are missing password fields.');
  }

  sheet.getRange(teacher.__rowNumber, headers.passwordSalt + 1).setValue(salt);
  sheet.getRange(teacher.__rowNumber, headers.passwordVerifier + 1).setValue(verifier);

  if (headers.updatedAt != null) {
    sheet.getRange(teacher.__rowNumber, headers.updatedAt + 1).setValue(new Date());
  }
}

function appendAuditLog_(actorTeacherId, action, targetType, targetId, result, detail) {
  try {
    const sheet = getRegistrySheet_('AuditLog');
    sheet.appendRow([
      'LOG-' + Utilities.getUuid(),
      new Date(),
      actorTeacherId || '',
      action || '',
      targetType || '',
      targetId || '',
      result || '',
      detail || ''
    ]);
  } catch (e) {
    console.error('appendAuditLog_ failed', e);
  }
}
