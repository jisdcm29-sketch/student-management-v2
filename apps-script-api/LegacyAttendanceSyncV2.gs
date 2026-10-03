const LEGACY_ATTENDANCE_SYNC_V2_ = Object.freeze({
  LEGACY_ID: '1Y5qoA0mQp-7EAQoXM7GsOF6y3MkdT5XhVEmjGjQc_TY',
  TARGET_ID: '1fRrmhub9KbvqJRsBZv1UL0KS-48I5BYybe6ANblrw8Q',
  TARGET_SHEET: 'LegacyAttendanceHistory',
  ELIGIBILITY_SHEET: 'StudentLearningProgress'
});

const LEGACY_ATTENDANCE_HEADERS_V2_ = [
  'sourceKey','sourceAttendanceId','studentId','currentClassId','normalizedPhone',
  'sourceStudentId','sourceStudentName','sourceStudentStatus','sourceClassId',
  'date','status','memo','sourceUpdatedAt','sourceRow','checksum','syncedAt'
];

function legacyAttendanceTextV2_(value) {
  return String(value == null ? '' : value).trim();
}

function legacyAttendancePhoneV2_(value) {
  return String(value || '').replace(/\D/g, '');
}

function legacyAttendanceHexV2_(bytes) {
  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function legacyAttendanceChecksumV2_(values) {
  return legacyAttendanceHexV2_(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(values || []),
    Utilities.Charset.UTF_8
  ));
}

function legacyAttendanceReadObjectsV2_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + sheetName);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return [];
  const values = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(legacyAttendanceTextV2_);
  const out = [];
  for (let r = 1; r < values.length; r++) {
    if (!values[r].some(function(v){ return legacyAttendanceTextV2_(v) !== ''; })) continue;
    const obj = { __row: r + 1 };
    headers.forEach(function(h, i){ if (h) obj[h] = values[r][i]; });
    out.push(obj);
  }
  return out;
}

function legacyAttendanceEligibleMapV2_(targetSs) {
  const rows = legacyAttendanceReadObjectsV2_(targetSs, LEGACY_ATTENDANCE_SYNC_V2_.ELIGIBILITY_SHEET);
  const map = {};
  rows.forEach(function(row) {
    const phone = legacyAttendancePhoneV2_(row.normalizedPhone);
    if (!phone) return;
    map[phone] = {
      studentId: legacyAttendanceTextV2_(row.studentId),
      currentClassId: ''
    };
  });

  const students = legacyAttendanceReadObjectsV2_(targetSs, 'Students');
  students.forEach(function(row) {
    const phone = legacyAttendancePhoneV2_(row.phone);
    if (!map[phone]) return;
    map[phone].studentId = legacyAttendanceTextV2_(row.studentId) || map[phone].studentId;
    map[phone].currentClassId = legacyAttendanceTextV2_(row.classId);
  });
  return map;
}

function legacyAttendanceCollectV2_() {
  const legacySs = SpreadsheetApp.openById(LEGACY_ATTENDANCE_SYNC_V2_.LEGACY_ID);
  const targetSs = SpreadsheetApp.openById(LEGACY_ATTENDANCE_SYNC_V2_.TARGET_ID);
  const eligible = legacyAttendanceEligibleMapV2_(targetSs);

  const oldStudents = legacyAttendanceReadObjectsV2_(legacySs, 'Students');
  const studentById = {};
  oldStudents.forEach(function(row){
    const sid = legacyAttendanceTextV2_(row.studentId);
    if (sid) studentById[sid] = row;
  });

  const raw = legacyAttendanceReadObjectsV2_(legacySs, 'Attendance');
  const now = new Date();
  const records = [];
  const seenKeys = {};

  raw.forEach(function(row) {
    const sourceStudentId = legacyAttendanceTextV2_(row.studentId);
    const sourceStudent = studentById[sourceStudentId];
    if (!sourceStudent) return;
    const phone = legacyAttendancePhoneV2_(sourceStudent.phone);
    const link = eligible[phone];
    if (!link || !link.studentId) return;

    const sourceAttendanceId = legacyAttendanceTextV2_(row.attendanceId);
    if (!sourceAttendanceId) return;

    // 이전 운영본에는 attendanceId 중복이 실제로 존재한다.
    // 따라서 attendanceId 단독이 아니라 원본 의미를 보존하는 복합키를 사용한다.
    const sourceKey = [
      sourceAttendanceId,
      legacyAttendanceTextV2_(row.date),
      sourceStudentId,
      legacyAttendanceTextV2_(row.classId)
    ].join('|');
    if (seenKeys[sourceKey]) {
      throw new Error('원본 Attendance에 동일 복합키가 중복됩니다: ' + sourceKey);
    }
    seenKeys[sourceKey] = true;

    const checksum = legacyAttendanceChecksumV2_([
      row.attendanceId,row.date,row.classId,row.studentId,row.status,row.memo,row.updatedAt,
      sourceStudent.name,sourceStudent.phone,sourceStudent.status,link.studentId,link.currentClassId
    ]);

    records.push({
      sourceKey: sourceKey,
      sourceAttendanceId: sourceAttendanceId,
      studentId: link.studentId,
      currentClassId: link.currentClassId,
      normalizedPhone: phone,
      sourceStudentId: sourceStudentId,
      sourceStudentName: legacyAttendanceTextV2_(sourceStudent.name),
      sourceStudentStatus: legacyAttendanceTextV2_(sourceStudent.status),
      sourceClassId: legacyAttendanceTextV2_(row.classId),
      date: legacyAttendanceTextV2_(row.date),
      status: legacyAttendanceTextV2_(row.status),
      memo: legacyAttendanceTextV2_(row.memo),
      sourceUpdatedAt: legacyAttendanceTextV2_(row.updatedAt),
      sourceRow: row.__row,
      checksum: checksum,
      syncedAt: now
    });
  });

  return {
    targetSs: targetSs,
    scannedRows: raw.length,
    eligibleStudentCount: Object.keys(eligible).length,
    records: records
  };
}

function legacyAttendanceEnsureSheetV2_(targetSs) {
  let sheet = targetSs.getSheetByName(LEGACY_ATTENDANCE_SYNC_V2_.TARGET_SHEET);
  if (!sheet) {
    sheet = targetSs.insertSheet(LEGACY_ATTENDANCE_SYNC_V2_.TARGET_SHEET);
    sheet.getRange(1, 1, 1, LEGACY_ATTENDANCE_HEADERS_V2_.length).setValues([LEGACY_ATTENDANCE_HEADERS_V2_]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const current = sheet.getRange(1, 1, 1, LEGACY_ATTENDANCE_HEADERS_V2_.length).getDisplayValues()[0];
  const same = LEGACY_ATTENDANCE_HEADERS_V2_.every(function(h, i){ return legacyAttendanceTextV2_(current[i]) === h; });
  if (!same) throw new Error('LegacyAttendanceHistory 헤더가 예상 구조와 다릅니다. 자동 수정하지 않습니다.');
  return sheet;
}

function legacyAttendanceExistingMapV2_(sheet) {
  const out = {};
  if (sheet.getLastRow() < 2) return out;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, LEGACY_ATTENDANCE_HEADERS_V2_.length).getDisplayValues();
  const keyIndex = LEGACY_ATTENDANCE_HEADERS_V2_.indexOf('sourceKey');
  const checksumIndex = LEGACY_ATTENDANCE_HEADERS_V2_.indexOf('checksum');
  values.forEach(function(row, i) {
    const key = legacyAttendanceTextV2_(row[keyIndex]);
    if (!key) return;
    out[key] = { rowNumber: i + 2, checksum: legacyAttendanceTextV2_(row[checksumIndex]) };
  });
  return out;
}

function legacyAttendanceRowV2_(record) {
  return LEGACY_ATTENDANCE_HEADERS_V2_.map(function(h){ return Object.prototype.hasOwnProperty.call(record,h) ? record[h] : ''; });
}

function previewLegacyAttendanceSyncV2() {
  const data = legacyAttendanceCollectV2_();
  const sheet = data.targetSs.getSheetByName(LEGACY_ATTENDANCE_SYNC_V2_.TARGET_SHEET);
  let wouldInsert = data.records.length, wouldUpdate = 0, wouldSkip = 0;
  if (sheet) {
    const existing = legacyAttendanceExistingMapV2_(sheet);
    wouldInsert = 0;
    data.records.forEach(function(record){
      const found = existing[record.sourceKey];
      if (!found) wouldInsert++;
      else if (found.checksum === record.checksum) wouldSkip++;
      else wouldUpdate++;
    });
  }
  const byStatus = {};
  const bySourceClass = {};
  data.records.forEach(function(r){
    byStatus[r.status] = (byStatus[r.status] || 0) + 1;
    bySourceClass[r.sourceClassId] = (bySourceClass[r.sourceClassId] || 0) + 1;
  });
  const result = {
    mode: 'PREVIEW_LEGACY_ATTENDANCE',
    eligibleStudentCount: data.eligibleStudentCount,
    scannedRows: data.scannedRows,
    matchedRows: data.records.length,
    changes: { wouldInsert:wouldInsert, wouldUpdate:wouldUpdate, wouldSkip:wouldSkip },
    byStatus: byStatus,
    bySourceClass: bySourceClass,
    targetSheetExists: !!sheet,
    note: '이 함수는 기존 Attendance를 수정하지 않습니다. LegacyAttendanceHistory 전용 보관 시트만 대상으로 합니다.'
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function applyLegacyAttendanceSyncV2() {
  const data = legacyAttendanceCollectV2_();
  const sheet = legacyAttendanceEnsureSheetV2_(data.targetSs);
  const existing = legacyAttendanceExistingMapV2_(sheet);
  const appendRows = [];
  let updated = 0, skipped = 0;

  data.records.forEach(function(record){
    const found = existing[record.sourceKey];
    if (!found) {
      appendRows.push(legacyAttendanceRowV2_(record));
      return;
    }
    if (found.checksum === record.checksum) {
      skipped++;
      return;
    }
    sheet.getRange(found.rowNumber, 1, 1, LEGACY_ATTENDANCE_HEADERS_V2_.length)
      .setValues([legacyAttendanceRowV2_(record)]);
    updated++;
  });

  if (appendRows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, LEGACY_ATTENDANCE_HEADERS_V2_.length)
      .setValues(appendRows);
  }

  const result = {
    mode: 'APPLY_LEGACY_ATTENDANCE',
    eligibleStudentCount: data.eligibleStudentCount,
    scannedRows: data.scannedRows,
    matchedRows: data.records.length,
    results: { inserted:appendRows.length, updated:updated, skipped:skipped },
    targetSheet: LEGACY_ATTENDANCE_SYNC_V2_.TARGET_SHEET
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}
