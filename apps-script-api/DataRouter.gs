function getTeacherDataSpreadsheet_(teacher) {
  const spreadsheetId = String(teacher && teacher.dataSpreadsheetId || '').trim();
  if (!spreadsheetId) {
    const error = new Error('교사 데이터 Spreadsheet가 연결되지 않았습니다.');
    error.code = 'DATASET_NOT_LINKED';
    throw error;
  }
  return SpreadsheetApp.openById(spreadsheetId);
}

function readSheetObjects_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || lastColumn < 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  const headers = values[0].map(function(v) { return String(v || '').trim(); });

  return values.slice(1).filter(function(row) {
    return row.some(function(v) { return String(v || '').trim() !== ''; });
  }).map(function(row) {
    const obj = {};
    headers.forEach(function(h, i) {
      if (h) obj[h] = row[i];
    });
    return obj;
  });
}

function getBootstrapData_(auth) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  const classes = readSheetObjects_(ss, 'Classes');
  const students = readSheetObjects_(ss, 'Students');
  const activeStudents = students.filter(function(s) {
    return String(s.status || '').trim() === '재학';
  });

  return {
    teacher: {
      teacherId: auth.teacher.teacherId,
      displayName: auth.teacher.displayName || '',
      role: auth.teacher.role || 'TEACHER'
    },
    dataset: {
      spreadsheetTitle: ss.getName()
    },
    summary: {
      classCount: classes.length,
      studentCount: students.length,
      activeStudentCount: activeStudents.length
    }
  };
}

function listClasses_(auth) {
  const ss = getTeacherDataSpreadsheet_(auth.teacher);
  return {
    classes: readSheetObjects_(ss, 'Classes')
  };
}
