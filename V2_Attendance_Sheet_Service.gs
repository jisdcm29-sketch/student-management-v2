/**
 * V2_Attendance_Sheet_Service.gs
 * [수정된 전체 파일]
 * - 조회 함수 유지
 * - PDF 출력 레이아웃을 HTML 기반 A4 가로 인쇄 형태로 고도화
 * - 반 검색 유지
 */

/**
 * ✅ [핵심 유지] 출석부 조회
 */
function V2_AttendanceSheetService_getClassAttendanceSheet(request) {
  try {
    var normalizedRequest = V2_ASS_normalizeRequest_(request);

    var data = V2_AttendanceSheetRepository_getClassAttendanceSheet({
      classId: normalizedRequest.classId,
      startDate: normalizedRequest.startDate,
      endDate: normalizedRequest.endDate,
      includeInactive: normalizedRequest.includeInactive
    });

    return {
      success: true,
      message: '반별 기간 출석부 데이터를 불러왔습니다.',
      data: data
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '출석부 조회 중 오류가 발생했습니다.',
      data: null
    };
  }
}

/**
 * ✅ PDF 출력 데이터 생성
 * 기존 함수명 유지: 프론트 충돌 방지
 */
function V2_AttendanceSheetService_getAttendanceSheetExportData(request) {
  try {
    var normalizedRequest = V2_ASS_normalizeRequest_(request);
    var sheetResponse = V2_AttendanceSheetService_getClassAttendanceSheet(normalizedRequest);

    if (!sheetResponse.success || !sheetResponse.data) {
      throw new Error(sheetResponse.message || '출석부 출력 데이터를 생성할 수 없습니다.');
    }

    var attendanceData = sheetResponse.data || {};
    var pdfInfo = V2_ASS_createAttendanceSheetPdf_(attendanceData);

    return {
      success: true,
      message: '출석부 PDF가 생성되었습니다.',
      data: {
        fileName: pdfInfo.fileName,
        mimeType: pdfInfo.mimeType,
        base64: pdfInfo.base64
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '출석부 PDF 생성 중 오류가 발생했습니다.',
      data: {
        fileName: '',
        mimeType: 'application/pdf',
        base64: ''
      }
    };
  }
}

/**
 * 반 검색
 */
function V2_AttendanceSheetService_getClassSearchList(searchText) {
  try {
    var list = V2_AttendanceSheetRepository_getClassSearchList({
      searchText: searchText
    });

    return {
      success: true,
      message: '반 목록 조회 완료',
      data: {
        items: list
      }
    };
  } catch (error) {
    return {
      success: false,
      message: error.message || '반 목록 조회 중 오류가 발생했습니다.',
      data: {
        items: []
      }
    };
  }
}

/**
 * PDF 생성
 * - HTML 기반
 * - A4 landscape 고정
 * - 상단부터 꽉 차게 배치
 */
function V2_ASS_createAttendanceSheetPdf_(attendanceData) {
  try {
    attendanceData = attendanceData || {};

    var html = V2_ASS_buildAttendancePdfHtml_(attendanceData);
    var fileName = V2_ASS_buildPdfFileName_(attendanceData);

    var blob = HtmlService
      .createHtmlOutput(html)
      .getBlob()
      .getAs(MimeType.PDF)
      .setName(fileName);

    var base64 = Utilities.base64Encode(blob.getBytes());

    return {
      fileName: fileName,
      mimeType: 'application/pdf',
      base64: base64
    };
  } catch (error) {
    throw error;
  }
}

/**
 * HTML PDF 본문 생성
 */
function V2_ASS_buildAttendancePdfHtml_(attendanceData) {
  try {
    var classInfo = attendanceData.classInfo || {};
    var period = attendanceData.period || {};
    var summary = attendanceData.summary || {};
    var students = Array.isArray(attendanceData.studentRows) ? attendanceData.studentRows : [];
    var dateColumns = Array.isArray(attendanceData.dateColumns) ? attendanceData.dateColumns : [];

    var layout = V2_ASS_getPdfLayoutOptions_(dateColumns.length);

    var html = '';

    html += '<!DOCTYPE html>';
    html += '<html>';
    html += '<head>';
    html += '<meta charset="UTF-8">';
    html += '<style>';
    html += '@page { size: A4 landscape; margin: 6mm 6mm 6mm 6mm; }';
    html += 'html, body { width: 297mm; min-height: 210mm; margin:0; padding:0; }';
    html += 'body {';
    html += '  font-family: Arial, "Malgun Gothic", "Apple SD Gothic Neo", sans-serif;';
    html += '  color:#0f172a;';
    html += '  background:#ffffff;';
    html += '  -webkit-print-color-adjust: exact !important;';
    html += '  print-color-adjust: exact !important;';
    html += '  font-size:' + layout.baseFontSize + 'px;';
    html += '}';
    html += '.page {';
    html += '  width: 285mm;';
    html += '  min-height: 198mm;';
    html += '  box-sizing: border-box;';
    html += '  margin: 0;';
    html += '  padding: 0;';
    html += '}';
    html += '.title-wrap { margin: 0 0 6px 0; padding-bottom: 5px; border-bottom: 2px solid #1d4ed8; }';
    html += '.title { font-size: 19px; font-weight: 700; text-align: center; margin: 0 0 2px 0; color:#0f172a; letter-spacing: -0.2px; }';
    html += '.subtitle { font-size: 10px; text-align: center; color:#475569; margin: 0; }';
    html += '.meta-table { width: 100%; border-collapse: collapse; table-layout: fixed; margin: 0 0 6px 0; }';
    html += '.meta-table td { border: 1px solid #cbd5e1; padding: 5px 6px; vertical-align: middle; line-height: 1.25; }';
    html += '.meta-label { background: #eff6ff; font-weight: 700; text-align: center; width: 8%; }';
    html += '.meta-value { background: #ffffff; width: 17%; }';
    html += '.summary-table { width: 100%; border-collapse: collapse; table-layout: fixed; margin: 0 0 7px 0; }';
    html += '.summary-table td { border: 1px solid #cbd5e1; padding: 5px 4px; text-align: center; vertical-align: middle; line-height: 1.2; }';
    html += '.summary-label { background: #f8fafc; font-weight: 700; color:#334155; }';
    html += '.summary-value-cell { background:#ffffff; font-weight:700; }';
    html += '.summary-present { background:#dcfce7; color:#166534; }';
    html += '.summary-late { background:#ffedd5; color:#9a3412; }';
    html += '.summary-absent { background:#fee2e2; color:#991b1b; }';
    html += '.summary-excused { background:#dbeafe; color:#1d4ed8; }';
    html += '.legend { margin: 0 0 6px 0; font-size: 9px; color:#475569; text-align:right; }';
    html += '.legend span { display:inline-block; margin-left:8px; padding:2px 6px; border-radius:10px; font-weight:700; }';
    html += '.legend .lg-present { background:#dcfce7; color:#166534; }';
    html += '.legend .lg-late { background:#ffedd5; color:#9a3412; }';
    html += '.legend .lg-absent { background:#fee2e2; color:#991b1b; }';
    html += '.legend .lg-excused { background:#dbeafe; color:#1d4ed8; }';
    html += '.attendance-table { width: 100%; border-collapse: collapse; table-layout: fixed; border: 1.4px solid #64748b; }';
    html += '.attendance-table th, .attendance-table td {';
    html += '  border: 1px solid #94a3b8;';
    html += '  padding: ' + layout.cellPaddingVertical + 'px ' + layout.cellPaddingHorizontal + 'px;';
    html += '  text-align: center;';
    html += '  vertical-align: middle;';
    html += '  line-height: 1.15;';
    html += '  word-break: keep-all;';
    html += '  overflow: hidden;';
    html += '}';
    html += '.attendance-table thead th {';
    html += '  background: #dbeafe;';
    html += '  color:#0f172a;';
    html += '  font-weight: 700;';
    html += '  font-size: ' + layout.headerFontSize + 'px;';
    html += '}';
    html += '.attendance-table tbody td { font-size: ' + layout.bodyFontSize + 'px; }';
    html += '.attendance-table thead .name-header { background:#bfdbfe; }';
    html += '.attendance-table thead .sum-header { background:#dbeafe; }';
    html += '.attendance-table tbody tr:nth-child(even) td { background:#f8fafc; }';
    html += '.col-no { width: 4.5%; }';
    html += '.col-name { width: 10%; }';
    html += '.date-col { width: ' + layout.dateColumnWidth + '%; }';
    html += '.sum-col { width: 4.8%; }';
    html += '.name-cell { text-align: left !important; padding-left: 6px !important; font-weight: 700; color:#1e293b; }';
    html += '.status-present { background:#dcfce7 !important; color:#166534; font-weight:700; }';
    html += '.status-late { background:#ffedd5 !important; color:#9a3412; font-weight:700; }';
    html += '.status-absent { background:#fee2e2 !important; color:#991b1b; font-weight:700; }';
    html += '.status-excused { background:#dbeafe !important; color:#1d4ed8; font-weight:700; }';
    html += '.status-empty { color:#94a3b8; }';
    html += '.empty-row td { padding: 18px 8px; color:#6b7280; background:#ffffff !important; }';
    html += '</style>';
    html += '</head>';
    html += '<body>';
    html += '<div class="page">';

    html += '<div class="title-wrap">';
    html += '<div class="title">V2 반별 기간 출석부</div>';
    html += '<div class="subtitle">학생 행 / 날짜 열 기준 A4 출력용 표준 양식</div>';
    html += '</div>';

    html += '<table class="meta-table">';
    html += '<tr>';
    html += '<td class="meta-label">반명</td>';
    html += '<td class="meta-value">' + V2_ASS_escapeHtml_(classInfo.className || '') + '</td>';
    html += '<td class="meta-label">반ID</td>';
    html += '<td class="meta-value">' + V2_ASS_escapeHtml_(classInfo.classId || '') + '</td>';
    html += '<td class="meta-label">조회 기간</td>';
    html += '<td class="meta-value">' + V2_ASS_escapeHtml_((period.startDate || '') + ' ~ ' + (period.endDate || '')) + '</td>';
    html += '<td class="meta-label">출력일시</td>';
    html += '<td class="meta-value">' + V2_ASS_escapeHtml_(attendanceData.generatedAt || '') + '</td>';
    html += '</tr>';
    html += '</table>';

    html += '<table class="summary-table">';
    html += '<tr>';
    html += '<td class="summary-label">학생 수</td>';
    html += '<td class="summary-value-cell">' + V2_ASS_toNumber_(summary.totalStudentCount) + '</td>';
    html += '<td class="summary-label">출석 입력 건수</td>';
    html += '<td class="summary-value-cell">' + V2_ASS_toNumber_(summary.totalAttendanceCount) + '</td>';
    html += '<td class="summary-label">출석</td>';
    html += '<td class="summary-value-cell summary-present">' + V2_ASS_toNumber_(summary.presentCount) + '</td>';
    html += '<td class="summary-label">지각</td>';
    html += '<td class="summary-value-cell summary-late">' + V2_ASS_toNumber_(summary.lateCount) + '</td>';
    html += '<td class="summary-label">결석</td>';
    html += '<td class="summary-value-cell summary-absent">' + V2_ASS_toNumber_(summary.absentCount) + '</td>';
    html += '<td class="summary-label">인정결석</td>';
    html += '<td class="summary-value-cell summary-excused">' + V2_ASS_toNumber_(summary.excusedCount) + '</td>';
    html += '</tr>';
    html += '</table>';

    html += '<div class="legend">';
    html += '<span class="lg-present">출석</span>';
    html += '<span class="lg-late">지각</span>';
    html += '<span class="lg-absent">결석</span>';
    html += '<span class="lg-excused">인정결석</span>';
    html += '</div>';

    html += '<table class="attendance-table">';
    html += '<thead>';
    html += '<tr>';
    html += '<th class="col-no">번호</th>';
    html += '<th class="col-name name-header">학생명</th>';

    for (var i = 0; i < dateColumns.length; i++) {
      html += '<th class="date-col">' + V2_ASS_escapeHtml_(V2_ASS_buildPdfDateLabel_(dateColumns[i])) + '</th>';
    }

    html += '<th class="sum-col sum-header">출석</th>';
    html += '<th class="sum-col sum-header">지각</th>';
    html += '<th class="sum-col sum-header">결석</th>';
    html += '<th class="sum-col sum-header">인정결석</th>';
    html += '</tr>';
    html += '</thead>';
    html += '<tbody>';

    if (students.length < 1) {
      html += '<tr class="empty-row">';
      html += '<td colspan="' + (dateColumns.length + 6) + '">표시할 학생이 없습니다.</td>';
      html += '</tr>';
    } else {
      for (var s = 0; s < students.length; s++) {
        html += V2_ASS_buildAttendancePdfBodyRowHtml_(students[s], s + 1, dateColumns);
      }
    }

    html += '</tbody>';
    html += '</table>';

    html += '</div>';
    html += '</body>';
    html += '</html>';

    return html;
  } catch (error) {
    throw new Error('PDF HTML 구성 중 오류가 발생했습니다. ' + error.message);
  }
}

function V2_ASS_buildAttendancePdfBodyRowHtml_(studentRow, number, dateColumns) {
  try {
    studentRow = studentRow || {};
    dateColumns = Array.isArray(dateColumns) ? dateColumns : [];

    var counts = studentRow.counts || {};
    var html = '';

    html += '<tr>';
    html += '<td>' + number + '</td>';
    html += '<td class="name-cell">' + V2_ASS_escapeHtml_(studentRow.studentName || '') + '</td>';

    for (var i = 0; i < dateColumns.length; i++) {
      var statusText = V2_ASS_findStatusForDate_(studentRow.dateStatuses, dateColumns[i]);
      var statusClass = V2_ASS_getPdfStatusClass_(statusText);

      html += '<td class="' + statusClass + '">';
      html += V2_ASS_escapeHtml_(statusText || '');
      html += '</td>';
    }

    html += '<td>' + V2_ASS_toNumber_(counts.presentCount) + '</td>';
    html += '<td>' + V2_ASS_toNumber_(counts.lateCount) + '</td>';
    html += '<td>' + V2_ASS_toNumber_(counts.absentCount) + '</td>';
    html += '<td>' + V2_ASS_toNumber_(counts.excusedCount) + '</td>';
    html += '</tr>';

    return html;
  } catch (error) {
    return '';
  }
}

function V2_ASS_getPdfLayoutOptions_(dateColumnCount) {
  try {
    var count = Number(dateColumnCount || 0);

    if (count <= 7) {
      return {
        dateColumnWidth: 6.2,
        baseFontSize: 11,
        headerFontSize: 11,
        bodyFontSize: 11,
        cellPaddingVertical: 5,
        cellPaddingHorizontal: 4
      };
    }

    if (count <= 10) {
      return {
        dateColumnWidth: 5.1,
        baseFontSize: 10,
        headerFontSize: 10,
        bodyFontSize: 10,
        cellPaddingVertical: 4,
        cellPaddingHorizontal: 3
      };
    }

    if (count <= 14) {
      return {
        dateColumnWidth: 4.0,
        baseFontSize: 9,
        headerFontSize: 9,
        bodyFontSize: 9,
        cellPaddingVertical: 3,
        cellPaddingHorizontal: 2
      };
    }

    return {
      dateColumnWidth: 3.5,
      baseFontSize: 8,
      headerFontSize: 8,
      bodyFontSize: 8,
      cellPaddingVertical: 2,
      cellPaddingHorizontal: 1
    };
  } catch (error) {
    return {
      dateColumnWidth: 5,
      baseFontSize: 10,
      headerFontSize: 10,
      bodyFontSize: 10,
      cellPaddingVertical: 4,
      cellPaddingHorizontal: 3
    };
  }
}

function V2_ASS_getPdfStatusClass_(status) {
  try {
    var text = String(status || '').trim().toLowerCase();

    if (!text) {
      return 'status-empty';
    }

    if (text === '출석' || text === '재석' || text === 'present') {
      return 'status-present';
    }

    if (text === '지각' || text === 'late') {
      return 'status-late';
    }

    if (text === '공결' || text === '인정결석' || text === 'excused') {
      return 'status-excused';
    }

    return 'status-absent';
  } catch (error) {
    return 'status-empty';
  }
}

function V2_ASS_findStatusForDate_(dateStatuses, dateColumn) {
  try {
    dateStatuses = Array.isArray(dateStatuses) ? dateStatuses : [];
    var targetDate = V2_ASS_extractDateText_(dateColumn);

    for (var i = 0; i < dateStatuses.length; i++) {
      var item = dateStatuses[i] || {};
      if (V2_ASS_normalizeDateText_(item.date) === targetDate) {
        return item.status || '';
      }
    }

    return '';
  } catch (error) {
    return '';
  }
}

function V2_ASS_buildPdfDateLabel_(dateColumn) {
  try {
    var dateText = V2_ASS_extractDateText_(dateColumn);
    if (!dateText) {
      return '';
    }

    var date = new Date(dateText + 'T00:00:00');
    if (isNaN(date.getTime())) {
      return dateText;
    }

    var dayNames = ['일', '월', '화', '수', '목', '금', '토'];
    var month = ('0' + (date.getMonth() + 1)).slice(-2);
    var day = ('0' + date.getDate()).slice(-2);

    return month + '/' + day + '(' + dayNames[date.getDay()] + ')';
  } catch (error) {
    return '';
  }
}

function V2_ASS_extractDateText_(dateColumn) {
  try {
    if (!dateColumn) {
      return '';
    }

    if (typeof dateColumn === 'string') {
      return V2_ASS_normalizeDateText_(dateColumn);
    }

    return V2_ASS_normalizeDateText_(dateColumn.date || '');
  } catch (error) {
    return '';
  }
}

function V2_ASS_buildPdfFileName_(attendanceData) {
  try {
    var classInfo = attendanceData.classInfo || {};
    var period = attendanceData.period || {};

    return [
      'V2_출석부',
      V2_ASS_sanitizeFilePart_(classInfo.className || classInfo.classId || '반'),
      period.startDate || 'start',
      period.endDate || 'end'
    ].join('_') + '.pdf';
  } catch (error) {
    return 'V2_출석부.pdf';
  }
}

function V2_ASS_sanitizeFilePart_(value) {
  try {
    return String(value || '').replace(/[\\\/:*?"<>|]/g, '_');
  } catch (error) {
    return 'FILE';
  }
}

function V2_ASS_escapeHtml_(value) {
  try {
    return String(value === null || value === undefined ? '' : value)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  } catch (error) {
    return '';
  }
}

/**
 * 공통 요청 정규화
 */
function V2_ASS_normalizeRequest_(request) {
  request = request || {};

  return {
    classId: String(request.classId || '').trim(),
    startDate: String(request.startDate || '').trim(),
    endDate: String(request.endDate || '').trim(),
    includeInactive: !!request.includeInactive
  };
}

function V2_ASS_toNumber_(value) {
  try {
    var numberValue = Number(value);
    return isNaN(numberValue) ? 0 : numberValue;
  } catch (error) {
    return 0;
  }
}

function V2_ASS_normalizeDateText_(value) {
  try {
    if (value === null || value === undefined || value === '') {
      return '';
    }

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    var textValue = String(value).trim();

    if (!textValue) {
      return '';
    }

    textValue = textValue.replace(/\./g, '-').replace(/\//g, '-');
    textValue = textValue.replace(/\s+/g, '');

    var directMatch = textValue.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (directMatch) {
      return [
        directMatch[1],
        ('0' + directMatch[2]).slice(-2),
        ('0' + directMatch[3]).slice(-2)
      ].join('-');
    }

    var parsedDate = new Date(textValue);
    if (!isNaN(parsedDate.getTime())) {
      return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }

    return textValue;
  } catch (error) {
    return String(value || '').trim();
  }
}
