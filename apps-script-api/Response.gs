function newRequestId_() {
  return Utilities.getUuid();
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function okResponse_(requestId, data) {
  return jsonResponse_({
    ok: true,
    requestId: requestId || newRequestId_(),
    data: data == null ? null : data,
    error: null
  });
}

function errorResponse_(requestId, code, message) {
  return jsonResponse_({
    ok: false,
    requestId: requestId || newRequestId_(),
    data: null,
    error: {
      code: String(code || 'ERROR'),
      message: String(message || '요청 처리 중 오류가 발생했습니다.')
    }
  });
}
