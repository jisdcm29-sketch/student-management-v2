/**
 * 수정된 전체 파일: V2_Absence_Contact_Service.gs
 * 결석 연락 기록 비즈니스 로직
 */

function V2_AbsenceContactService_getAllowedStatuses_() {
  try {
    return [
      V2_CONFIG.ABSENCE_CONTACT_STATUS.NOT_CONTACTED,
      V2_CONFIG.ABSENCE_CONTACT_STATUS.CONTACTED,
      V2_CONFIG.ABSENCE_CONTACT_STATUS.NO_ANSWER,
      V2_CONFIG.ABSENCE_CONTACT_STATUS.CALLBACK_NEEDED
    ];
  } catch (error) {
    throw new Error('V2_AbsenceContactService_getAllowedStatuses_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactService_normalizeInput_(data) {
  try {
    data = data || {};

    return {
      attendanceId: String(data.attendanceId || '').trim(),
      studentId: String(data.studentId || '').trim(),
      classId: String(data.classId || '').trim(),
      contactDate: String(data.contactDate || '').trim(),
      contactStatus: String(
        data.contactStatus || V2_CONFIG.ABSENCE_CONTACT_STATUS.NOT_CONTACTED
      ).trim(),
      parentPhone: String(data.parentPhone || '').trim(),
      memo: String(data.memo || '').trim()
    };
  } catch (error) {
    throw new Error('V2_AbsenceContactService_normalizeInput_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactService_validateInput_(data) {
  try {
    if (!data.attendanceId) {
      throw new Error('attendanceId는 필수입니다.');
    }

    if (!data.studentId) {
      throw new Error('studentId는 필수입니다.');
    }

    if (!data.classId) {
      throw new Error('classId는 필수입니다.');
    }

    if (!data.contactDate) {
      throw new Error('contactDate는 필수입니다.');
    }

    if (V2_AbsenceContactService_getAllowedStatuses_().indexOf(data.contactStatus) === -1) {
      throw new Error('허용되지 않은 연락 상태입니다: ' + data.contactStatus);
    }
  } catch (error) {
    throw new Error('V2_AbsenceContactService_validateInput_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactService_buildEntity_(data, existing) {
  try {
    var now = V2_nowText_();
    var teacherName = String(V2_getCurrentUserName_() || '').trim();

    if (existing) {
      return {
        contactId: existing.contactId || V2_createId_(),
        attendanceId: data.attendanceId,
        studentId: data.studentId,
        classId: data.classId,
        contactDate: data.contactDate,
        contactStatus: data.contactStatus,
        parentPhone: data.parentPhone,
        memo: data.memo,
        teacherName: teacherName || String(existing.teacherName || '').trim(),
        createdAt: existing.createdAt || now,
        updatedAt: now
      };
    }

    return {
      contactId: V2_createId_(),
      attendanceId: data.attendanceId,
      studentId: data.studentId,
      classId: data.classId,
      contactDate: data.contactDate,
      contactStatus: data.contactStatus,
      parentPhone: data.parentPhone,
      memo: data.memo,
      teacherName: teacherName,
      createdAt: now,
      updatedAt: now
    };
  } catch (error) {
    throw new Error('V2_AbsenceContactService_buildEntity_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactService_normalizeAttendanceIds_(attendanceIds) {
  try {
    attendanceIds = Array.isArray(attendanceIds) ? attendanceIds : [];

    var result = [];
    var seen = {};

    attendanceIds.forEach(function(attendanceId) {
      var value = String(attendanceId || '').trim();
      if (!value) {
        return;
      }

      if (seen[value]) {
        return;
      }

      seen[value] = true;
      result.push(value);
    });

    return result;
  } catch (error) {
    throw new Error('V2_AbsenceContactService_normalizeAttendanceIds_ 오류: ' + error.message);
  }
}

function V2_AbsenceContactService_create(data) {
  try {
    var normalized = V2_AbsenceContactService_normalizeInput_(data);
    V2_AbsenceContactService_validateInput_(normalized);

    var existing = V2_AbsenceContactRepository_findByAttendanceId(normalized.attendanceId);
    var entity = V2_AbsenceContactService_buildEntity_(normalized, existing);
    var result = V2_AbsenceContactRepository_upsertByAttendanceId(entity);

    if (!result || !result.data) {
      throw new Error('결석 연락 기록 저장 결과가 올바르지 않습니다.');
    }

    var savedRecord = V2_AbsenceContactRepository_findByAttendanceId(normalized.attendanceId);

    if (!savedRecord) {
      throw new Error('저장 후 결석 연락 기록 재조회에 실패했습니다.');
    }

    V2_log_(
      'INFO',
      'V2_AbsenceContactService_create',
      result.mode === 'insert' ? '결석 연락 기록 신규 저장' : '결석 연락 기록 갱신',
      JSON.stringify({
        mode: result.mode || '',
        attendanceId: savedRecord.attendanceId || '',
        studentId: savedRecord.studentId || '',
        classId: savedRecord.classId || '',
        contactStatus: savedRecord.contactStatus || '',
        updatedAt: savedRecord.updatedAt || ''
      })
    );

    return V2_createSuccessResponse_(savedRecord);
  } catch (error) {
    V2_log_('ERROR', 'V2_AbsenceContactService_create', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_AbsenceContactService_getByAttendanceId(attendanceId) {
  try {
    attendanceId = String(attendanceId || '').trim();

    if (!attendanceId) {
      throw new Error('attendanceId는 필수입니다.');
    }

    var found = V2_AbsenceContactRepository_findByAttendanceId(attendanceId);

    return V2_createSuccessResponse_(found);
  } catch (error) {
    V2_log_('ERROR', 'V2_AbsenceContactService_getByAttendanceId', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}

function V2_AbsenceContactService_getMapByAttendanceIds(attendanceIds) {
  try {
    var normalizedIds = V2_AbsenceContactService_normalizeAttendanceIds_(attendanceIds);
    var dataMap = V2_AbsenceContactRepository_findMapByAttendanceIds(normalizedIds);

    return V2_createSuccessResponse_({
      attendanceIds: normalizedIds,
      recordMap: dataMap
    });
  } catch (error) {
    V2_log_('ERROR', 'V2_AbsenceContactService_getMapByAttendanceIds', error.message, error.stack || '');
    return V2_createErrorResponse_(error);
  }
}
