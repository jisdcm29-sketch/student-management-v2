# STUDENT MANAGEMENT V2 — 수업 내용 관리 시방서

작성일: 2026-10-04
대상: student-management-v2
상태: Step 7C 과거 수업 데이터 안전 이관 준비

## 1. 목적
기존 Google 기반 학생관리 프로그램의 수업 내용 입력/조회 기능을 계승하되, 장기 데이터 누적에도 프로그램 전체 속도에 영향을 최소화하고, 기존 V2 구조를 깨지 않는 방식으로 단계적으로 이전한다.

## 2. 최우선 원칙
1. 기존 정상 기능을 변경하거나 깨뜨리지 않는다.
2. 수업 내용 데이터는 수업 내용 메뉴를 열기 전에는 읽지 않는다.
3. 기본 조회는 전체가 아니라 최근 기록만 읽는다.
4. 반(classId) + 날짜 범위를 우선 조회 조건으로 사용한다.
5. 신규 저장은 한 행 추가, 수정은 해당 lessonId 행만 수정한다.
6. 전체 Lessons 시트를 매번 다시 쓰는 방식은 금지한다.
7. 학생별 과제는 중복 저장하지 않고 기존 개별 지도 기록과 역할을 분명히 한다.
8. 화면에는 필요한 데이터만 렌더링하고 수백/수천 건을 한 번에 출력하지 않는다.
9. 기존 Google 기반 수업 내용의 장점은 유지하고, 불명확한 진도 단위·긴 목록·반 식별 혼재는 개선한다.
10. 구현 전/후 이 문서를 다시 확인하고 변경 사항은 이 문서의 변경 이력에 기록한다.

## 3. 데이터 증가와 성능 원칙
핵심 원칙: "데이터는 많이 저장해도 되지만, 한 번에 많이 읽지 않는다."

### 3.1 초기 로딩
- 프로그램 로그인 시 Lessons 데이터를 읽지 않는다.
- 홈/학생/출석/성적 초기화 과정에서 Lessons 조회를 호출하지 않는다.
- 수업 내용 메뉴 진입 시에만 필요한 데이터를 호출한다.

### 3.2 기본 조회
- 기본 조회는 최근 30건이다.
- 전체 기록을 한 번에 반환하지 않는다.
- Apps Script는 시트 끝에서 일정 청크만 역방향으로 읽어 필요한 기록을 찾는다.
- Step 2 기준: 200행 청크, 최대 4,000행 스캔, 반환 최대 100건.
- 향후 실제 데이터량/속도 측정 후 인덱스 시트 도입 여부를 판단한다.

### 3.3 조건 조회
우선 조건:
- classId
- startDate
- endDate
- 필요 시 lessonId

### 3.4 저장
- 신규: append 방식
- 수정: lessonId로 해당 행만 갱신
- 저장 한 번에 전체 시트 rewrite 금지
- 삭제 정책은 Step 3에서 확정

### 3.5 렌더링
- 최근 기록만 화면에 표시
- 긴 목록은 더 보기/기간 조회 방식
- 입력 화면과 목록 화면을 분리하지 않고 기존 Google 기반의 빠른 흐름을 유지

## 4. 기존 Google 기반 화면에서 유지할 기능
반드시 유지:
- 반 선택
- 날짜
- 주제
- 학습 내용
- 숙제
- 다음 계획
- 누적 진도
- 학생별 과제
- 저장된 수업 목록
- 반별 조회
- 수정
- 삭제/보관 기능
- 입력 중 초안 자동저장(기존 JSLesson의 장점, 화면 단계에서 재도입 검토)

## 5. Step 1 실제 구조 확인 결과

### 5.1 현재 V2 교사 데이터 스프레드시트
현재 실제 V2 교사 데이터 스프레드시트에 다음 시트가 이미 존재한다.
- Lessons
- LessonAssignments
- Classes
- StudentDailyRecords

### 5.2 현재 Lessons 실제 헤더
현재 헤더는 다음 9개다.
- lessonId
- date
- classId
- topic
- content
- homework
- nextPlan
- progressCurrentCount
- createdAt

현재 Step 1 확인 시 데이터 행은 없고 헤더만 존재한다.

결론:
- 새 Lessons 시트를 만들지 않는다.
- 기존 헤더를 기준으로 Step 2를 구현한다.
- progressValue/progressUnit 새 열은 지금 추가하지 않는다.
- 현재 progressCurrentCount는 유지하고 단위는 Classes.targetProgressUnit에서 화면 표시 시 결합한다.

### 5.3 현재 LessonAssignments 실제 헤더
- assignmentId
- lessonId
- classId
- studentId
- assignmentText
- createdAt
- legacyCreatedAt2
- assignment
- studentName

현재 Step 1 확인 시 데이터 행은 없고 헤더만 존재한다.

결론:
- 새 LessonAssignments 시트를 만들지 않는다.
- legacyCreatedAt2 / assignment 같은 이전 호환 필드는 지금 삭제하지 않는다.
- 학생별 과제 저장 방식은 Step 5에서 DailyCoachingV2와 역할 중복을 다시 비교한 뒤 확정한다.

### 5.4 Classes 진도 필드
현재 Classes에는 다음 필드가 이미 있다.
- targetProgressCount
- targetProgressUnit

현재 운영 반 확인 결과:
- C-013: targetProgressCount=0, targetProgressUnit=과
- C-014: targetProgressCount=0, targetProgressUnit=빈 값

따라서 진도율(%) 표시는 targetProgressCount가 0보다 클 때만 계산한다.
운영 반의 실제 목표값 설정은 별도 단계에서 교사 확인 후 진행한다.

### 5.5 기존 Google 기반 JSLesson에서 확인한 장점
- 반 선택 시 학생별 과제 입력칸 자동 생성
- 오늘 날짜 자동 입력
- 수정 시 기존 수업/학생별 과제 다시 불러오기
- 저장 후 방금 저장한 행 강조
- 입력 중 초안 자동저장/복원
- classId로 내부 연결하고 화면에는 className 표시

### 5.6 기존 Google 기반 JSLesson에서 확인한 성능 단점
기존 화면은 수업 메뉴 진입 시 `getLessonsList()`로 전체 수업 목록을 가져온 뒤 브라우저에서 반 필터를 수행한다.
이 방식은 기록이 누적될수록 전송량/렌더링 시간이 증가할 수 있다.

V2에서는 서버에서 먼저 조건을 적용하고 최근 제한 건수만 반환한다.

### 5.7 기존 데이터 이전 판단
현재 V2 Lessons/LessonAssignments는 헤더만 있고 비어 있다.
반면 기존 Google 기반 화면에는 과거 수업 기록이 다수 존재한다.
따라서 과거 기록을 V2에서 계속 사용하려면 별도 이관이 필요하다.

이관 원칙:
- 기존 원본은 삭제하지 않는다.
- 먼저 원본 수업 시트/필드 구조를 정확히 확인한다.
- classId 매핑을 검증한 후 미리보기 → 복사 → 검증 순서로 진행한다.
- 이관은 Step 2 읽기 API 검증 후 별도 단계로 진행한다.

## 6. 데이터 구조 확정(현재 단계)

Lessons는 현재 헤더를 유지한다.
- lessonId
- date
- classId
- topic
- content
- homework
- nextPlan
- progressCurrentCount
- createdAt

표시 시 계산 필드(시트에 새 열 추가하지 않음):
- className ← Classes
- progressUnit ← Classes.targetProgressUnit
- targetProgressCount ← Classes.targetProgressCount
- progressPercent ← progressCurrentCount / targetProgressCount (target > 0일 때만)

## 7. 화면 구성 초안
### A. 수업 기록 입력
- 반
- 날짜
- 지난 수업 요약
- 주제
- 학습 내용
- 숙제
- 다음 계획
- 누적 진도
- 목표 진도 자동 표시
- 학생별 과제(접기/펼치기)
- 수업 저장
- 입력 초기화

### B. 저장된 수업 기록
- 반
- 시작일
- 종료일
- 최근 30건 / 전체 기간
- 조회
- 수정
- 삭제/보관

## 8. 단계별 개발 순서
### Step 1 — 현재 구조 확인 — 완료
- V2 실제 Lessons/LessonAssignments/Classes 확인 완료
- 기존 Google 기반 JSLesson 동작 확인 완료
- V2 Lessons/LessonAssignments가 비어 있음을 확인
- 기존 역사 데이터는 별도 이관이 필요함을 확인

### Step 2 — 읽기 전용 API — 완료
- 최근 30건 조회
- 반/기간 조회
- 지난 수업 1건 조회
- 전체 시트 대신 뒤쪽 청크 제한 스캔
- 화면 연결/저장 기능은 아직 하지 않음

### Step 3 — 저장/수정 API — 완료
- 신규 수업은 Lessons에 한 행만 append
- 수정은 lessonId로 찾은 해당 행만 갱신
- lessons.get / lessons.save API 추가
- 저장 시 LockService로 동시 ID 생성 충돌 방지
- LessonAssignments는 Step 5 전까지 저장하지 않음
- 삭제 정책은 아직 구현하지 않음

### Step 4 — V2 화면 연결 — 완료
- 수업 내용 메뉴 활성화 완료
- 기존 Google 화면 장점 유지
- 지난 수업/진도 자동 표시 완료
- 신규 저장/같은 lessonId 수정/최근 목록 조회 검증 완료
- 입력 초안 자동저장/복원 동작 확인

### Step 5 — 학생별 과제 연결 — 완료
#### Step 5A — API
- 수업별 학생 과제는 LessonAssignments에만 저장
- StudentDailyRecords는 개별 지도 기록으로 분리 유지
- lessonId + studentId 조합으로 1학생 1행 유지
- 신규 과제는 단일 행 append
- 수정은 기존 assignmentId 행만 update
- 빈 과제는 해당 학생의 LessonAssignments 행만 delete
- StudentDailyRecords에는 읽기/쓰기 모두 하지 않음
- 조회는 LessonAssignments 전체 로드 대신 lessonId 열 TextFinder 사용

#### Step 5B — 화면 연결 — 완료
- 학생별 과제 영역은 기본 접힘 상태의 선택 기능으로 제공
- 영역을 펼칠 때만 선택 반의 재학생을 렌더링
- 기존 수업 수정 시 영역을 펼칠 때만 lessonId 기준 과제를 지연 조회
- 과제를 입력한 경우에만 수업 저장 후 LessonAssignments API를 호출
- 학생별 과제를 전혀 사용하지 않으면 추가 API 호출/저장 없음
- 공통 수업 저장과 학생별 과제 저장을 분리하여 StudentDailyRecords 중복 저장 방지

### Step 6 — 성능 검증 — 완료
- Step 6A 실제 진단 완료: 구조상 lazy 로딩, 최근 30건/최대 100건, fullSheetRead=false, maxScanRows=4000, chunkSize=200 확인.
- 실제 참고 응답시간(2026-10-04): session 2535ms, lessons.list 30 3453ms, lessons.list 100 5072ms, lessons.previous 4524ms, 학생별 과제 조회 5871ms. 응답시간은 Apps Script/네트워크 영향을 받으므로 반복 중앙값으로 재확인.
- Step 6B: 반복 호출 중앙값과 100/500/1000/4000건 누적 가정으로 인덱스 도입 시점 판단.
- Step 6A: 현재 운영 데이터 기준 실제 API/화면 구조 진단
- 로그인/홈 초기 로딩에 Lessons 호출이 없는지 정적 구조 확인
- 수업 메뉴 진입 시에만 Lessons 조회가 연결되는지 확인
- lessons.list 기본 30건 / 최대 100건 제한 확인
- 전체 시트 대신 청크 제한 스캔(fullSheetRead=false) 확인
- 학생별 과제 조회도 fullSheetRead=false 확인
- API 응답시간은 참고값으로 기록하되 네트워크 상태 때문에 절대 실패 기준으로 사용하지 않음
- Step 6B: 100/500/1000건 이상 누적 가정 및 필요 시 인덱스 도입 여부 판단

## 9. 매 단계 점검 체크리스트
- [x] 기존 정상 기능을 건드리지 않는가? (Step 2는 읽기 전용만 추가)
- [x] 로그인 시 Lessons를 읽지 않는가?
- [x] 수업 메뉴 진입 전 Lessons API를 호출하지 않는가?
- [x] 기본 조회가 최근 제한 건수인가?
- [x] classId/날짜 조건을 지원하는가?
- [x] 저장 시 전체 시트를 다시 쓰지 않는가? (Step 3: 신규 append / 수정 단일 행)
- [x] 한 번에 너무 많은 행을 브라우저에 보내지 않는가?
- [x] 학생별 과제는 LessonAssignments로 분리하고 StudentDailyRecords와 중복 저장하지 않도록 Step 5A에서 확정
- [x] 반 식별은 classId 기준인가?
- [x] 진도 단위는 Classes에서 가져오도록 정리했는가?
- [x] 변경 사항을 이 문서에 기록했는가?
- [x] 학생별 과제 화면은 기본 접힘/필요 시 지연 조회인가? (Step 5B)

## 10. 변경 이력
- 2026-10-04: 최초 작성. 기존 Google 기반 수업 내용 화면 검토 결과와 장기 성능 원칙 반영.
- 2026-10-04: Step 1 실제 구조 확인 결과 반영. Lessons/LessonAssignments 현행 헤더 유지 결정. 기존 V2 시트가 비어 있고 과거 데이터 별도 이관이 필요함을 기록.
- 2026-10-04: Step 2 읽기 전용 API 설계 반영. 기본 30건, 200행 청크, 최대 4,000행 스캔, 최대 반환 100건으로 제한.

- 2026-10-04: Step 2 실제 진단 완료. count=0, previousLesson=null, readOnly=true, fullSheetRead=false 확인.
- 2026-10-04: Step 3 저장/수정 API 추가. 전체 시트 rewrite 금지, 신규 append, 수정 단일 행, 학생별 과제 저장은 보류.

- 2026-10-04: Step 4 실제 화면 검증 완료. 지난 수업 조회, 신규 L-00002 저장, 같은 lessonId 수정, 목록 2건 유지 확인.
- 2026-10-04: Step 5A 학생별 과제 API 추가. LessonAssignments 전용 저장, StudentDailyRecords 분리 유지, lessonId+studentId 중복 방지, 행 단위 생성/수정/삭제 원칙 반영.

- 2026-10-04: Step 5A 실제 진단 완료. LA-00001 생성/동일 ID 수정/빈 값 삭제, StudentDailyRecords 0건 유지, fullSheetRead=false 확인.
- 2026-10-04: Step 5B 선택형 학생별 과제 UI 연결. 기본 접힘, 펼칠 때 재학생 표시, 수정 시 지연 조회, 실제 변경 시에만 LessonAssignments 저장 호출.

- 2026-10-04: Step 5B 실제 화면 검증 완료. 선택형 학생별 과제 영역에서 기존 과제 1건 복원 및 재학생별 입력 표시 확인.
- 2026-10-04: Step 5C 세션 정리 안전성 보완. clearV2SessionState 재귀 호출 제거, 로그아웃 시 sessionStorage/localStorage의 V2 세션 토큰·현재 페이지를 함께 정리하도록 수정. 수업/과제 API 및 Apps Script는 변경하지 않음.

- 2026-10-04: Step 6A 성능 진단 파일 추가. 로그인/홈의 Lessons 비호출 구조, 최근 제한 조회, fullSheetRead=false, 학생별 과제 제한 조회를 실제 환경에서 점검하도록 함.

- 2026-10-04: Step 6A 실제 진단 정상. lazy 초기화, 메뉴 조건 조회, 최근 제한, fullSheetRead=false, StudentDailyRecords 비조회 확인. 실제 응답시간 참고값 기록.
- 2026-10-04: Step 6B 누적 데이터/인덱스 판단 진단 파일 추가. 운영 데이터는 수정하지 않고 반복 응답시간과 스캔 상한을 기준으로 인덱스 도입 시점 판단.


### Step 7 — 기존 Google 기반 수업 데이터 이관
#### Step 7A — 원본 읽기 전용 확인 — 완료
- 기존 Google 기반 `한국어학당 학생 관리 시스템`의 Lessons 188건 확인.
- 현재 V2 Classes에 실제 존재하는 C-013 19건, C-014 12건만 1차 이관 대상으로 확정.
- 종료된 과거 반 C-004~C-012의 157건은 반 정보 없는 고아 기록 방지를 위해 보류.
- 기존 LessonAssignments 15건은 모두 C-004/C-005/C-006에 속하므로 이번 1차 이관에서 제외.

#### Step 7B — 충돌/중복 미리보기 — 완료
- 현재 V2 Lessons 2건(L-00001, L-00002) 기준 검사.
- C-013/C-014 후보 31건과 lessonId 충돌 0건.
- 날짜/반/주제/학습내용/숙제/다음계획/진도 완전 동일 중복 0건.
- legacy lessonId L-159~L-189를 그대로 보존하기로 결정.

#### Step 7C — 실제 이관 도구 — 준비
- SUPER_ADMIN만 실행 가능.
- 원본 Spreadsheet ID와 승인 lessonId L-159~L-189를 코드에서 고정하여 예상 외 행을 가져오지 않는다.
- 실행 직전 원본 31건/반별 19+12건이 Step 7B 상태와 같은지 다시 검증한다.
- 실제 쓰기 직전에 현재 V2 `Lessons`를 숨김 백업 시트로 복제한다.
- 기존 원본은 읽기 전용이며 수정하지 않는다.
- 일반 최근 목록 API가 시트 끝에서 역방향 조회하므로, 과거 31건을 현재 최신 V2 행 뒤에 append하지 않는다.
- 대신 과거 기록을 날짜 오름차순으로 정리하여 `Lessons`의 2행 앞쪽에 삽입하고 현재 V2 최신 행은 아래쪽에 유지한다.
- 전체 시트 rewrite는 하지 않는다.
- 이미 같은 lessonId 또는 완전 동일 기록이 존재하면 자동 skip한다.
- 실행 후 총건수, 승인 ID 31건, C-013 19건/C-014 12건을 재검증한다.
- 확인 문구 `IMPORT_C013_C014_31` 없이는 실제 이관이 실행되지 않는다.

### Step 7C 주의사항
- Step 7C Apps Script 파일을 적용하면 새 버전 재배포가 필요하다.
- 먼저 미리보기만 실행하고 결과가 `readyCount=31`, 충돌 0건인지 확인한 후 실제 이관 버튼을 사용한다.
- 실제 이관 후 재실행하면 동일 lessonId가 이미 있으므로 31건이 skip되어 중복 삽입되지 않아야 한다.
- 이관 후 신규 수업의 자동 lessonId는 기존 보존 ID의 최대 숫자를 기준으로 다음 번호가 생성될 수 있다.

## Step 7 변경 이력 추가
- 2026-10-04: Step 7A 원본 분석 완료. 전체 188건 중 현재 V2 반과 직접 연결되는 C-013/C-014 31건만 1차 이관 대상으로 확정.
- 2026-10-04: Step 7B 충돌 검사 완료. V2 기존 2건과 ID 충돌/완전 동일 중복 모두 0건.
- 2026-10-04: Step 7C 안전 이관 도구 준비. 백업 후 과거 기록을 현재 최신 V2 행 위쪽에 삽입하고 검증하도록 설계.

## 11. Step 7 — 기존 Google 기반 수업 기록 이관
### Step 7A/7B — 읽기 전용 분석/충돌 검사 — 완료
- 기존 Google 기반 Lessons 188건 확인.
- 현재 V2에서 그대로 유지되는 C-013 19건, C-014 12건, 총 31건을 1차 이관 대상으로 확정.
- lessonId 충돌 0건, 완전 동일 중복 0건 확인.
- 과거 종료 반 C-004~C-012의 157건은 현재 V2 Classes에 대응 반이 없어 이관 보류.
- 기존 LessonAssignments 15건은 모두 과거 종료 반(C-004/C-005/C-006)에 속해 1차 이관에서 제외.

### Step 7C — 실제 이관 — 완료
- 승인된 L-159~L-189 31건을 원본 lessonId/createdAt을 보존하여 이관.
- 기존 V2 최신 기록이 최근 목록에서 우선되도록 과거 31건을 row 2 앞쪽에 삽입.
- 이관 직전 숨김 백업 시트 `Lessons_BACKUP_STEP7C_...` 자동 생성.
- 실제 검증 결과: Lessons 2건 → 33건, C-013=21건, C-014=12건.
- 원본 Google 기반 스프레드시트는 수정하지 않음.

### Step 7D — 실제 화면 검증 — 완료
- C-014에서 2026-09-30 `토픽 듣기 연습` 포함 과거 목록 정상 조회.
- C-013에서 2026-09-30 `7과 날씨가 어떻습니까?` 포함 과거 목록 정상 조회.
- 이관된 데이터가 기존 V2 수업 목록 UI에서 반별로 정상 표시됨.

### Step 7E — 개발용 테스트 기록 정리 — 완료
- 개발 과정에서 생성한 `L-00001`(Step 3 진단), `L-00002`(실제 저장/수정 테스트)만 제거 대상.
- `L-00002`에 연결된 LessonAssignments 테스트 과제도 함께 제거.
- 실제 삭제 전에 Lessons와 LessonAssignments를 각각 숨김 백업 시트로 복사.
- ID와 주제의 테스트 표식이 모두 일치할 때만 정리 실행.
- 실제 최종 상태 확인: Lessons 31건, C-013=19건, C-014=12건, 테스트 lessonId 0건, LessonAssignments 0건.
- 정리 직전 Lessons와 LessonAssignments 숨김 백업 시트 생성 확인.

## 12. Step 8 — 수업 기록 보관/복원 정책
### Step 8A — 보관/복원 기능
- 영구 삭제는 구현하지 않는다. 기본 관리 기능은 `보관`으로 통일한다.
- 활성 `Lessons` 행을 보관할 때 `LessonsArchive`로 복사한 뒤 활성 행만 제거한다.
- 연결된 `LessonAssignments`가 있으면 `LessonAssignmentsArchive`로 함께 이동한다.
- 보관 시 전체 시트 rewrite를 하지 않고 해당 행만 이동한다.
- 보관 목록은 기본 30건, 최대 100건, 200행 청크/최대 4,000행 스캔 원칙을 유지한다.
- 보관된 수업은 `복원`할 수 있으며, 원래 수업 날짜를 기준으로 활성 `Lessons`의 시간 순서에 맞춰 다시 삽입한다.
- 복원 시 연결된 학생별 과제도 함께 복원한다.
- 같은 lessonId가 활성 `Lessons`에 이미 있으면 복원을 차단한다.
- 보관/복원은 로그인/홈 초기 로딩에 영향을 주지 않고 수업 내용 화면에서만 호출한다.

## Step 8 변경 이력
- 2026-10-04: Step 7E 완료 후 운영 데이터는 Lessons 31건(C-013 19건, C-014 12건), 테스트 LessonAssignments 0건 상태로 정리됨.
- 2026-10-04: Step 8A 보관/복원 정책 확정. 영구 삭제 대신 LessonsArchive/LessonAssignmentsArchive를 이용한 행 단위 보관/복원을 도입.

