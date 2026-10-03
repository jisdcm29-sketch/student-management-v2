# STUDENT MANAGEMENT V2 - 통합 성적 허브 시방서 V1.0
작성일: 2026-10-03

## 1. 목적
성적 관리는 교사가 수동 입력하는 시험 점수만 표시하는 화면이 아니라, 학생이 한국어 모바일/평가/TOPIK 프로그램에서 남긴 학습 성취와 활동 기록을 한 곳에서 확인하는 통합 성적 허브로 만든다.

핵심 원칙:
- 원본 프로그램의 점수/활동 데이터를 지우거나 옮기지 않는다.
- 학생관리 V2는 원본을 읽어 통합 표시하는 것을 우선한다.
- 기존 Scores 시트는 교사가 직접 입력하는 공식 성적/평가 기록용으로 유지한다.
- 외부 프로그램 데이터는 phone을 연결키로 사용하되, 학생관리 V2의 studentId를 최종 내부 식별자로 사용한다.
- 자동 활동 데이터만으로 학생의 태도를 단정하거나 자동 등급화하지 않는다.
- 활동/참여 지표와 교사의 태도 메모를 분리한다.

## 2. 확인된 실제 데이터 소스
### A. 학생관리 V2
- Students: studentId, name, phone, classId, status
- Attendance: 출석/지각/결석/조퇴
- Scores: 교사 직접 입력 성적
- DailyRecords/수업 기록: 교사 관찰 기록

### B. 한국어 모바일 프로그램 이용 기록
확인된 시트:
- TestResults
  - date, phone, name, klass, book, lesson, testType
  - bestScore, attemptsToday, bestCorrect, total, bestTimeout
  - firstAt, bestAt, lastAt, status
- 워크북읽기평가
  - 교재, 복습, 점수, 정답수, 전체문항, 미응답, 응시시간, 시간초과, 답안
- 워크북듣기평가
  - 교재, 복습, 상태, 점수, 정답수, 전체문항, 응시시간
- 진도현황
  - 최초접속, 최근접속, 시작교재/과, 현재교재 등
- Sessions / Log
  - 로그인/세션/앱 활동 이벤트
- 학생별 시트
  - SNU 단원 테스트, 워크북 평가, TOPIK 평가를 학생별로 누적

### C. TOPIK I 읽기
- 별도 응시기록의 All_Results
- 실전, 문항연습, 오답복습을 구분
- attempt_id, phone, student_name, result_type, mode, 회차, 시험명
- 응시시간, 문항수, 정답/오답/미응답, 점수, 문항 번호 목록

### D. TOPIK I 듣기
- 이번 1차 성적 허브 범위에서 제외한다.
- TOPIK I 듣기 앱이 운영에 포함되는 시점에 별도 연결한다.

### E. TOPIK II 읽기/듣기/쓰기
- 이번 1차 성적 허브 범위에서 제외한다.
- 서울대/워크북/TOPIK I 통합이 안정화된 뒤 별도 단계에서 연결한다.
- 현재 화면과 API는 TOPIK II 데이터가 없어도 정상 동작하도록 설계한다.

## 3. 데이터 통합 방식
원본 시트를 복사하여 Scores에 섞지 않는다.
서버에서 읽은 데이터를 공통 포맷으로 변환하여 화면에 제공한다.

공통 이벤트 예:
- source: SNU_VOCAB / SNU_GRAMMAR / SNU_MIXED / REVIEW_READING / REVIEW_LISTENING / TOPIK1_COLLOCATION / TOPIK1_GRAMMAR / TOPIK1_READING
- studentId
- phone
- occurredAt
- program
- book
- unit
- activityType
- mode
- score100
- correct
- total
- unanswered
- attempts
- durationSec
- status
- sourceAttemptId

중복키:
source + sourceAttemptId

## 4. 학생 성적 화면 구성
### 상단: 학생 요약
- 학생 / 반 / 재학상태
- 최근 학습일
- 현재 교재/과
- 출석률
- 최근 30일 학습 활동 횟수

### 핵심 성취 카드
- 서울대 각 과 어휘 테스트
- 서울대 각 과 문법 테스트
- 서울대 각 과 종합 테스트
- 복습 읽기 평가
- 복습 듣기 평가
- TOPIK I 연어 시험
- TOPIK I 문법 시험
- TOPIK I 읽기평가
각 카드는 최신점수, 최고점수, 최근 응시일, 응시횟수를 보여준다.
TOPIK I 듣기와 TOPIK II는 후속 단계에서 추가한다.

### 세부 탭
1. 종합
2. 서울대 교재
3. 복습 읽기·듣기
4. TOPIK I
5. 학습 활동
6. 교사 기록

## 5. 반 전체 화면
교사가 반을 선택하면 한 화면에서 학생별 핵심 지표를 본다.
권장 열:
- 학생
- 현재 교재/진도
- 서울대 최근 성취
- 워크북 읽기
- 워크북 듣기
- TOPIK I
- 최근 활동
- 출석
- 교사 확인

점수 한 칸에 모든 것을 압축하지 않고, 영역별 요약을 제공한다.
학생 이름을 누르면 개인 상세로 진입한다.

## 6. 활동/태도 처리 원칙
'태도'는 자동 점수로 만들지 않는다.
객관적 활동 근거를 별도로 보여준다:
- 최근 접속일
- 학습일 수
- 시험/연습 응시 횟수
- 재시도 횟수
- 오답 복습 여부
- 미응답/시간초과 기록
- 진도 진행
- 출석 기록

교사는 이 근거를 보고 기존 attitudeNote/teacherNote 또는 별도 교사 기록에 판단을 남긴다.

## 7. 교사 편의 UX
기존 Google 기반 성적 화면의 장점은 유지한다:
- 반 → 학생 순서 선택
- 날짜 기본값
- 성적 수정/조회가 한 화면에서 가능
- 최근 성적 바로 확인

개선:
- 수동 입력보다 자동 수집 성적이 기본 화면의 중심
- 반 전체 요약 → 학생 상세 드릴다운
- 탭별로 데이터 종류 분리
- 긴 원시 로그는 기본 화면에서 숨기고 필요할 때 펼치기
- 필터: 기간 / 교재 / 평가유형 / 프로그램
- 최근 30일을 기본 기간으로 하되 전체 기간 선택 가능

## 8. 성능 원칙
- 로그인 시 모든 외부 시트를 한 번에 읽지 않는다.
- 성적 메뉴 진입 시 필요한 데이터만 로딩한다.
- 반 전체는 요약 데이터만 반환한다.
- 학생 상세를 열 때 해당 학생 데이터만 추가 조회한다.
- 외부 데이터 소스별 어댑터를 분리한다.
- CacheService 또는 요약 캐시를 사용하되 원본 수정은 하지 않는다.
- 데이터가 커지면 source + phone + date 기반 인덱스를 추가 검토한다.

## 9. 권장 API
- scoreHub.sources
- scoreHub.classSummary
- scoreHub.studentSummary
- scoreHub.studentTimeline
- scoreHub.manual.list
- scoreHub.manual.save
- scoreHub.manual.delete

기존 단순 scores.* API는 교사 직접 입력용으로만 제한하고, 통합 허브의 메인 API로 사용하지 않는다.

## 10. 단계별 구현 순서
1. 데이터 소스 카탈로그 확정
2. 학생 전화번호 <-> studentId 연결 검증
3. 한국어 모바일 TestResults/워크북/진도 어댑터
4. TOPIK I 연어/문법/읽기 어댑터
5. 반 전체 요약 API
6. 학생 상세 API
7. 성적 관리 테스트 화면
8. 교사 사용성 검증
9. 운영 반영
10. TOPIK I 듣기와 TOPIK II는 후속 단계에서 별도 연결

## 11. 현재 중단/보류 사항
이전에 준비한 단순 ScoreServiceV2(읽기/쓰기/듣기/말하기 수동 점수 중심)는 최종 성적 화면의 기반으로 사용하지 않는다.
기존 Scores 시트의 수동 입력 기능만 별도 영역에서 재사용한다.

## 12. 데이터 해석 주의
접속 횟수, 페이지 열람, 사용시간 같은 활동 데이터는 학업 성취나 태도를 직접 의미하지 않는다.
따라서 시스템은 사실 데이터와 경향을 표시하고, 교사의 최종 판단을 대신하지 않는다.


## 13. 진도 및 응시 참고자료
학생 상세 상단에 현재 학습 위치와 응시량을 함께 표시한다.
- 시작교재/시작과
- 현재교재/현재과
- 현재 과 통과현황
- 현재 과 어휘/문법/종합 최고점
- 현재상태 및 다음단계
- TOPIK I 연어/문법 진행상태
- 최근활동 시각
- 최근시험 및 최근점수
- 총응시 횟수

서울대 및 TOPIK I 연어/문법 시험의 응시수는 TestResults의 행 개수가 아니라 attemptsToday의 합계를 사용한다.
복습 읽기/듣기와 TOPIK I 읽기는 제출/응시 기록 1행을 1회 응시로 계산한다.
