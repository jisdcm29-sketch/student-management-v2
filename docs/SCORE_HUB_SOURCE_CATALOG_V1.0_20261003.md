# SCORE HUB SOURCE CATALOG V1.0
작성일: 2026-10-03
범위: 서울대 각 과 어휘·문법·종합 + 복습 읽기·듣기 평가 + TOPIK I 연어·문법·읽기평가
TOPIK I 듣기 및 TOPIK II: 이번 단계에서 제외하고 추후 연결.

## 학생 연결키
학생관리 V2 Students의 studentId를 내부 기준키로 사용하고, 외부 프로그램은 phone exact match로 연결한다. 이름만으로 자동 연결하지 않는다.

## 서울대/모바일
Spreadsheet: 한국어 모바일 프로그램이용 기록
ID: 1y4xaZD8SQUVLztDhBSytvi-_naVCGYX5gTOyZqmhMUE

- TestResults: 서울대 단원 어휘/문법/종합 성취, 최고점, 응시횟수, PASS/RETRY
- 워크북읽기평가: 점수, 정답수, 미응답, 응시시간, 시간초과
- 워크북듣기평가: SUBMITTED 결과의 점수, 정답수, 응시시간
- 진도현황: 최초/최근 접속, 시작교재/과, 현재교재
- 학생별 시트: 상세 검증용. 1차 집계에서는 중복 방지를 위해 원본 집계 소스로 사용하지 않음.

## TOPIK I 읽기
Spreadsheet: TOPIK1_읽기_응시기록
ID: 18HXty992Riii2-csrB2aHpQ7vVt8qD2NOWFMp1yG63M
Sheet: All_Results

- FULL_FIXED 등: 실전 성취
- QUESTION_PRACTICE: 유형 연습
- WRONG_REVIEW: 오답 복습
세 범주를 서로 섞어 단일 평균으로 만들지 않는다.

## TOPIK I 듣기
Spreadsheet: TOPIK1_듣기_응시기록
ID: 1F4Bpcb4tIwuRMRy7aLb9j1YLxzxCg_wWyYEiLxkg798
Sheet: All_Results
TOPIK I 읽기와 동일 계열 스키마로 실전/연습/오답복습을 분리 집계한다.

## 1차 응답 구조
student, progress, snu, workbookReading, workbookListening, topik1Collocation, topik1Grammar, topik1Reading, activity

## 태도 처리
자동 태도점수는 만들지 않는다. 최근 접속, 학습 이벤트, 재시도, 오답복습, 미응답, 시간초과, 진도, 출석 등 객관적 근거와 교사 메모를 분리한다.

## 제외
TOPIK I 듣기 및 TOPIK II 읽기/듣기/쓰기: 추후 연결.
