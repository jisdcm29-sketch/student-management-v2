# 학생관리시스템 V2 Apps Script API

이 폴더는 V1 운영 Apps Script와 완전히 분리된 V2 API 프로젝트용 전체 코드입니다.

## 적용 순서
1. script.google.com에서 새 프로젝트를 만듭니다.
2. 프로젝트 이름을 `학생관리시스템_V2_API`로 지정합니다.
3. 이 폴더의 .gs 파일과 appsscript.json 내용을 그대로 만듭니다.
4. 프로젝트 설정 > 스크립트 속성에 다음 값을 추가합니다.
   - `SETUP_ADMIN_PASSWORD` = 본인이 사용할 새 V2 관리자 비밀번호
5. `setupInitialAdminPasswordFromProperty_` 함수를 한 번 실행합니다.
6. 실행이 성공하면 SETUP_ADMIN_PASSWORD 속성은 코드가 자동 삭제합니다.
7. 배포 > 새 배포 > 웹 앱
   - 실행 사용자: 나
   - 액세스 권한: 모든 사용자(로그인하지 않은 사용자 포함 가능)
8. 생성된 /exec URL을 GitHub의 `src/config.js`에 입력합니다.
9. GitHub Pages 테스트 화면에서 API 상태 확인 → 로그인 테스트를 진행합니다.

## 중요
- V1 Apps Script 프로젝트에 이 코드를 붙여넣지 마십시오.
- V1 Google Sheet를 수정하지 않습니다.
- 실제 학생 데이터는 아직 V2로 복사하지 않습니다.
