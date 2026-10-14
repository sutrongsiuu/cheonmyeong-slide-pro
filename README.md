# 천명 프로그램 (CheonMyeong Slide Pro)

예배 슬라이드 프레젠테이션 앱입니다.

## 다운로드

**[최신 버전 다운로드 (Releases)](../../releases/latest)**

1. 위 링크에서 `CheonMyeong-x.x.x-Setup.exe` 를 내려받습니다.
2. 실행하여 설치합니다.
   - Windows에서 "PC 보호" 경고가 뜨면 **추가 정보 → 실행**을 누르세요. (코드 서명이 없는 앱이라 표시되는 정상 경고입니다)

## AI 기능 사용

설정 화면에서 본인의 AI API 키(Gemini / OpenAI / Claude)를 입력하면 AI 기능을 사용할 수 있습니다.

## 개발자용: 소스에서 실행 / 빌드

```bash
npm install
npm start          # 실행
npm run build-win  # Windows 설치파일 빌드 (dist/ 폴더에 생성)
```

### 새 버전 배포 (자동 업데이트)

1. `package.json`의 `version`을 올립니다. (예: 1.5.1 → 1.6.0)
2. 아래 명령으로 빌드와 동시에 GitHub Releases에 업로드합니다.
   ```powershell
   $env:GH_TOKEN = gh auth token
   npm run release
   ```
3. 이미 설치한 사용자는 프로그램을 켤 때 새 버전을 자동으로 내려받고, 재시작하거나 종료할 때 업데이트됩니다.
   (메뉴 **Cheon Myeong Slide Pro → 업데이트 확인**으로 직접 확인할 수도 있습니다.)

## 라이선스

MIT
