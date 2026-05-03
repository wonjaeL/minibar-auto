# MinibarAuto

호텔 미니바 일일 판매 엑셀에서 특정 날짜의 Room No 리스트를 추출하고, finjrnl XML 트랜잭션 로그를 정렬/필터 가능한 표로 보여주는 Electron 앱.

## Windows 실행파일 다운로드

1. [Releases](../../releases) 페이지에서 최신 `MinibarAuto-<version>-portable.exe` 다운로드
2. 더블클릭으로 바로 실행 (설치 불필요)

릴리스가 비어 있으면 [Actions](../../actions) 탭의 최신 `build-windows` 워크플로우 결과 → `MinibarAuto-windows-portable` artifact에서도 받을 수 있음.

## 기능

### Excel → Room List
- `7.Daily Minibar Sold_YYYY_MM.xlsx` 첨부 → 날짜 선택 → 콤마 구분 Room No 문자열 출력
- 3자리 → 4자리 zero-pad (예: 601 → 0601)
- 분할 시트(`17 (1)`, `17 (2)`) 자동 병합

### XML Viewer
- `finjrnl_articles_*.XML` 첨부 → 모든 트랜잭션 평탄화
- 컬럼 체크박스로 표시 토글
- 헤더 클릭 정렬 (asc↔desc, 숫자/문자열 자동 판별)
- 컬럼별 substring 필터 (대소문자 무시)

## 개발

```bash
npm install
npm start          # Electron 실행
npm run dist       # 로컬에서 Windows .exe 빌드 (Wine 필요할 수 있음)
```

## 릴리스 만들기

```bash
git tag v0.1.0 && git push origin v0.1.0
```

GitHub Actions가 Windows 러너에서 빌드 후 자동으로 Release에 `.exe` 첨부.
