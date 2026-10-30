# 고천중 가을밤 천체관측 신청 페이지

- 학생·학부모용 페이지: https://syjh0412.github.io/stargazing/
- 접수 창구와 기록: Google Apps Script(`Code.gs`) + 구글 시트 '신청' 탭

## 파일
- `index.html` — GitHub Pages로 보이는 신청 페이지 (접수는 Apps Script 주소로 보냄)
- `Code.gs` — Apps Script에 붙여넣는 서버 코드
- `apps-script-Index.html` — Apps Script 쪽 `Index` 파일 (index.html과 같은 내용, Apps Script 주소로 직접 열 때 쓰임)

## 고칠 때
- 신청 시작·마감·정원: `Code.gs` 맨 위 `EVENT`
- Apps Script를 고친 뒤에는 **배포 → 배포 관리 → 연필 → 새 버전**으로 배포 (주소 유지)
