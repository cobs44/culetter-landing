// ─────────────────────────────────────────────
// Culetter Beta Application - Google Apps Script
// ─────────────────────────────────────────────
// 이 스크립트는 베타 신청 폼에서 받은 데이터를
// 현재 Google Sheet의 "Beta Applications" 탭에 한 줄씩 추가합니다.
//
// 세팅 방법:
// 1. Google Sheets에서 새 시트를 만든다
// 2. Extensions > Apps Script 메뉴 클릭
// 3. 기본 코드(myFunction)를 모두 지우고 이 파일 내용을 붙여넣는다
// 4. 디스크 아이콘으로 저장 (이름 아무거나)
// 5. 우측 상단 [Deploy] > [New deployment]
// 6. ⚙️(톱니) 클릭 > "Web app" 선택
// 7. 다음과 같이 설정:
//      - Description: 아무거나 (예: "Culetter Beta v1")
//      - Execute as: Me (본인)
//      - Who has access: Anyone (필수 — 폼이 외부에서 호출됨)
// 8. [Deploy] 클릭 > 권한 승인 (구글 계정 로그인 → "고급" → "안전하지 않은 페이지로 이동" → 허용)
// 9. 발급된 "Web app URL"을 복사
// 10. beta.html에서 SCRIPT_URL 변수에 그 URL을 붙여넣기
// ─────────────────────────────────────────────

const SHEET_NAME = 'Beta Applications';

function doPost(e) {
  try {
    // 동시 요청 방지 (50명 정도면 거의 발생 안 함)
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    try {
      const data = JSON.parse(e.postData.contents);
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName(SHEET_NAME);

      // 시트가 없으면 새로 만들기
      if (!sheet) {
        sheet = ss.insertSheet(SHEET_NAME);
      }

      // 헤더가 비어있으면 한 번만 추가
      if (sheet.getLastRow() === 0) {
        sheet.appendRow([
          'Submitted At',
          'Name',
          'Email',
          'Moment',
          'Frequency',
          'Agreed'
        ]);
        // 헤더 스타일링
        sheet.getRange(1, 1, 1, 6)
          .setFontWeight('bold')
          .setBackground('#F0F0F0');
        sheet.setFrozenRows(1);
      }

      // 한국 시간으로 타임스탬프
      const kstTime = Utilities.formatDate(
        new Date(),
        'Asia/Seoul',
        'yyyy-MM-dd HH:mm:ss'
      );

      // 새 행 추가
      sheet.appendRow([
        kstTime,
        data.name || '',
        data.email || '',
        data.moment || '',
        data.frequency || '',
        data.agree ? 'Yes' : 'No'
      ]);

      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);

    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// (선택) 브라우저로 URL을 직접 열었을 때 노출되는 페이지
function doGet() {
  return ContentService
    .createTextOutput('Culetter Beta endpoint is live.')
    .setMimeType(ContentService.MimeType.TEXT);
}
