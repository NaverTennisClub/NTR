const SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();
const scoreInputSheet = SPREADSHEET.getSheetByName('스코어 입력');
const NTRScoreSheet = SPREADSHEET.getSheetByName('NTR 점수');

/**
 * 스프레드시트 열릴 때 메뉴 추가
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('NTR 시스템')
    .addItem('NTR 계산하기', 'calculateNTR')      // 새로 추가된 행만 처리
    .addItem('NTR 전체 재계산', 'recalcAllNTR')   // 시트1의 모든 경기 재계산
    .addItem("점수 입력창 열기", "openNTRDialog")
    .addToUi();

    openNTRDialog()
}

function openNTRDialog() {
  const html = HtmlService.createTemplateFromFile("dialog").evaluate()
    .setWidth(700)      // 가로 크기 (px)
    .setHeight(600);    // 세로 크기 (px)
  SpreadsheetApp.getUi().showModalDialog(html, "데이터 입력");
}

function doPost(e) {
  try {
    Logger.log("doPost() 호출됨"); // 로그 기록

    if (!e || !e.postData || !e.postData.contents) {
      Logger.log("❌ 데이터가 전달되지 않음");
      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "No data received" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    Logger.log("✅ 데이터 수신됨: " + e.postData.contents); // JSON 데이터 로그

    const data = JSON.parse(e.postData.contents);

    // 필수 값 검증
    if (!data.p1 || !data.p2 || !data.p3 || !data.p4 || !data.score) {
      Logger.log("❌ 필수 필드 누락됨");
      return ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "Missing required fields" })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    insertDataToSheet(data);
    Logger.log("✅ 데이터 스프레드시트에 저장됨");

    return ContentService.createTextOutput(
      JSON.stringify({ status: "success", message: "Data inserted successfully" })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    Logger.log("❌ 오류 발생: " + error.message);
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: error.message })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
