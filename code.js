// 最初のみ起動させる定期実行のトリガー
// 毎日AM9:00にシートを生成する
function setActivateTrigger() {
    ScriptApp.newTrigger('main').timeBased().atHour(9).everyDays(1).create();
}

// 最新のシートをコピー
function copySheet() {
  let mySheet = SpreadsheetApp.getActiveSpreadsheet();
  let copySheet = mySheet.getActiveSheet();
  copySheet.copyTo(mySheet);
}

// xxのコピーのコピーを削除する
function changeSheetNameToCurrentDate() {
    let mySheet = SpreadsheetApp.getActiveSpreadsheet();
    let date = new Date();

    let month = date.getMonth() + 1;
    let day = date.getDate();
    let yesterday = date.getDate() - 1;
    let targetSheetName = `${month}/${yesterday} のコピー`;

    // コピーしたシート名を現在日時に変更する
    let dateString = `${month}/${day}`
    let targetSheet = mySheet.getSheetByName(targetSheetName);
    targetSheet.setName(`${dateString}`);

    // コピーしたシートを先頭に持ってくる
    moveFirstSheet(mySheet, targetSheet);
}

// コピーしたシートを先頭に持ってくる
function moveFirstSheet(mySheet, targetSheet) {
    targetSheet.activate();
    mySheet.moveActiveSheet(1);
}

// main
function main() {
    copySheet();
    changeSheetNameToCurrentDate();
}
