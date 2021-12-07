// 最初のみ起動させる定期実行のトリガー
// 毎日AM9:00にシートを生成する
function setActivateTrigger() {
    ScriptApp.newTrigger('main').timeBased().atHour(9).everyDays(1).create();
}

// デフォルトのシートをコピー
function copyDefaultSheet(mySheet, defaultSheetName) {
    let defaultSheet = mySheet.getSheetByName(defaultSheetName);
    defaultSheet.copyTo(mySheet);
}

// コピーしたシート名を現在日時に変更する
function changeSheetNameToCurrentDate(mySheet, defaultSheetName) {
    let date = new Date();
    let month = date.getMonth() + 1;
    let day = date.getDate();

    let targetSheetName = `${defaultSheetName} のコピー`;

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
    const defaultSheetName = "default";
    const mySheet = SpreadsheetApp.getActiveSpreadsheet();
    copyDefaultSheet(mySheet, defaultSheetName);
    changeSheetNameToCurrentDate(mySheet, defaultSheetName);
}
