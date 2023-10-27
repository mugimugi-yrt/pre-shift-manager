// シートをグローバル化
var targetSpreadsheetId = "Sheet ID"; // 目標のスプレッドシートのIDを指定
var targetSheetName = "Sheet Name"; // 目標のシート名を指定
var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);

function onFormSubmit(e) {
  var formResponses = e.values;

  // シフト希望：Trueにすることで希望のシフトに入る
  // shiftdata[i][0] = 月曜日のシフト
  var shiftData = [
    [false, false, false, false, false], // 2限
    [false, false, false, false, false], // L限
    [false, false, false, false, false], // 3限
    [false, false, false, false, false], // 4限
    [false, false, false, false, false]  // 5限
  ];

  // データ抽出
  var responseData = [
    formResponses[5].split(','),  // 2限
    formResponses[6].split(','),  // L限
    formResponses[7].split(','),  // 3限
    formResponses[8].split(','),  // 4限
    formResponses[9].split(',')   // 5限
  ]

  // シフト希望格納
  shiftData = updateShiftData(shiftData, responseData);
  // Logger.log(shiftData);  // デバック用

  // 表で表示されるスタンプを作成
  var nameL = formResponses[2];
  var same = formResponses[4];
  var stamp = nameL;
  if (same === "yes") {
    var nameF = formResponses[3];
    stamp = stamp + "(" + nameF.charAt(0) + ")";
  }
  
  // 学年をで番号割り振り
  var studentYear = formResponses[1];
  var yearnum = giveNumber(studentYear);
  
  // シフト表作成
  makeSheet(shiftData, yearnum, stamp);
}

// 個人シフト表リストとシフト希望回答から、個人シフト表リストを更新
function updateShiftData(data, respose) {
  for (var i = 0; i < 5; i++) {
    if (respose[i].length !== 0) {
      // Logger.log("hoge");  // デバック用
      for (var j = 0; j < respose[i].length; j++) {
        // Logger.log("hogehoge");  // デバック用
        // resposeに格納されているものの空白文字を消去
        var request = respose[i][j].trim();
        if (request === "月曜日")      { data[i][0] = true; }
        else if (request === "火曜日") { data[i][1] = true; }
        else if (request === "水曜日") { data[i][2] = true; }
        else if (request === "木曜日") { data[i][3] = true; }
        else if (request === "金曜日") { data[i][4] = true; }
      }
    }
  }
  return data;
}

// 学年から数値を返す
function giveNumber(year) {
  if (year === "3")      { return 2; }
  else if (year === "2") { return 3; }
  else if (year === "1") { return 4; }
}

// 全体シフト表のシート作成
function makeSheet(data, num, moji) {
  var column = ["C", "D", "E", "F", "G"];
  for (var i = 0; i < 5; i++) {
    var x = num + (i * 3);
    for (var j = 0; j < 5; j++) {
      if (data[i][j]) {
        var cell = column[j] + String(x);
        var targetCell = targetSheet.getRange(cell);
        var currentText = targetCell.getValue();
        if(currentText) {
          var updateText = currentText + ", " + moji;
          targetCell.setValue(updateText);
        }
        else {
          targetCell.setValue(moji);
        }
      }
    }
  }
}
