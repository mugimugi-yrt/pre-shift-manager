// シートをグローバル化
var targetSpreadsheetId = "Sheet ID"; // 使用スプレッドシートのID
var targetSheetName = "Sheet Name"; // 調整表作成シート名を指定
var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
// ここがフォーム受信時に実行されるmain関数
// フォームが提出された時のシート作成(要フォーム受信トリガー)
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
    formResponses[4].split(','),  // 2限
    formResponses[5].split(','),  // L限
    formResponses[6].split(','),  // 3限
    formResponses[7].split(','),  // 4限
    formResponses[8].split(',')   // 5限
  ]

  // シフト希望格納
  shiftData = updateShiftData(shiftData, responseData);

  // 表で表示されるスタンプを作成
  var nameL = formResponses[2];
  var nameF = formResponses[3];
  var nameFull = nameL + ' ' + nameF;
  
  // 学年をで番号割り振り
  var studentYear = formResponses[1];
  var yearnum = giveNumber(studentYear);

  // 名簿表更新
  updateNameList(studentYear, nameFull);

  // 備考表記入
  var note = formResponses[9];
  if (note !== "") { updateNote(note, nameFull); }

  // シフト表作成
  makeSheet(shiftData, yearnum, nameL);
  makePulldown();
}

// 個人シフト表リストとシフト希望回答から、個人シフト表リストを更新
function updateShiftData(data, respose) {
  for (var i = 0; i < 5; i++) {
    if (respose[i].length !== 0) {
      for (var j = 0; j < respose[i].length; j++) {
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

// 名簿リストを更新
function updateNameList(year, name) {
  var targetWord = [
    ["I", "J", "K"],  // 3年
    ["L", "M", "N"],  // 2年
    ["O", "P", "Q"]   // 1年
  ]
  if (year === "3")      { var num = 0; }
  else if (year === "2") { var num = 1; }
  else if (year === "1") { var num = 2; }

  // 提出人数更新
  var grade = targetSheet.getRange(targetWord[num][1] + "3");
  var cellValue = grade.getValue().split(' ');
  var submitNum = parseInt(cellValue[0], 10) + 1;
  grade.setValue(String(submitNum) +"人");

  // 空欄部分を調査
  var n = 5;
  var targetCell = targetSheet.getRange(targetWord[num][0] + String(n));
  var usedChecker = targetCell.getValue();
  while(usedChecker) {
    n = n + 1;
    targetCell = targetSheet.getRange(targetWord[num][0] + String(n));
    usedChecker = targetCell.getValue();
  }

  // 空欄部分に新たな提出者情報を挿入
  for (var i = 0; i < 3; i++) {
    targetCell = targetSheet.getRange(targetWord[num][i] + String(n));
    if (i === 0) {
      targetCell.setValue(name);
    }
    else {
      targetCell.insertCheckboxes();
    }
  }
}

// 備考リストを更新
function updateNote(sentence, name) {
  // 空欄部分を調査
  var n = 24;
  var targetCell = targetSheet.getRange("I" + String(n));
  var usedChecker = targetCell.getValue();
  while(usedChecker) {
    n = n + 1;
    targetCell = targetSheet.getRange("I" + String(n));
    usedChecker = targetCell.getValue();
  }

  // 空欄部分に新たな備考情報を挿入
  targetCell.setValue(name);
  targetCell = targetSheet.getRange("J" + String(n));
  targetCell.setValue(sentence);
}

// 学年から数値を返す
function giveNumber(year) {
  if (year === "3")      { return 3; }
  else if (year === "2") { return 4; }
  else if (year === "1") { return 5; }
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
          var updateText = currentText + ', ' + moji;
          targetCell.setValue(updateText);
        }
        else {
          targetCell.setValue(moji);
        }
      }
    }
  }
}

// 決定版にプルダウンを埋め込む
function makePulldown() {
  var shiftDataIn = ["C", "D", "E", "F", "G"];
  // タテで確認
  for (var i = 0; i < shiftDataIn.length; i++) {
    var n = 0
    for (var j = 3; j <= 17; j++) {
      var num = j % 3;
      if (num === 0) { var nameMenu = [["3年"], ["2年"], ["1年"]]; }
      var cell = targetSheet.getRange(shiftDataIn[i] + String(j));
      var data = cell.getValue();
      if (data) {
        var check = data.split(',');
        nameMenu[num] = makeNameMenu(nameMenu[num], check, num);
      }
      else {
        nameMenu[num] = [];
      }
      if (num === 2) {
        var pulldown = nameMenu.flat();
        var sc = shiftDataIn[i] + String(j + 16 + n) + ":" + shiftDataIn[i] + String(j + 19 + n);
        var setCell = targetSheet.getRange(sc);
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(pulldown).build();
        setCell.setDataValidation(rule);
        n = n + 1;
      }
    }
  }
}

// 学年用に用意されたリストとセルの値から、学年リストを作成
function makeNameMenu(grades, names, num) {
  var gradeWord = ["I", "L", "O"];
  var checkWord = ["K", "N", "Q"];
  for (var i = 0; i < names.length; i++) {
    // Logger.log(names[i].trim());
    for(var j = 5; j <= 19; j++) {
      var targetCell = targetSheet.getRange(gradeWord[num] + String(j));
      var checkCell = targetSheet.getRange(checkWord[num] + String(j));
      var fullName = targetCell.getValue();
      var lastName = fullName.split(' ')[0];
      if (names[i].trim() === fullName) { 
        if (checkCell.isChecked() === false) {grades.push(fullName);}
      }
      else if(names[i].trim() === lastName) { 
        if (checkCell.isChecked() === false) {grades.push(fullName);}
      }
    }
  }
  return grades;
}

// ここが変更時に実行されるmain関数
// フォームが提出された時のシート作成(要変更時トリガー)
function onCheck(e) {
  var sheet = e.source.getActiveSheet(); // 編集されたシートを取得
  var range = e.source.getActiveRange();
  var col = range.getColumn();
  var row = range.getRow();
  var cellValue = range.getValue();

  // 編集されたシートが調整表でなければ、実行終了
  if (sheet.getName() !== targetSheetName) { return; }

  // 確定版処理
  if (col === 7 && row === 41) { endEditShift(); }

  // 各アップデート
  else if (row >= 5 && row <= 19) {
    // 名前表示アップデート
    if      (col === 10) { updateSheet(0, row, col, cellValue); }
    else if (col === 13) { updateSheet(1, row, col, cellValue); }
    else if (col === 16) { updateSheet(2, row, col, cellValue); }
    // プルダウンアップデート
    else if (col === 11 || col === 14 || col === 17) { makePulldown(); }
  }
}

// 決定版が作成完了した後のリセット
function endEditShift() {
  // G41のセルのチェックボックスを判定
  var checkBox = targetSheet.getRange("G41");
  var check = box.getValue();

  if (check === true) {
    // チェックボックスがチェックされたときの処理
    var response = Browser.msgBox("警告", "決定表以外リセットします！本当によろしいですか？", Browser.Buttons.OK_CANCEL);
    if (response == "ok") {
      // 調査表リセット
      var range = targetSheet.getRange("C3:G17");
      range.clearContent();
      // 名簿表リセット
      range = targetSheet.getRange("I5:Q19");
      range.clearContent();
      range.clearDataValidations();
      var firstnum = "0人";
      range = targetSheet.getRange("J3");
      range.setValue(firstnum);
      range = targetSheet.getRange("M3");
      range.setValue(firstnum);
      range = targetSheet.getRange("P3");
      range.setValue(firstnum);
      // 備考リセット
      range = targetSheet.getRange("I24:N42");
      range.clearContent();
      // プルダウンの削除
      range = targetSheet.getRange("C21:G40");
      range.clearDataValidations();
    }
  }
  checkBox.setValue(false);
}

// 名前表示にチェックが付いたときの更新(要変更時トリガー)
function updateSheet(gnum, row, col, tf) {
  var fullName = targetSheet.getRange(row, col - 1).getValue();
  var lastName = fullName.split(' ')[0];
  var names = [lastName, fullName];
  var oldName = tf ? 0 : 1;
  var newName = tf ? 1 : 0;

  for (var i = 3; i <= 7; i++) {
    for (var j = 0; j <= 5; j++) {
      var num = 3 * (j + 1) + gnum;
      var nameListCell = targetSheet.getRange(num, i);
      var nameList = nameListCell.getValue().split(',');
      var updateText = "";
      for (var n = 0; n < nameList.length; n++) {
        if (nameList[n].trim() === names[oldName]) { nameList[n] = names[newName]; }
        if (n === 0) { updateText = nameList[n].trim(); }
        else { updateText = updateText + ', ' + nameList[n].trim();}
      }
      nameListCell.setValue(updateText);
    }
  }   
}
