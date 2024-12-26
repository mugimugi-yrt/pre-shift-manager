// シートをグローバル化
var targetSpreadsheetId = "Sheet ID"; // 使用スプレッドシートのID
var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
var shiftSheetName = "調整表";                                            // 調整表作成シート名を指定
var shiftSheet = targetSpreadsheet.getSheetByName(shiftSheetName);
var nameSheetName = "名簿表";
var nameSheet = targetSpreadsheet.getSheetByName(nameSheetName);
var noteSheetName = "備考";
var noteSheet = targetSpreadsheet.getSheetByName(noteSheetName);

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
    formResponses[6].split(','),  // 2限
    formResponses[7].split(','),  // L限
    formResponses[8].split(','),  // 3限
    formResponses[9].split(','),  // 4限
    formResponses[10].split(',')  // 5限
  ]

  // シフト希望格納
  shiftData = updateShiftData(shiftData, responseData);
  var count = shiftNum(shiftData);

  // 名前データ取得(苗字&フルネーム)
  var lastName = formResponses[4];
  var fullName = lastName + ' ' + formResponses[5];
  
  // 学年・学科・コースデータを取得
  var studentYear = formResponses[1];
  var course = formResponses[2] + ' / ' + formResponses[3];
  if (studentYear === "1") {
    course = course + "希望";
  }

  // 名簿表更新
  updateNameList(studentYear, course, fullName, count);

  // 備考表記入
  var note = formResponses[11];
  if (note !== "") { updateNote(note, fullName); }

  // シフト表作成
  makeResearchSheet(shiftData, studentYear, lastName);
  makePulldown();
}

// 個人シフト表リストとシフト希望回答から、個人シフト表リストを更新
function updateShiftData(data, respose) {
  for (var i = 0; i < 5; i++) {
    if (respose[i].length !== 0) {
      for (var j = 0; j < respose[i].length; j++) {
        var request = respose[i][j].trim();
        if      (request === "月曜日") { data[i][0] = true; }
        else if (request === "火曜日") { data[i][1] = true; }
        else if (request === "水曜日") { data[i][2] = true; }
        else if (request === "木曜日") { data[i][3] = true; }
        else if (request === "金曜日") { data[i][4] = true; }
      }
    }
  }
  return data;
}

// シフトに入っているコマ数を返す
function shiftNum(data) {
  var count = 0;
  for (var i = 0; i < 5; i++) {
    for (var j = 0; j < 5; j++) {
      if (data[i][j] === true) {
        count = count + 1;
      }
    }
  }
  return count;
}

// 名簿リストを更新
function updateNameList(year, course, name, cnum) {
  var targetWord = [
    ["A", "B", "C", "D"],  // 4年
    ["E", "F", "G", "H"],  // 3年
    ["I", "J", "K", "L"],  // 2年
    ["M", "N", "O", "P"]   // 1年
  ]
  if      (year === "4") { var num = 0; }
  else if (year === "3") { var num = 1; }
  else if (year === "2") { var num = 2; }
  else if (year === "1") { var num = 3; }

  // 提出人数更新
  var grade = nameSheet.getRange(targetWord[num][1] + "3");
  var cellValue = grade.getValue().split(' ');
  var submitNum = parseInt(cellValue[0], 10) + 1;
  grade.setValue(String(submitNum) + "人");

  // 空欄部分を調査
  var n = 5;
  var targetCell = nameSheet.getRange(targetWord[num][0] + String(n));
  var usedChecker = targetCell.getValue();
  while(usedChecker) {
    n = n + 1;
    targetCell = nameSheet.getRange(targetWord[num][0] + String(n));
    usedChecker = targetCell.getValue();
  }

  // 空欄部分に新たな提出者情報を挿入
  for (var i = 0; i < 4; i++) {
    targetCell = nameSheet.getRange(targetWord[num][i] + String(n));
    if (i === 0) {
      targetCell.setValue(name + ' (' + cnum + ')');
    }
    else if (i === 1) {
      targetCell.setValue(course);
    }
    else {
      targetCell.insertCheckboxes();
    }
  }
}

// 備考リストを更新
function updateNote(sentence, name) {
  // 空欄部分を調査
  var n = 3;
  var targetCell = noteSheet.getRange("A" + String(n));
  var usedChecker = targetCell.getValue();
  while(usedChecker) {
    n = n + 1;
    targetCell = noteSheet.getRange("A" + String(n));
    usedChecker = targetCell.getValue();
  }

  // 空欄部分に新たな備考情報を挿入
  targetCell.setValue(name);
  targetCell = noteSheet.getRange("B" + String(n));
  targetCell.setValue(sentence);
}

// 全体シフト表のシート作成
function makeResearchSheet(data, year, moji) {
  var column = ["C", "D", "E", "F", "G"];

  if      (year === "4") { var num = 3; }
  else if (year === "3") { var num = 4; }
  else if (year === "2") { var num = 5; }
  else if (year === "1") { var num = 6; }

  // ヨコでデータ格納
  for (var i = 0; i < 5; i++) {
    var x = num + (i * 3);
    for (var j = 0; j < 5; j++) {
      if (data[i][j]) {
        var cell = column[j] + String(x);
        var targetCell = shiftSheet.getRange(cell);
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
    for (var j = 3; j <= 22; j++) {
      var num = j % 4;
      if (num === 3) { var nameMenu = [["4年"], ["3年"], ["2年"], ["1年"]]; }
      var cell = shiftSheet.getRange(shiftDataIn[i] + String(j));
      var data = cell.getValue();
      if (data) {
        var check = data.split(',');
        nameMenu[(num + 1) % 4] = makeNameMenu(nameMenu[(num + 1) % 4], check, (num + 1) % 4);
      }
      else {
        nameMenu[(num + 1) % 4] = [];
      }
      if (num === 2) {
        var pulldown = nameMenu.flat();
        var sc = shiftDataIn[i] + String(j + 20 + (n * 4)) + ":" + shiftDataIn[i] + String(j + 24 + (n * 4));
        var setCell = shiftSheet.getRange(sc);
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(pulldown).build();
        setCell.setDataValidation(rule);
        n = n + 1;
      }
    }
  }
}

// 学年用に用意されたリストとセルの値から、学年リストを作成
function makeNameMenu(grades, names, num) {
  var gradeWord = ["A", "E", "I"];
  var checkWord = ["D", "H", "L"];
  for (var i = 0; i < names.length; i++) {
    for(var j = 5; j <= 19; j++) {
      var targetCell = nameSheet.getRange(gradeWord[num] + String(j));
      var checkCell = nameSheet.getRange(checkWord[num] + String(j));
      var name = targetCell.getValue();
      var fullName = name.split(' ')[0] + ' ' + name.split(' ')[1]
      var lastName = name.split(' ')[0];
      if (names[i].trim() === fullName) { 
        if (checkCell.isChecked() === false) { grades.push(fullName); }
      }
      else if(names[i].trim() === lastName) { 
        if (checkCell.isChecked() === false) { grades.push(fullName); }
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

  // 編集されたシートが調整表もしくは名簿表でなければ、実行終了
  if (sheet.getName() !== shiftSheetName && sheet.getName() !== nameSheetName) { return; }

  // 確定版処理
  if (sheet.getName() === shiftSheetName) {
    if (col === 7) {
      if (row === 67 || row === 68) { endEditShift(row); }
    }
  }

  // 各アップデート
  else if (sheet.getName() === nameSheetName) {
    if (row >= 5 && row <= 19) {
      // 名前表示アップデート
      if      (col === 3)  { updateSheet(0, row, col, cellValue); }  // 3年
      else if (col === 7)  { updateSheet(1, row, col, cellValue); }  // 2年
      else if (col === 11) { updateSheet(2, row, col, cellValue); }  // 1年
      // プルダウンアップデート
      else if (col === 4 || col === 8 || col === 12 || col === 16) { makePulldown(); }
    }
  }
}

// 決定版が作成完了した後の処理
function endEditShift(num) {
  // G41のセルのチェックボックスを判定
  var checkBox = shiftSheet.getRange("G" + String(num));
  var check = checkBox.getValue();

  if (check === true) {
    // チェックボックスがチェックされたときの処理
    var response = Browser.msgBox("完成！", "固定シフト表の作成に入ります。よろしいですか？", Browser.Buttons.OK_CANCEL);
    if (response == "ok") {
      // 決定版シフト表の作成
      var userInput = SpreadsheetApp.getUi().prompt('作成年度の入力:', SpreadsheetApp.getUi().ButtonSet.OK);
      var yearText = userInput.getResponseText();
      var newSpreadsheet = makeShiftSheet(yearText);

      // 調査資料のコピー作成
      copyData(newSpreadsheet);

      // シートリセット(SA体験ありの場合のみ実行)
      if (num === 68) { resetSheet(); }
      SpreadsheetApp.getUi().alert('シフト表作成完了しました');
    }
  }
  checkBox.setValue(false);
  if(response == "ok" && num === 68) { SpreadsheetApp.getUi().alert('リセットが完了しました'); }
}

// 配布するシフト表のスプシデータを作成
function makeShiftSheet(year) {
  var newSpreadsheet = SpreadsheetApp.create(year + "年度春学期_固定シフト");
  var spreadsheetId = newSpreadsheet.getId();
  var folderId = "1GzDJCqlK5Fef6ZpXhzNL60xnGtKvmWDp";
  DriveApp.getFileById(spreadsheetId).moveTo(DriveApp.getFolderById(folderId));

  // シート1を固定シフトシートとして作成
  var sheet = newSpreadsheet.getSheetByName('シート1');
  if (sheet) { sheet.setName('固定シフト'); }

  // 固定シフトシートの形を作る
  makeSheetShape(sheet, year);

  // 決定表からソートデータの取得(月, 火, 水, 木, 金で取得)
  var decidedShiftData = [];
  for (var i = 3; i <= 7; i++) {
    var weekData = [];
    for (var j = 0; j <= 4; j++) {
      var data = [];
      for(var k = 26; k <= 33; k++) {
        var charge = shiftSheet.getRange(k + (j * 8), i).getValue();
        data.push(charge);
      }
      data = addGradeData(data);
      weekData.push(data);
    }
    decidedShiftData.push(weekData);
  }

  // 相談員さんデータの取得
  var counselorData = getCounselorData();

  // 決定表作成
  drawShiftSheet(decidedShiftData, counselorData, sheet);

  return newSpreadsheet;
}

// 固定シフト表のシートの形を作る
function makeSheetShape(sheet, year) {
  // 列の幅を設定
  for (var col = 1; col <= 11; col++) {
    sheet.setColumnWidth(col, 100);
  }

  // 行の高さを設定
  sheet.setRowHeight(1, 40);
  for (var row = 2; row <= 30; row++) {
    sheet.setRowHeight(row, 22);
  }
  sheet.setRowHeight(31, 21);

  // タイトルを付ける(20XX年春学期 固定シフト)
  var title = sheet.getRange("E1:G1");
  title.merge();
  title.setHorizontalAlignment('center');
  title.setVerticalAlignment('middle');
  title.setFontSize(15);
  title.setValue(year + "春学期 固定シフト");

  // 枠線設定
  var style = SpreadsheetApp.BorderStyle.SOLID;

  // 時限セット
  var periods = [
    sheet.getRange("A3:A8"),   // 2限
    sheet.getRange("A9:A12"),  // L限
    sheet.getRange("A13:A18"), // 3限
    sheet.getRange("A19:A24"), // 4限
    sheet.getRange("A25:A30")  // 5限
  ];
  var timeLabel = [
    "2限\n(10:50\n    ~ 13:30)",
    "L限\n(13:40\n    ~ 13:10)",
    "3限\n(13:20\n    ~ 15:00)",
    "4限\n(15:10\n    ~ 16:50)",
    "5限\n(19:00\n    ~ 19:40)"
  ];
  for (var i = 0; i < periods.length; i++) {
    periods[i].merge();
    periods[i].setBorder(true, true, true, true, false, false, "black", style);
    periods[i].setVerticalAlignment('top');
    periods[i].setFontSize(12);
    periods[i].setValue(timeLabel[i]);
  }
  
  // 曜日セット
  var blank = sheet.getRange("A2");
  blank.setBorder(true, true, true, true, false, false, "black", style);
  var weeks = [
    sheet.getRange("B2:C2"),
    sheet.getRange("D2:E2"),
    sheet.getRange("F2:G2"),
    sheet.getRange("H2:I2"),
    sheet.getRange("J2:K2")
  ];
  var weekLabel = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日"];
  for (var j = 0; j < weeks.length; j++) {
    weeks[j].merge();
    weeks[j].setBorder(true, true, true, true, false, false, "black", style);
    weeks[j].setHorizontalAlignment('center');
    weeks[j].setVerticalAlignment('top');
    weeks[j].setFontSize(11);
    weeks[j].setValue(weekLabel[j]);
  }

  // コマセット
  var boxes = [
    sheet.getRange("B3:C8"),   sheet.getRange("D3:E8"),   sheet.getRange("F3:G8"),   sheet.getRange("H3:I8"),   sheet.getRange("J3:K8"),
    sheet.getRange("B9:C12"),  sheet.getRange("D9:E12"),  sheet.getRange("F9:G12"),  sheet.getRange("H9:I12"),  sheet.getRange("J9:K12"),
    sheet.getRange("B13:C18"), sheet.getRange("D13:E18"), sheet.getRange("F13:G18"), sheet.getRange("H13:I18"), sheet.getRange("J13:K18"),
    sheet.getRange("B19:C24"), sheet.getRange("D19:E24"), sheet.getRange("F19:G24"), sheet.getRange("H19:I24"), sheet.getRange("J19:K24"),
    sheet.getRange("B25:C30"), sheet.getRange("D25:E30"), sheet.getRange("F25:G30"), sheet.getRange("H25:I30"), sheet.getRange("J25:K30")
  ]
  for (var n = 0; n < boxes.length; n++) {
    boxes[n].setFontSize(11);
    boxes[n].setBorder(true, true, true, true, false, false, "black", style);
  }

  // 相談員さんセット
  var counscell = [sheet.getRange("A31"), sheet.getRange("A32")];
  var counslabel = ["担当相談員", "滞在時間"];
  for (var x = 0; x < counscell.length; x++) {
    counscell[x].setBorder(true, true, true, true, false, false, "black", style);
    counscell[x].setHorizontalAlignment('center');
    counscell[x].setFontSize(10);
    counscell[x].setValue(counslabel[x]);
  }
  counscell[0].setBorder(true, true, false, true, false, false, "black", style);
  var nameCell = [sheet.getRange("B31:C31"), sheet.getRange("D31:E31"), sheet.getRange("F31:G31"), sheet.getRange("H31:I31"), sheet.getRange("J31:K31"),]
  var timeCell = [sheet.getRange("B32:C32"), sheet.getRange("D32:E32"), sheet.getRange("F32:G32"), sheet.getRange("H32:I32"), sheet.getRange("J32:K32"),]
  for (var y = 0; y < 5; y++) {
    nameCell[y].merge();
    timeCell[y].merge();
    nameCell[y].setBorder(true, true, false, true, false, false, "black", style);
    nameCell[y].setHorizontalAlignment('center');
    nameCell[y].setFontSize(11);
    timeCell[y].setBorder(false, true, true, true, false, false, "black", style);
    timeCell[y].setHorizontalAlignment('center');
    timeCell[y].setFontSize(11);
  }
}

// 決定版データリスト(コマ単位)から学年をつけたものを取得
function addGradeData(perdata) {
  // 名簿表から名前を取得
  var nameData = getFullnameData();

  // 学年順(降順)にソートし、出力する文字列を格納
  perdata = sortGrade(perdata, nameData);

  return perdata;
}

// 提出者名簿の名前リストを作成
function getFullnameData() {
  var gradeWord = ["A", "E", "I", "M"];
  var nameList = [[], [], [], []]; // 4, 3, 2, 1年の順で格納
  for (var i = 0; i < gradeWord.length; i++) {
    var n = 5;
    var fullName = nameSheet.getRange(gradeWord[i] + String(n)).getValue();
    while(fullName) {
      n = n + 1;
      fullName = fullName.split(' ')[0] + ' ' + fullName.split(' ')[1];
      nameList[i].push(fullName);
      fullName = nameSheet.getRange(gradeWord[i] + String(n)).getValue();
    }
  }
  return nameList;
}

// コマのデータと全体名簿リストから学年順にソート
function sortGrade(data, namelist) {
  var returnData = [[], [], [], []]  // このリストに4, 3, 2, 1年の情報を格納し、flatして最終的に返却
  var addWord = ["④" ,"③", "②", "①"];

  for (var i = 0; i <= 4; i++) {
    var name = data[i];
    for (var n = 0; n < namelist.length; n++) {
      for (var m = 0; m < namelist[n].length; m++) {
        if (name === namelist[n][m]) {
          if      (returnData[n].length === 0) { name = addWord[n] + name }
          else if (returnData[n].length > 0)   { name = "    " + name; }
          returnData[n].push(name);
        }
      }
    }
  }
  returnData = returnData.flat();
  if (returnData.length < 6) {
    for (i = 0; i < 6 - returnData.length; i++) {
      returnData.push('');
    }
  }
  for (var x = 5; x <= 7; x++) { returnData.push(data[x]); }
  return returnData;
}

// カウンセラーさんの曜日と時間を取得
function getCounselorData() {
  var returnData = []
  for (var i = 3; i <= 7; i++) {
    var name = shiftSheet.getRange(66, i).getValue();
    for (var j = 4; j <= 6; j++) {
      var check = nameSheet.getRange("R" + String(j)).getValue();
      if (name === check) { var time = nameSheet.getRange("S" + String(j)).getValue(); }
    }
    returnData.push([name, time]);
  }
  return returnData;
}

// ソート済みシフトデータからシフト表作成
function drawShiftSheet(data, cdata, sheet) {
  for (var i = 0; i < data.length; i++) {
    for (var j = 0; j < data[i].length; j++) {
      // 研修会のコマ
      if (i === 4 && j === 1) {
        var training = sheet.getRange("J9:K10");
        training.merge();
        training.setHorizontalAlignment('center');
        training.setVerticalAlignment('top');
        training.setFontSize(15);
        training.setValue("(研修会)");
      }

      else {
        // L限のコマ
        if (j === 1) {
          for (var k = 0; k <= 4; k++) {
            if (k <= 2) {
              var setCell = sheet.getRange(9 + k, 2 * (i + 1));
              setCell.setValue(data[i][j][k]);
            }
            else if (k === 3) {
              var setCell = sheet.getRange(12, 2 * (i + 1));
              setCell.setValue(data[i][j][k]);
            }
            else if (k === 4) {
              var setCell = sheet.getRange(12, 2 * (i + 1));
              if (setCell.getValue()) { setCell = sheet.getRange(12, 2 * (i + 1) + 1); }
              setCell.setValue(data[i][j][k]);
            }
          }
        }
        // 2限のコマ(L限の関係でズレるため)
        else if (j === 0) {
          for (var k = 0; k < data[i][j].length; k++) {
            if(data[i][j][k] !== '') {
              // SAのコマ
              if (k <= 2) {
                var setSA = sheet.getRange(3 + k, 2 * (i + 1));
                setSA.setValue(data[i][j][k]);
              }
              // SA体験のコマ
              else if (k === 3) {
                var setSApre = sheet.getRange(4, 2 * (i + 1));
                setSApre.setValue(data[i][j][k]);
              } 
              else if (k === 4) {
                var setSApre = sheet.getRange(4, 2 * (i + 1));
                if (setSApre.getValue()) { setSApre = sheet.getRange(4, 2 * (i + 1) + 1); }
                setSApre.setValue(data[i][j][k]);
              }
              // TAのコマ
              else if (k === 5) {
                var setTA = sheet.getRange(7, 2 * (i + 1));
                setTA.setBackground(tacolor(data[i][j][k]));
                if(tacolor(data[i][j][k]) !== "#ffffff") { setTA.setHorizontalAlignment('center'); }
                setTA.setValue('TA ' + data[i][j][k]);
              }
              // 先生のオフィスアワー
              else if (k === 6) {
                var setTeacher = sheet.getRange(8, 2 * (i + 1));
                setTeacher.setFontColor(teachercolor(data[i][j][k]));
                setTeacher.setValue(data[i][j][k]);
              }
              else if (k === 7) {
                var setTeacher = sheet.getRange(8, 2 * (i + 1));
                if (setTeacher.getValue()) { setTeacher = sheet.getRange(8, 2 * (i + 1) + 1); }
                setTeacher.setFontColor(teachercolor(data[i][j][k]));
                setTeacher.setValue(data[i][j][k]);
              }
            }
          }
        }
        else {
          for (var k = 0; k < data[i][j].length; k++) {
            if(data[i][j][k] !== '') {
              // SAのコマ
              if (k <= 2) {
                var setSA = sheet.getRange(13 + (j - 2) * 6 + k, 2 * (i + 1));
                setSA.setValue(data[i][j][k]);
              }
              // SA体験のコマ
              else if (k === 3) {
                var setSApre = sheet.getRange(16 + (j - 2) * 6, 2 * (i + 1));
                setSApre.setValue(data[i][j][k]);
              } 
              else if (k === 4) {
                var setSApre = sheet.getRange(16 + (j - 2) * 6, 2 * (i + 1));
                if (setSApre.getValue()) { setSApre = sheet.getRange(16 + (j - 2) * 6, 2 * (i + 1) + 1); }
                setSApre.setValue(data[i][j][k]);
              }
              // TAのコマ
              else if (k === 5) {
                var setTA = sheet.getRange(17 + (j - 2) * 6, 2 * (i + 1));
                setTA.setBackground(tacolor(data[i][j][k]));
                if(tacolor(data[i][j][k]) !== "#ffffff") { setTA.setHorizontalAlignment('center'); }
                setTA.setValue('TA ' + data[i][j][k]);
              }
              // 先生のオフィスアワー
              else if (k === 6) {
                var setTeacher = sheet.getRange(18 + (j - 2) * 6, 2 * (i + 1));
                setTeacher.setFontColor(teachercolor(data[i][j][k]));
                setTeacher.setValue(data[i][j][k]);
              }
              else if (k === 7) {
                var setTeacher = sheet.getRange(18 + (j - 2) * 6, 2 * (i + 1));
                if (setTeacher.getValue()) { setTeacher = sheet.getRange(18 + (j - 2) * 6, 2 * (i + 1) + 1); }
                setTeacher.setFontColor(teachercolor(data[i][j][k]));
                setTeacher.setValue(data[i][j][k]);
              }
            }
          }
        }
      }
    }
  }

  // カウンセラーデータ記入
  for (var n = 0; n < cdata.length; n++) {
    var setName = sheet.getRange(31, 2 * (n + 1));
    setName.setValue(cdata[n][0]);
    var setTime = sheet.getRange(32, 2 * (n + 1));
    setTime.setValue(cdata[n][1]);
  }
}

// TA識別用関数
function tacolor(data) {
  if      (data === "Java対応")   { return "#f0f8ff" }  // Java：ライトブルー
  else if (data === "python対応") { return "#fabca5" }  // python：淡いオレンジ
  else if (data === "MATLAB対応") { return "#c8fcc0" }  // MATLAB：淡い緑
  else if (data === "C++対応")    { return "#f9fcc0" }  // C++：淡い黄色
  else                            { return "#ffffff" }  // GBC TA：白
}

// 先生識別用関数
function teachercolor(data) {
  var ecteacher = ["李先生", "黄先生", "デルグレゴ先生", "花崎先生", "馬先生"];
  for (var i = 0; i < ecteacher.length; i++) {
    if(data == ecteacher[i]) { return "#fc778e" }  // EC対応の先生：ピンク
  }
  return "#0000ff" // その他の先生：青
}

// シフト決定のために使用したデータをコピー
function copyData(sheet) {
  shiftSheet.copyTo(sheet).setName("データ(調整表)");
  nameSheet.copyTo(sheet).setName("データ(名簿表)");
  noteSheet.copyTo(sheet).setName("データ(備考)");
}

// 全てのデータリセット
function resetSheet() {
  var clearCell = [
    shiftSheet.getRange("C3:G22"),   // 調査表
    shiftSheet.getRange("C26:G30"),  // 決定表2限
    shiftSheet.getRange("C34:G38"),  // 決定表L限
    shiftSheet.getRange("C42:G46"),  // 決定表3限
    shiftSheet.getRange("C50:G54"),  // 決定表4限
    shiftSheet.getRange("C58:G62"),  // 決定表5限
    nameSheet.getRange("A5:P19:"),    // 名簿表
    noteSheet.getRange("A3:B21")     // 備考
  ]
  for (var i = 0; i < 8; i++) {
    clearCell[i].clearContent();
    clearCell[i].clearDataValidations();
  }


  var setCell = [
    nameSheet.getRange("B3"),
    nameSheet.getRange("F3"),
    nameSheet.getRange("J3")
  ]
  for (var j = 0; j < 3; j++) {
    setCell[j].setValue("0人");
  }
}

// 名前表示にチェックが付いたときの更新
function updateSheet(gnum, row, col, tf) {
  var name = nameSheet.getRange(row, col - 2).getValue();
  var fullName = name.split(' ')[0] + ' ' + name.split(' ')[1];
  var lastName = name.split(' ')[0];
  var names = [lastName, fullName];
  var oldName = tf ? 0 : 1;
  var newName = tf ? 1 : 0;

  for (var i = 3; i <= 7; i++) {
    for (var j = 0; j <= 5; j++) {
      var num = 3 * (j + 1) + gnum;
      var nameListCell = shiftSheet.getRange(num, i);
      var nameList = nameListCell.getValue().split(',');
      if (nameList.length > 0) {
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
}
