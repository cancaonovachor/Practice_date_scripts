// 日付用ライブラリとしてMoment.jsを使用
// https://tonari-it.com/gas-moment-js-moment/

var FOLDER_ID_TARGET = PropertiesService.getScriptProperties("FOLDER_ID_TARGET"); // 練習時間調整シート作成先フォルダのID
var TEMPLATE_SS_ID = PropertiesService.getScriptProperties("TEMPLATE_SS_ID"); // 書式のテンプレートにするスプレッドシート
var TEMPLATE_DB_ID = PropertiesService.getScriptProperties("TEMPLATE_DB_ID"); // practice_date_template_db(member_dbからqueryで取ってきた名簿)
var TEMPLATE_DB_NAME = 'practice_date_template_db';
var CALCULATE_SHEET_NAME = 'calculate';
var BRANCH_NUM = 3; //支部の数

function practice_date_sheet_create() {
  var target_moment = Moment.moment();
  
  // 作成するスプレッドシート用に日付をあわせる
  var month = target_moment.add(3, "months").month(); // 3ヶ月後の月(1月 = 0)
  var month_string = ("00" + (month + 1)).slice(-2); // 3ヶ月後の月の文字列(1月 = 1)
  var year = target_moment.get("year");
  
  Logger.log(month);
  
  var template = DriveApp.getFileById(TEMPLATE_SS_ID);
  var folder = DriveApp.getFolderById(FOLDER_ID_TARGET);
  
  // テンプレートを「練習時間調整_YYYYMM」としてコピー
  var target_spreadsheet = template.makeCopy('練習時間調整_' + year + month_string, folder).getId();
  
  
  // 練習候補日生成
  // FIXME: 各日付に対して祝日かどうかを確かめるために、毎回祝日と本番日程のカレンダーを開いているため非常に重い。
  // 先に祝日と本番日程をリストに入れておいて、そこから検索するように要変更
  var candidates = [];
  
  var start = Moment.moment({ year: year, month: Number(month), day: 1});
  var end = Moment.moment({ year: year, month: Number(month) + 1, day: 1});
  
  Logger.log(start.month());
  Logger.log(end.month());
  
  for (var d = start; d < end; d.add(1, 'days')) {
    if (isHoliday(d)) {
      candidates.push(d.month() + 1 + '/' + d.date());
    }
  }
  
  Logger.log(candidates);
  
  var target_sheet_name = copyFromTemplate(target_spreadsheet, candidates, month_string);
  copyMembersFromDB(target_spreadsheet, target_sheet_name);
  createCalculate(target_spreadsheet, candidates);
  moveCalculate(target_spreadsheet, month_string, candidates.length);
  shareAndNotify(target_spreadsheet);
}

// 団員id、名前、所属、パートをコピー
function copyMembersFromDB(practice_date_sheet_id, sheet_name) {
  var target_sheet = SpreadsheetApp.openById(practice_date_sheet_id).getSheetByName(sheet_name);
  var db = SpreadsheetApp.openById(TEMPLATE_DB_ID).getSheetByName(TEMPLATE_DB_NAME);
  
  var num_db = db.getLastRow()-1; // stateがmemberの団員の数
  var datas = db.getRange(2, 1, num_db, 4).getValues();
  
  // コピー実行
  var target_cells = target_sheet.getRange(4, 1, num_db, 4)
  target_cells.setValues(datas);
  
  // 書式をコピー
  
  var lastCol = target_sheet.getLastColumn();
  target_sheet.getRange(4, 1, 1, lastCol - 2).copyFormatToRange(target_sheet, 1, lastCol - 2, 5, 3 + num_db);
  
  target_sheet.getRange(4, lastCol - 1, 1, 2).copyTo(target_sheet.getRange(5, lastCol - 1, num_db - 1, lastCol));
}

// calculateシートを作成
function createCalculate(practice_date_sheet_id, candidate_list) {
  var calculate_sheet = SpreadsheetApp.openById(practice_date_sheet_id).getSheetByName(CALCULATE_SHEET_NAME); // 作成したスプレッドシートのcalculateシート
  var branches = ["E1", "S1", "AG1"];  // テンプレートのセル(支部)
  var times = ["E3:R8", "S3:AF8", "AG3:AT8"]; // テンプレートのセル(時刻と計算セル)
  
  var lastCol = calculate_sheet.getLastColumn();
  
  for (var i = 0; i < (candidate_list.length - 1) * BRANCH_NUM; i++) {
    calculate_sheet.insertColumnsAfter(lastCol + 14*i, 14);
    var branch = calculate_sheet.getRange(branches[i % 3]);
    var targetToCopy = calculate_sheet.getRange(1, lastCol + 1 + i*14);
    branch.copyTo(targetToCopy);
    branch.copyTo(targetToCopy.offset(1, 0), {formatOnly: true});
    
    var targetToCopy_time = calculate_sheet.getRange(3, lastCol + 1 + i*14);
    calculate_sheet.getRange(times[i % 3]).copyTo(targetToCopy_time);
    
    calculate_sheet.getRange(1, lastCol + 1+ 14*i, 1, 14).merge();
    calculate_sheet.getRange(2, lastCol + 1 + 14*i, 1, 14).merge();
  }
  
  var column_next = 5;
  candidate_list.forEach (function (date) {
    for (var i = 0; i < 3; i++) {
      var target_cell = calculate_sheet.getRange(2, column_next);
      target_cell.setValue(date);
      column_next += 14;
    }
  });
}

// calculateシートを練習時間調整シートに移動する
function moveCalculate(spreadsheet_id, month, candidate_num) {
  var targetSpreadSheet = SpreadsheetApp.openById(spreadsheet_id);
  var practice_date_sheet = targetSpreadSheet.getSheetByName(month+"月");
  var calculate_sheet = targetSpreadSheet.getSheetByName("calculate");
  
  var calculate_range = calculate_sheet.getRange(5, 1, 5, 14*3+4);
  
  var target_range = practice_date_sheet.getRange(practice_date_sheet.getLastRow(), 1);
  
  calculate_range.copyTo(target_range);
  
  var copyrange = practice_date_sheet.getRange(practice_date_sheet.getLastRow() - 4, 14 * BRANCH_NUM);
  var column_next = 14 + 4 + 1;
  for (var i = 0; i < candidate_num; i++) {
    //copyrange.copyTo(destination);
  }
  
}

function copyFromTemplate(practice_date_sheet_id, candidate_dates, month) {
  var practice_date_sheet = SpreadsheetApp.openById(practice_date_sheet_id).getSheetByName('template');
  practice_date_sheet.setName(month+'月')
  
  // 9-22時:14列
    
  var lastCol = practice_date_sheet.getLastColumn() - 2;  // 最後の2つ前の列(完了ステータスの前の列)の番号
  var branches = ["E1", "S1", "AG1"]; // テンプレートのセル(支部)
  var times = ["E3:R4", "S3:AF4", "AG3:AT4"];  // テンプレートのセル(時刻)
  
  for (var i = 0; i < (candidate_dates.length - 1)*3; i++) {
    // 支部の部分をコピーする
    practice_date_sheet.insertColumnsAfter(lastCol + 14*i, 14); // 現在のシートの最後のセルの後に14列追加
    var branch = practice_date_sheet.getRange(branches[i % 3]); // 作りたい支部のテンプレートを選択
    var targetToCopy = practice_date_sheet.getRange(1, lastCol + 1 + i*14); // コピー先のセルを選択
    branch.copyTo(targetToCopy);  // コピーを実行
    branch.copyTo(targetToCopy.offset(1, 0), {formatOnly:true}); // 支部部分の書式を日付部分にもコピー
    
    // 時刻の部分をコピーする
    var targetToCopy_time = practice_date_sheet.getRange(3, lastCol + 1 + i*14); //コピー先のセル
    practice_date_sheet.getRange(times[i % 3]).copyTo(targetToCopy_time);
    
    practice_date_sheet.getRange(1, lastCol + 1 + 14 * i, 1, 14).merge(); // 支部のセルを結合
    practice_date_sheet.getRange(2, lastCol + 1 + 14 * i, 1, 14).merge(); // 日付のセルを結合
  }
  
  // 日付を入力

  var column_next = 5;
  candidate_dates.forEach (function (date) {
    for (var i = 0; i < 3; i++) {
      var target_cell = practice_date_sheet.getRange(2, column_next);
      target_cell.setValue(date);
      column_next += 14;
    }
  });
  
  return practice_date_sheet.getSheetName();
}

function shareAndNotify(target_spreadsheet_id) {
  var db = SpreadsheetApp.openById(TEMPLATE_DB_ID).getSheetByName('practice_date_template_db');
  var lastRow = db.getLastRow();
  var lastCol = db.getLastColumn();
  var emails = db.getRange(2, 5, lastRow - 1, 1).getValues();  // 共有相手のメンバーのメールアドレスを取得
  var targetSpreadsheet = DriveApp.getFileById(target_spreadsheet_id);
  
  Logger.log(emails);
  
  // 団員のメールアドレスに共有
  for (var i = 0; i < emails.length; i++) {
    //targetSpreadsheet.addEditor(emails[i][0]);
    Logger.log(emails[i][0]);
  }
  
}

