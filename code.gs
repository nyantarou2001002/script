function doGet(e) {
  var mode = e.parameter.mode;
  var page = e.parameter.page;
  var selectedEmpId = e.parameter.empId;
  var password = e.parameter.password;  // パスワードのパラメータを取得

  // スクリプトプロパティに保存したパスワードと照合
  var storedPassword = PropertiesService.getScriptProperties().getProperty("adminPassword");

  // パスワードが一致しない場合はエラーメッセージを表示
  if (password && password !== storedPassword) {
    return HtmlService.createHtmlOutput("パスワードが一致しません。");
  }

  // 月次レポート表示用
  if (mode === "report") {
    return HtmlService.createTemplateFromFile("view_report")
                      .evaluate()
                      .setTitle("月次勤怠レポート");
  }

  // 従業員詳細表示
  if (selectedEmpId) {
    PropertiesService.getUserProperties().setProperty('selectedEmpId', selectedEmpId.toString());
    var selectedMonth = e.parameter.month || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM");
    var template = HtmlService.createTemplateFromFile("view_detail");
    template.timeClocks = getTimeClocksForMonth(selectedMonth);
    template.selectedMonth = selectedMonth;
    return template.evaluate().setTitle("Detail: " + selectedEmpId.toString());
  }

  // ログインページ
  if (page === "login") {
    return HtmlService.createTemplateFromFile("login")
                      .evaluate()
                      .setTitle("ログイン");
  }

  // 従業員一覧ページ
  if (page === "view") {
    return HtmlService.createTemplateFromFile("view_home")
                      .evaluate()
                      .setTitle("Home");
  }

  return HtmlService.createTemplateFromFile("kintai_home")
                    .evaluate()
                    .setTitle("Home");
}







/**
 * このアプリのURLを返す
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * 従業員一覧を取得する関数（出勤回数情報も付与）
 */
function getEmployees() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var empSheet = ss.getSheetByName("従業員名簿");
  if (!empSheet) return [];
  
  var lastRow = empSheet.getLastRow();
  if (lastRow < 2) return []; // データがない場合は空配列
  
  // 6列分：従業員ID, 名前, 職種, 出勤回数上限, 有給休暇上限, 有給休暇付与日
  var empRange = empSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  var employees = [];
  var now = new Date();
  
  // 現在月の開始・終了日時
  var currentMonthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  var currentMonthEnd = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);
  
  for (var i = 0; i < empRange.length; i++) {
    var empId = empRange[i][0];
    var empName = empRange[i][1];
    var empJob = empRange[i][2];
    var maxAttendanceCount = empRange[i][3];
    var paidVacationLimit = empRange[i][4];
    var paidVacationGrantDate = empRange[i][5];
    if (!empId || empId.toString().trim() === "") continue;
    
    // 現在月の出勤回数（「出勤」レコードのみ）
    var attendanceCountCurrentMonth = getAttendanceCount(empId, currentMonthStart, currentMonthEnd);
    
    // ※変更点：有給休暇付与日から9ヶ月以降の出勤回数をチェック
    var attendanceCountAfterNineMonths = 0;
    if (paidVacationGrantDate) {
      var grantDate = new Date(paidVacationGrantDate);
      // 9ヶ月後の日付
      var nineMonthsLater = new Date(grantDate.getFullYear(), grantDate.getMonth() + 9, grantDate.getDate(), 23, 59, 59);
      // 9ヶ月以降、つまりnineMonthsLaterから「今」までの出勤回数を算出
      attendanceCountAfterNineMonths = getAttendanceCount(empId, nineMonthsLater, now);
    }
    
    employees.push({
      id: empId,
      name: empName,
      job: empJob,
      maxAttendanceCount: maxAttendanceCount,
      paidVacationLimit: paidVacationLimit,
      paidVacationGrantDate: paidVacationGrantDate,
      attendanceCountCurrentMonth: attendanceCountCurrentMonth,
      attendanceCountAfterNineMonths: attendanceCountAfterNineMonths
    });
  }
  return employees;
}


/**
 * 指定された従業員IDについて、startDate～endDateの間の「出勤」レコードの数を返す
 */
function getAttendanceCount(empId, startDate, endDate) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  if (!sheet) return 0;
  var last_row = sheet.getLastRow();
  if (last_row < 2) return 0;
  var data = sheet.getRange(2, 1, last_row - 1, 3).getValues();
  var count = 0;
  for (var i = 0; i < data.length; i++) {
    if (data[i][0].toString() === empId && data[i][1] === "出勤") {
      var dt = new Date(data[i][2]);
      if (dt >= startDate && dt <= endDate) {
        count++;
      }
    }
  }
  return count;
}




/**
 * 従業員情報の取得
 * ※ デバッグするときにはselectedEmpIdを存在するIDで書き換えてください
 */
function getEmployeeName() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿");
  if (!empSheet) return "不明";

  var lastRow = empSheet.getLastRow();
  if (lastRow < 2) return "不明"; // ヘッダー以外のデータがない場合

  var empData = empSheet.getRange(2, 1, lastRow - 1, 2).getValues(); // IDと名前のデータ取得

  for (var i = 0; i < empData.length; i++) {
    var empId = empData[i][0];  // A列（従業員ID）
    var empName = empData[i][1]; // B列（名前）

    Logger.log("Row " + (i + 2) + ": ID = " + empId + ", Name = " + empName);

    if (empId && empId.toString() === selectedEmpId.toString()) {
      return empName; // 名前を返す
    }
  }
  return "不明"; // 該当なし
}


/**
 * 📌 勤怠情報の取得
 * - `date` のフォーマットを「yyyy-MM-dd」に統一する（メモと一致させる）
 */
function getTimeClocks() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!selectedEmpId) return [];

  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  if (!timeClocksSheet) return [];

  var last_row = timeClocksSheet.getLastRow();
  if (last_row < 2) return [];

  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row - 1, 3);
  var data = timeClocksRange.getValues();
  var empTimeClocks = [];
  var now = new Date();
  var firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  for (var i = 0; i < data.length; i++) {
    var empId = data[i][0].toString();  // 🔥 ここで型を統一
    var type = data[i][1];
    var datetime = new Date(data[i][2]);

    if (!empId || empId.trim() === "" || empId !== selectedEmpId) continue;
    if (isNaN(datetime.getTime()) || datetime < firstDayOfMonth || datetime > now) continue;
    var formattedDate = Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd");

    var formattedDateTime = Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd HH:mm");
    empTimeClocks.push({
      'date': formattedDate,         // キー作成用は日付のみ (yyyy-MM-dd)
      'datetime': formattedDateTime, // 表示用の日時（yyyy-MM-dd HH:mm）
      'type': type,
      'rawDateTime': datetime
    });

  }

  // 日時順にソート
  empTimeClocks.sort(function(a, b) {
    return a.rawDateTime - b.rawDateTime;
  });

  return empTimeClocks;
}


/**
 * 勤怠情報登録
 */
function saveWorkRecord(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!selectedEmpId) {
    return 'エラー: 従業員が選択されていません。';
  }
  var targetDate = form.target_date;
  var targetTime = form.target_time;
  if (!targetDate || !targetTime) {
    return 'エラー: 日付または時刻が未入力です。';
  }
  var targetType = '';
  switch (form.target_type) {
    case 'clock_in': 
      targetType = '出勤'; 
      break;
    case 'clock_out': 
      targetType = '退勤'; 
      break;
    case 'paid_vacation': 
      targetType = '有給休暇';
      break;
    default: 
      return 'エラー: 無効な登録種別です。';
  }
  
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  if (!timeClocksSheet) {
    return '⚠ エラー: 打刻履歴シートが見つかりません。管理者にお問い合わせください。';
  }
  
  var targetRow = timeClocksSheet.getLastRow() + 1;
  var timestamp = new Date(`${targetDate}T${targetTime}:00`);
  if (isNaN(timestamp.getTime())) {
    return 'エラー: 無効な日付または時刻です。';
  }
  
  timeClocksSheet.getRange(targetRow, 1).setValue(selectedEmpId);
  timeClocksSheet.getRange(targetRow, 2).setValue(targetType);
  timeClocksSheet.getRange(targetRow, 3).setValue(timestamp);
  
  return '登録しました';
}


/**
 * 選択している従業員のメモカラムの値をspread sheetから取得する
 */
function getEmpMemo() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var checkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]// 「チェック結果」のシート
  var last_row = checkSheet.getLastRow()
  var timeClocksRange = checkSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var checkResult = "";
  var i = 1;
  while (true) {
    var empId =timeClocksRange.getCell(i, 1).getValue();
    var result =timeClocksRange.getCell(i, 2).getValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    if (empId == selectedEmpId){
        checkResult = result
        break;
    }
    i++
  }
  return checkResult
}

/**
 * メモの内容をSpreadSheetに保存する
 */
function saveMemo(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  // inputタグのnameで取得
  var memo = form.memo

  var targetRowNumber = getTargetEmpRowNumber(selectedEmpId)
  var sheet = SpreadsheetApp.getActiveSheet()
  if (targetRowNumber == null) {
    // targetRowNumberがない場合には新規に行を追加する
    // 現在の最終行に+1した行番号
    targetRowNumber = sheet.getLastRow() + 1
    // 1列目にempIdをセットして保存
    sheet.getRange(targetRowNumber, 1).setValue(selectedEmpId)
  }
  // memoの内容を保存
  sheet.getRange(targetRowNumber, 2).setValue(memo);
  return "登録しました";

}

/**
 * spreadSheetに保存されている指定のemployee_idの行番号を返す
 */
function getTargetEmpRowNumber(empId) {
  // 開いているシートを取得
  var sheet = SpreadsheetApp.getActiveSheet()
  // 最終行取得
  var last_row = sheet.getLastRow()
  // 2行目から最終行までの1列目(emp_id)の範囲を取得
  var data_range = sheet.getRange(1, 1, last_row, 1);
  // 該当範囲のデータを取得
  var sheetRows = data_range.getValues();
  // ループ内で検索
  for (var i = 0; i <= sheetRows.length - 1; i++) {
    var row = sheetRows[i]
    if (row[0] == empId) {
      // spread sheetの行番号は1から始まるが配列のindexは0から始まるため + 1して行番号を返す
      return i + 1;
    }
  }
  // 見つからない場合にはnullを返す
  return null
}

/**
 * 【A】打刻の不整合をチェックして結果を返す関数
 *   - 出勤があるのに退勤が無い
 *   - 休憩開始があるのに休憩終了が無い
 * 等を簡易的にチェック
 */
function checkInconsistenciesForEmp(empId) {
  // 「本日」の開始時刻（00:00）を算出
  var now = new Date();
  var startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  
  // 今月の開始日時
  var firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  
  // タイプ別に集計用のオブジェクト
  var dailyRecords = {}; 
  // dailyRecords[日付文字列] = { 
  //    punchIn: [Date1, Date2, ...],   // 出勤打刻のリスト
  //    punchOut: [...],               // 退勤打刻リスト
  //    breakBegin: [...],             // 休憩開始
  //    breakEnd: [...],               // 休憩終了
  // }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  var last_row = sheet.getLastRow();
  var range = sheet.getRange(2, 1, last_row, 3);
  var data = range.getValues();
  
  for (var i = 0; i < data.length; i++) {
    var rowEmpId = data[i][0];
    var rowType  = data[i][1];
    var rowDateTime = data[i][2];
    if (!rowEmpId || rowEmpId === "") break;
    
    // 対象従業員、かつ、当月かつ、本日以前のデータのみを対象
    if (rowEmpId == empId && rowDateTime >= firstDayOfMonth && rowDateTime < startOfToday) {
      var dateObj = new Date(rowDateTime);
      var key = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy-MM-dd");
      
      if (!dailyRecords[key]) {
        dailyRecords[key] = {
          punchIn: [],
          punchOut: [],
          breakBegin: [],
          breakEnd: []
        };
      }
      
      switch(rowType) {
        case '出勤':
          dailyRecords[key].punchIn.push(dateObj);
          break;
        case '退勤':
          dailyRecords[key].punchOut.push(dateObj);
          break;
        case '休憩開始':
          dailyRecords[key].breakBegin.push(dateObj);
          break;
        case '休憩終了':
          dailyRecords[key].breakEnd.push(dateObj);
          break;
      }
    }
  }
  
  // 不整合チェック
  var inconsistencies = [];
  for (var dateKey in dailyRecords) {
    var record = dailyRecords[dateKey];
    // 出勤があるのに退勤が無い
    if (record.punchIn.length > 0 && record.punchOut.length == 0) {
      inconsistencies.push({
        date: dateKey,
        message: "出勤打刻はあるが退勤打刻がありません。"
      });
    }
    // 休憩開始があるのに休憩終了が無い（ざっくりチェック）
    if (record.breakBegin.length > 0 && record.breakEnd.length == 0) {
      inconsistencies.push({
        date: dateKey,
        message: "休憩開始がありますが、休憩終了がありません。"
      });
    }
    // ※必要に応じて他のパターンも追加可能
  }
  
  return inconsistencies;
}

/**
 * 【B】全従業員の打刻漏れをリストアップして返す
 */
function getInconsistencyList() {
  var emps = getEmployees();
  var result = [];
  emps.forEach(function(emp) {
    var empId = emp.id;
    var empName = emp.name;
    var inconst = checkInconsistenciesForEmp(empId);
    if(inconst.length > 0) {
      inconst.forEach(function(ic) {
        result.push({
          empId: empId,
          empName: empName,
          date: ic.date,
          message: ic.message
        });
      });
    }
  });
  return result;
}

/**
 * 【C】打刻漏れアラートの結果を「チェック結果」シートに記録する
 */
function updateCheckResultsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("チェック結果");
  
  if (!sheet) {
    sheet = ss.insertSheet("チェック結果");
    sheet.appendRow(["従業員ID", "従業員名", "日付", "エラー内容"]);
  }
  
  // シートをクリア（毎回最新データだけを記録）
  sheet.getRange("A2:D").clearContent();

  var inconsistencies = getInconsistencyList(); // 打刻漏れリストを取得

  for (var i = 0; i < inconsistencies.length; i++) {
    var row = inconsistencies[i];
    sheet.appendRow([row.empId, row.empName, row.date, row.message]);
  }
}

/**
 * 【D】勤怠データをCSV形式で出力し、Googleドライブに保存
 */
function exportAttendanceToCSV() {
  var folderName = "勤怠データ"; // Googleドライブ内のフォルダ名
  var fileName = "attendance_" + Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd_HHmmss") + ".csv";
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("打刻履歴"); // 勤怠履歴シート
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return "データがありません。";
  }

  var csvContent = [];
  for (var i = 0; i < data.length; i++) {
    csvContent.push(data[i].join(","));
  }
  
  var csvBlob = Utilities.newBlob(csvContent.join("\n"), "text/csv", fileName);
  
  // Googleドライブ内のフォルダを取得または作成
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  
  // CSVファイルをGoogleドライブに保存
  var file = folder.createFile(csvBlob);
  return file.getUrl(); // ファイルのURLを返す
}

/**
 * 【E】シフトと実績の比較チェック
 * - 遅刻・早退・未出勤・無断残業のチェック
 */
function checkShiftVsActual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shiftSheet = ss.getSheetByName("シフト表");  // 予定シフト
  var timeSheet = ss.getSheetByName("打刻履歴");   // 実際の打刻
  var checkSheet = ss.getSheetByName("チェック結果"); // 結果シート

  if (!shiftSheet || !timeSheet) {
    return "シフト表または打刻履歴シートが見つかりません。";
  }
  
  // 結果シートのクリア＆ヘッダーセット
  if (!checkSheet) {
    checkSheet = ss.insertSheet("チェック結果");
    checkSheet.appendRow(["従業員ID", "従業員名", "日付", "エラー内容"]);
  } else {
    checkSheet.getRange("A2:D").clearContent();
  }

  // シフトデータを取得
  var shiftData = shiftSheet.getDataRange().getValues();
  var timeData = timeSheet.getDataRange().getValues();

  var checkResults = [];

  // シフトデータを走査
  for (var i = 1; i < shiftData.length; i++) {
    var empId = shiftData[i][0];
    var date = shiftData[i][1];
    var shiftStart = shiftData[i][2];
    var shiftEnd = shiftData[i][3];

    if (!empId || !date) continue;

    var actualStart = null;
    var actualEnd = null;

    // 実際の打刻データを検索
    for (var j = 1; j < timeData.length; j++) {
      if (timeData[j][0] == empId) {
        var punchType = timeData[j][1];
        var punchTime = new Date(timeData[j][2]);

        var punchDate = Utilities.formatDate(punchTime, "Asia/Tokyo", "yyyy-MM-dd");
        if (punchDate == date) {
          if (punchType == "出勤") actualStart = punchTime;
          if (punchType == "退勤") actualEnd = punchTime;
        }
      }
    }

    var issues = [];

    // 遅刻
    if (actualStart && actualStart > shiftStart) {
      issues.push("遅刻 (" + Utilities.formatDate(actualStart, "Asia/Tokyo", "HH:mm") + " 出勤)");
    }
    // 早退
    if (actualEnd && actualEnd < shiftEnd) {
      issues.push("早退 (" + Utilities.formatDate(actualEnd, "Asia/Tokyo", "HH:mm") + " 退勤)");
    }
    // 未出勤
    if (!actualStart && !actualEnd) {
      issues.push("未出勤");
    }
    // 無断残業
    if (actualEnd && actualEnd > shiftEnd.setMinutes(shiftEnd.getMinutes() + 30)) {
      issues.push("無断残業 (" + Utilities.formatDate(actualEnd, "Asia/Tokyo", "HH:mm") + " 退勤)");
    }

    if (issues.length > 0) {
      checkResults.push([empId, getEmployeeNameById(empId), date, issues.join(", ")]);
    }
  }

  // チェック結果をシートに書き込む
  if (checkResults.length > 0) {
    checkSheet.getRange(2, 1, checkResults.length, 4).setValues(checkResults);
  }
}

/**
 * 【F】従業員IDから従業員名を取得
 */
function getEmployeeNameById(empId) {
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿");
  var data = empSheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == empId) {
      return data[i][1];  // 従業員名を返す
    }
  }
  return "不明";
}

/**
 * 【G】Gmailでシフトと実績の不一致を通知
 */
function sendEmailNotification() {
  var checkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("チェック結果");
  if (!checkSheet) return;

  var data = checkSheet.getDataRange().getValues();
  if (data.length <= 1) return; // データがない場合は何もしない

  var recipient = "owner_email@example.com"; // 送信先（オーナーのメールアドレス）
  var subject = "【勤怠管理】シフトと実績の不一致アラート";
  var body = "以下の従業員に勤怠の不一致がありました。\n\n";

  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0];
    var empName = data[i][1];
    var date = data[i][2];
    var issue = data[i][3];
    if (empId && empName && date && issue) {
      body += `従業員ID: ${empId}\n`;
      body += `名前: ${empName}\n`;
      body += `日付: ${date}\n`;
      body += `エラー内容: ${issue}\n\n`;
    }
  }

  body += "詳細はアプリの「チェック結果」シートをご確認ください。\n";

  // Gmailで送信
  GmailApp.sendEmail(recipient, subject, body);
}

/**
 * 【H】管理者設定シートからメールアドレスを取得
 */
function getOwnerEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("管理者設定");

  if (!sheet) {
    sheet = ss.insertSheet("管理者設定");
    sheet.appendRow(["オーナーのメールアドレス", ""]);
    return "";
  }

  var email = sheet.getRange(1, 2).getValue();
  return email || "";
}

/**
 * 【I】管理者設定シートにメールアドレスを保存
 */
function setOwnerEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("管理者設定");

  if (!sheet) {
    sheet = ss.insertSheet("管理者設定");
    sheet.appendRow(["オーナーのメールアドレス", email]);
  } else {
    sheet.getRange(1, 2).setValue(email);
  }

  return "メールアドレスを更新しました: " + email;
}

/**
 * 【J】Gmailでシフトと実績の不一致を通知（オーナーのメールアドレスを取得）
 */
function sendEmailNotification() {
  var recipient = getOwnerEmail();
  if (!recipient) {
    Logger.log("オーナーのメールアドレスが設定されていません。");
    return;
  }

  var checkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("チェック結果");
  if (!checkSheet) return;

  var data = checkSheet.getDataRange().getValues();
  if (data.length <= 1) return; // データがない場合は何もしない

  var subject = "【勤怠管理】シフトと実績の不一致アラート";
  var body = "以下の従業員に勤怠の不一致がありました。\n\n";

  for (var i = 1; i < data.length; i++) {
    var empId = data[i][0];
    var empName = data[i][1];
    var date = data[i][2];
    var issue = data[i][3];
    if (empId && empName && date && issue) {
      body += `従業員ID: ${empId}\n`;
      body += `名前: ${empName}\n`;
      body += `日付: ${date}\n`;
      body += `エラー内容: ${issue}\n\n`;
    }
  }

  body += "詳細はアプリの「チェック結果」シートをご確認ください。\n";

  // Gmailで送信
  GmailApp.sendEmail(recipient, subject, body);
}

/**
 * 📌 新しい従業員を追加する（UUIDベースのユニークID）
 * - UUID（またはタイムスタンプ）を使用して、二度と重複しないIDを作成
 */
function addEmployee(empName, empJob, maxAttendanceCount, paidVacationLimit, paidVacationGrantDate) {
  if (!empName || empName.trim() === "" || !empJob || empJob.trim() === "" ||
      isNaN(maxAttendanceCount) || isNaN(paidVacationLimit) || !paidVacationGrantDate) {
    return JSON.stringify({ success: false, message: "エラー: 従業員名、職種、出勤回数上限、有給休暇上限、有給休暇付与日を正しく入力してください。" });
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("従業員名簿");
  if (!sheet) {
    // ヘッダー：従業員ID, 名前, 職種, 出勤回数上限, 有給休暇上限, 有給休暇付与日
    sheet = ss.insertSheet("従業員名簿");
    sheet.appendRow(["従業員ID", "名前", "職種", "出勤回数上限", "有給休暇上限", "有給休暇付与日"]);
  }
  var newEmpId = generateEmployeeId(empJob);
  sheet.appendRow([newEmpId, empName, empJob, maxAttendanceCount, paidVacationLimit, paidVacationGrantDate]);
  return JSON.stringify({
    success: true,
    empId: newEmpId,
    empName: empName,
    message: `従業員 ${empName} (ID: ${newEmpId}) を追加しました。`
  });
}



/**
 * 📌 ユニークなIDを生成する関数
 * - UUIDのようなランダムな英数字ID
 * - `Math.random()` + `タイムスタンプ` を組み合わせて衝突を防ぐ
 */
function generateUniqueId() {
  var timestamp = new Date().getTime(); // 現在の時間 (ミリ秒)
  var randomPart = Math.floor(Math.random() * 1000000); // 6桁のランダムな数字
  return "EMP-" + timestamp.toString(36) + "-" + randomPart.toString(36);
}



/**
 * 📌 従業員を削除する
 * 指定した従業員IDの行を「従業員名簿」シートから削除
 */
function deleteEmployee(empId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿");
  if (!sheet) return JSON.stringify({ success: false, message: "エラー: 従業員名簿シートが見つかりません。" });

  var data = sheet.getDataRange().getValues();
  var targetRow = -1;

  for (var i = 1; i < data.length; i++) { // 1行目はヘッダーなのでスキップ
    if (data[i][0].toString() === empId.toString()) {
      targetRow = i + 1; // Google Sheets の行番号は 1始まり
      break;
    }
  }

  if (targetRow !== -1) {
    sheet.deleteRow(targetRow);
    return JSON.stringify({ success: true, empId: empId, message: `従業員ID: ${empId} を削除しました。` });
  } else {
    return JSON.stringify({ success: false, message: "エラー: 従業員が見つかりませんでした。" });
  }
}

/**
 * 📌 勤怠メモを保存・更新する
 * - メモは「従業員ID」「日時」「種別（出勤・退勤など）」に紐付ける
 */
function saveTimeClockMemo(form) {
  var empId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!empId) {
    return JSON.stringify({ success: false, message: "エラー: 従業員IDが選択されていません。" });
  }

  var targetDate = form.target_date; // 例："2025-02-12"
  var targetType = form.target_type;
  var memo = form.memo_text;

  if (!targetDate || !targetType) {
    return JSON.stringify({ success: false, message: "エラー: 日付または種別が未入力です。" });
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("勤怠メモ");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("勤怠メモ");
    sheet.appendRow(["従業員ID", "日付", "種別", "メモ"]); // ヘッダー作成
  }

  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  
  // ※ 既存の行を探す際に、スプレッドシートの日付を「yyyy-MM-dd」に変換して比較する
  for (var i = 1; i < data.length; i++) {
    var sheetDate = Utilities.formatDate(new Date(data[i][1]), "Asia/Tokyo", "yyyy-MM-dd");
    if (data[i][0] == empId && sheetDate == targetDate && data[i][2] == targetType) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow !== -1) {
    sheet.getRange(targetRow, 4).setValue(memo);
  } else {
    // 新規登録時、日付は文字列のまま書き込むと自動で Date 型に変換される場合もあるので注意
    sheet.appendRow([empId, targetDate, targetType, memo]);
  }

  return JSON.stringify({ success: true, message: "メモを保存しました。" });
}



/**
 * 📌 指定された従業員の勤怠メモを取得
 * - `date` のフォーマットを「yyyy-MM-dd」に統一（getTimeClocks()と一致させる）
 */
function getTimeClockMemos() {
  var empId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!empId) return JSON.stringify({});

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("勤怠メモ");
  if (!sheet) return JSON.stringify({});

  var data = sheet.getDataRange().getValues();
  var memos = {};

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == empId) {
      // 日付データを「yyyy-MM-dd」に変換してキーに利用
      var dateKey = Utilities.formatDate(new Date(data[i][1]), "Asia/Tokyo", "yyyy-MM-dd");
      var type = data[i][2];
      var key = dateKey + "-" + type;
      memos[key] = data[i][3] || ""; 
    }
  }

  return JSON.stringify(memos);
}




/***********************************
 * 追加機能: 月次集計
 ***********************************/

/**
 * ざっくりと今月の合計勤務時間・休憩時間を集計
 * - ロジックは簡易的な例です。(出勤→退勤の差分を勤務時間に、休憩開始→休憩終了の差分を休憩時間に加算)
 * - 一日の中で複数回の休憩がある場合など複雑なケースには要追加実装
 */
function getMonthlySummary() {
  var emps = getEmployees();
  var summaryList = [];
  
  var now = new Date();
  var firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  
  emps.forEach(function(emp) {
    var empId = emp.id;
    var empName = emp.name;
    
    // 打刻記録を取得（時系列ソート済み）
    var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
    var last_row = timeClocksSheet.getLastRow();
    var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row - 1, 3);
    var data = timeClocksRange.getValues();
    
    // 各日の最初の出勤と最後の退勤をまとめる
    var daysWorked = {};  // { "yyyy-MM-dd": { in: Date, out: Date } }
    var paidVacationDaysTaken = {}; // 有給休暇取得日の記録
    
    data.forEach(function(row) {
      var recEmpId = row[0].toString();
      if (recEmpId !== empId) return;
      
      var type = row[1];
      var dt = new Date(row[2]);
      if (isNaN(dt.getTime()) || dt < firstDayOfMonth || dt > now) return;
      
      var dateStr = Utilities.formatDate(dt, "Asia/Tokyo", "yyyy-MM-dd");
      
      if (type === "有給休暇") {
        paidVacationDaysTaken[dateStr] = true;
        return;
      }
      
      if (!daysWorked[dateStr]) {
        daysWorked[dateStr] = { in: null, out: null };
      }
      if (type === "出勤") {
        if (!daysWorked[dateStr].in || dt < daysWorked[dateStr].in) {
          daysWorked[dateStr].in = dt;
        }
      }
      if (type === "退勤") {
        if (!daysWorked[dateStr].out || dt > daysWorked[dateStr].out) {
          daysWorked[dateStr].out = dt;
        }
      }
    });
    
    var totalWorkMin = 0;
    var totalNightShiftMin = 0;
    var attendanceDays = 0;
    var holidayWorkMin = 0;
    
    for (var day in daysWorked) {
      var rec = daysWorked[day];
      if (rec.in && rec.out) {
        var sessionMinutes = (rec.out - rec.in) / 1000 / 60;
        var breakDeduction = 0;
        if (sessionMinutes >= 360 && sessionMinutes < 480) {
          breakDeduction = 45;
        } else if (sessionMinutes >= 480) {
          breakDeduction = 60;
        }
        var effectiveMinutes = sessionMinutes - breakDeduction;
        totalWorkMin += effectiveMinutes;
        attendanceDays++;
        
        var d = new Date(day);
        var dayOfWeek = d.getDay();
        if (dayOfWeek === 0 || dayOfWeek === 6) {
          holidayWorkMin += effectiveMinutes;
        }
        var inHour = rec.in.getHours();
        if (inHour >= 22 || inHour < 6) {
          totalNightShiftMin += effectiveMinutes;
        }
      }
    }
    
    var paidVacationTakenCount = Object.keys(paidVacationDaysTaken).length;
    var allocatedPaidVacation = parseInt(emp.paidVacationLimit);
    if (isNaN(allocatedPaidVacation)) allocatedPaidVacation = 0;
    var remainingPaidVacation = allocatedPaidVacation - paidVacationTakenCount;
    
    // 追加：残り出勤回数 = maxAttendanceCount - 出勤日数
    var remainingAttendanceCount = emp.maxAttendanceCount - attendanceDays;
    
    summaryList.push({
      empId: empId,
      empName: empName,
      totalWorkMin: totalWorkMin,
      totalNightShiftMin: totalNightShiftMin,
      attendanceDays: attendanceDays,
      remainingAttendanceCount: remainingAttendanceCount, // 追加
      holidayWorkMin: holidayWorkMin,
      paidVacationTaken: paidVacationTakenCount,
      remainingPaidVacation: remainingPaidVacation,
      maxAttendanceCount: emp.maxAttendanceCount // 必要に応じて
    });
  });
  
  return summaryList;
}


/**
 * 新しい職種を追加する関数
 * 職種コードは2文字の大文字アルファベットであることをチェック
 */
function addJobType(code, name) {
  // 入力値のトリム
  code = code.trim();
  name = name.trim();
  
  // 2文字の大文字アルファベットかチェック
  if (!/^[A-Z]{2}$/.test(code)) {
    return "エラー: 職種コードは2文字の大文字アルファベットで入力してください。";
  }
  if (!name) {
    return "エラー: 職種名を入力してください。";
  }
  
  // 職種一覧シート（例："職種一覧"）を取得または作成
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("職種一覧");
  if (!sheet) {
    sheet = ss.insertSheet("職種一覧");
    // ヘッダーを作成
    sheet.appendRow(["職種コード", "職種名"]);
  }
  
  // 同じコードが既に存在していないかチェック
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) { // ヘッダーを除く
    if (data[i][0] && data[i][0].toString().trim() === code) {
      return "エラー: 職種コード「" + code + "」は既に存在しています。";
    }
  }
  
  // 新規追加
  sheet.appendRow([code, name]);
  
  return "職種「" + name + "」（コード：" + code + "）を追加しました。";
}


/**
 * 職種一覧シートから職種情報を取得する関数
 */
function getJobTypes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("職種一覧");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  var jobTypes = [];
  // ヘッダー行を除く
  for (var i = 1; i < data.length; i++) {
    var code = data[i][0];
    var name = data[i][1];
    if (code && name) {
      jobTypes.push({
        code: code.toString().trim(),
        name: name.toString().trim()
      });
    }
  }
  return jobTypes;
}


/**
 * 選択中の従業員の職種を取得する関数
 */
function getEmployeeJob() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!selectedEmpId) return "";
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("従業員名簿");
  if (!empSheet) return "";
  
  var lastRow = empSheet.getLastRow();
  if (lastRow < 2) return "";
  
  // 3列目（職種）も取得する
  var empData = empSheet.getRange(2, 1, lastRow - 1, 3).getValues();
  for (var i = 0; i < empData.length; i++) {
    var empId = empData[i][0];
    var job = empData[i][2]; // 3列目に職種が格納されている前提
    if (empId && empId.toString() === selectedEmpId.toString()) {
      return job;
    }
  }
  return "";
}


/**
 * 指定された職種コード(empJob)を先頭に、6桁の数字を連結して従業員番号を生成する関数
 */
function generateEmployeeId(empJob) {
  // 0～999,999 のランダムな整数を生成し、6桁の文字列に変換
  var num = Math.floor(Math.random() * 1000000);
  var numStr = num.toString().padStart(6, '0');
  return empJob + numStr;
}


/**
 * 指定された設定項目の値を取得する関数
 */
function getSetting(settingName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("店舗設定");
  if (!sheet) {
    // シートがなければ作成してヘッダーをセット
    sheet = ss.insertSheet("店舗設定");
    sheet.appendRow(["設定項目", "値"]);
    return "";
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === settingName) {
      return data[i][1];
    }
  }
  return "";
}

/**
 * 指定された設定項目の値を更新または新規追加する関数
 */
function setSetting(settingName, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("店舗設定");
  if (!sheet) {
    sheet = ss.insertSheet("店舗設定");
    sheet.appendRow(["設定項目", "値"]);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim() === settingName) {
      sheet.getRange(i + 1, 2).setValue(value);
      return settingName + "を更新しました。";
    }
  }
  sheet.appendRow([settingName, value]);
  return settingName + "を設定しました。";
}

/**
 * 有給休暇数の設定を更新する関数
 */
function setPaidVacationDays(days) {
  if (isNaN(days) || days < 0) {
    return "エラー: 有給休暇数は0以上の数字を入力してください。";
  }
  return setSetting("有給休暇数", days);
}

/**
 * 有給休暇数の設定値を取得する関数
 */
function getPaidVacationDays() {
  var val = getSetting("有給休暇数");
  return val ? val : "";
}

/**
 * 出勤回数上限の設定を更新する関数
 */
function setMaxAttendanceCount(count) {
  if (isNaN(count) || count < 0) {
    return "エラー: 出勤回数上限は0以上の数字を入力してください。";
  }
  return setSetting("出勤回数上限", count);
}

/**
 * 出勤回数上限の設定値を取得する関数
 */
function getMaxAttendanceCount() {
  var val = getSetting("出勤回数上限");
  return val ? val : "";
}


/**
 * 指定された月（形式 "YYYY-MM"）の打刻記録を取得する関数
 */
function getTimeClocksForMonth(monthStr) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!selectedEmpId) return [];
  
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  if (!timeClocksSheet) return [];
  
  var last_row = timeClocksSheet.getLastRow();
  if (last_row < 2) return [];
  
  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row - 1, 3);
  var data = timeClocksRange.getValues();
  var empTimeClocks = [];
  
  // monthStr is expected in "YYYY-MM" format.
  var parts = monthStr.split("-");
  var year = parseInt(parts[0], 10);
  var month = parseInt(parts[1], 10); // 1～12
  var startDate = new Date(year, month - 1, 1);
  // endDate: 最終日の23:59:59
  var endDate = new Date(year, month, 0, 23, 59, 59);
  
  for (var i = 0; i < data.length; i++) {
    var empId = data[i][0].toString();
    var type = data[i][1];
    var datetime = new Date(data[i][2]);
    if (!empId || empId.trim() === "" || empId !== selectedEmpId) continue;
    if (isNaN(datetime.getTime()) || datetime < startDate || datetime > endDate) continue;
    var formattedDate = Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd");
    var formattedDateTime = Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd HH:mm");
    empTimeClocks.push({
      'date': formattedDate,         // キー作成用は日付のみ
      'datetime': formattedDateTime, // 表示用の日時
      'type': type,
      'rawDateTime': datetime
    });
  }
  
  empTimeClocks.sort(function(a, b) {
    return a.rawDateTime - b.rawDateTime;
  });
  return empTimeClocks;
}


/**
 * 個別の月次集計結果をグループ別（従業員IDの先頭2文字＝職種コード）にまとめて合算する関数
 */
function getMonthlySummaryByJob() {
  var summaries = getMonthlySummary(); // 各従業員の集計結果
  var grouped = {};

  summaries.forEach(function(summary) {
    var jobCode = summary.empId.substring(0, 2);
    if (!grouped[jobCode]) {
      grouped[jobCode] = {
        jobCode: jobCode,
        totalWorkMin: 0,
        totalNightShiftMin: 0,
        attendanceDays: 0,
        holidayWorkMin: 0,
        paidVacationTaken: 0,
        remainingPaidVacation: 0,
        maxAttendanceCountSum: 0,
        attendanceDaysSum: 0,
        remainingAttendanceCount: 0
      };
    }
    grouped[jobCode].totalWorkMin += summary.totalWorkMin;
    grouped[jobCode].totalNightShiftMin += summary.totalNightShiftMin;
    grouped[jobCode].attendanceDays += summary.attendanceDays;
    grouped[jobCode].holidayWorkMin += summary.holidayWorkMin;
    grouped[jobCode].paidVacationTaken += summary.paidVacationTaken;
    grouped[jobCode].remainingPaidVacation += summary.remainingPaidVacation;

    // 合計の上限出勤回数・実際出勤日数を加算
    grouped[jobCode].maxAttendanceCountSum += summary.maxAttendanceCount;
    grouped[jobCode].attendanceDaysSum     += summary.attendanceDays;
  });

  // 各職種の残り出勤回数を計算
  for (var code in grouped) {
    grouped[code].remainingAttendanceCount =
        grouped[code].maxAttendanceCountSum - grouped[code].attendanceDaysSum;
  }

  var result = [];
  for (var jobCode in grouped) {
    result.push(grouped[jobCode]);
  }

  // 各グループに対して、職種名を付与
  var jobTypes = getJobTypes();
  result.forEach(function(item) {
    for (var i = 0; i < jobTypes.length; i++) {
      if (jobTypes[i].code === item.jobCode) {
        item.jobName = jobTypes[i].name;
        break;
      }
    }
    if (!item.jobName) {
      item.jobName = "";
    }
  });

  return result;
}



// レポート生成時に「残り出勤回数」を出力するように修正
function getReportDataArray(reportType) {
  if (reportType === "job") {
    var header = ["職種コード", "職種名", "合計勤務時間(分)", "合計夜勤時間(分)", "出勤日数", "残り出勤回数", "休日出勤の時間(分)", "有給取得日数", "残り有給日数"];
    var grouped = getMonthlySummaryByJob();
    var data = [header];
    grouped.forEach(function(row) {
      data.push([
         row.jobCode,
         row.jobName,
         Math.round(row.totalWorkMin),
         Math.round(row.totalNightShiftMin),
         row.attendanceDays,
         row.remainingAttendanceCount,  // ここで追加
         Math.round(row.holidayWorkMin),
         row.paidVacationTaken,
         row.remainingPaidVacation
      ]);
    });
    return data;
  } else {
    var header = ["従業員ID", "従業員名", "合計勤務時間(分)", "合計夜勤時間(分)", "出勤日数", "残り出勤回数", "休日出勤の時間(分)", "有給取得日数", "残り有給日数"];
    var monthly = getMonthlySummary();
    var data = [header];
    monthly.forEach(function(row) {
      data.push([
         row.empId,
         row.empName,
         Math.round(row.totalWorkMin),
         Math.round(row.totalNightShiftMin),
         row.attendanceDays,
         row.remainingAttendanceCount,  // ここで追加
         Math.round(row.holidayWorkMin),
         row.paidVacationTaken,
         row.remainingPaidVacation
      ]);
    });
    return data;
  }
}




// reportType: "monthly" なら個別月次レポート、"job" なら職種別レポート
// function getReportDataArray(reportType) {
//   if (reportType === "job") {
//     // 職種別レポート用
//     var header = [
//       "職種コード", 
//       "職種名", 
//       "合計勤務時間(分)", 
//       "合計夜勤時間(分)", 
//       "出勤日数", 
//       "残り出勤回数", 
//       "休日出勤の時間(分)", 
//       "有給取得日数", 
//       "残り有給日数"
//     ];

//     var grouped = getMonthlySummaryByJob();
//     var data = [header];

//     grouped.forEach(function(row) {
//       data.push([
//         row.jobCode,
//         row.jobName,
//         Math.round(row.totalWorkMin),
//         Math.round(row.totalNightShiftMin),
//         row.attendanceDays,
//         row.remainingAttendance,     // ←ここで残り出勤回数を出力
//         Math.round(row.holidayWorkMin),
//         row.paidVacationTaken,
//         row.remainingPaidVacation
//       ]);
//     });
//     return data;
//   } else {
//     // 個人別レポート用
//     var header = [
//       "従業員ID", 
//       "従業員名", 
//       "合計勤務時間(分)", 
//       "合計夜勤時間(分)", 
//       "出勤日数", 
//       "残り出勤回数", 
//       "休日出勤の時間(分)", 
//       "有給取得日数", 
//       "残り有給日数"
//     ];

//     var monthly = getMonthlySummary();
//     var data = [header];

//     monthly.forEach(function(row) {
//       data.push([
//         row.empId,
//         row.empName,
//         Math.round(row.totalWorkMin),
//         Math.round(row.totalNightShiftMin),
//         row.attendanceDays,
//         row.remainingAttendanceCount,  // こちらは個人の残り出勤回数
//         Math.round(row.holidayWorkMin),
//         row.paidVacationTaken,
//         row.remainingPaidVacation
//       ]);
//     });
//     return data;
//   }
// }



function exportReportToXLSX(reportType) {
  var data = getReportDataArray(reportType);
  var ss = SpreadsheetApp.create("Temp Report XLSX");
  var sheet = ss.getActiveSheet();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  var fileId = ss.getId();
  var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + fileId + "&exportFormat=xlsx";
  // 一時ファイルはゴミ箱へ移動
  DriveApp.getFileById(fileId).setTrashed(true);
  return url;
}

function exportReportToODS(reportType) {
  var data = getReportDataArray(reportType);
  var ss = SpreadsheetApp.create("Temp Report ODS");
  var sheet = ss.getActiveSheet();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  var fileId = ss.getId();
  var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + fileId + "&exportFormat=ods";
  DriveApp.getFileById(fileId).setTrashed(true);
  return url;
}

function exportReportToPDF(reportType) {
  var data = getReportDataArray(reportType);
  var ss = SpreadsheetApp.create("Temp Report PDF");
  var sheet = ss.getActiveSheet();
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  var fileId = ss.getId();
  var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + fileId + "&exportFormat=pdf";
  DriveApp.getFileById(fileId).setTrashed(true);
  return url;
}

function exportReportToHTML(reportType) {
  var data = getReportDataArray(reportType);
  // シンプルなHTMLテーブルに変換
  var html = '<html><head><meta charset="UTF-8"><title>Report</title></head><body><table border="1">';
  data.forEach(function(row) {
    html += "<tr><td>" + row.join("</td><td>") + "</td></tr>";
  });
  html += "</table></body></html>";
  var blob = Utilities.newBlob(html, "text/html", "report.html");
  var file = DriveApp.createFile(blob);
  return file.getUrl();
}

function exportReportToCSV(reportType) {
  var data = getReportDataArray(reportType);
  var csvContent = data.map(function(row) {
    return row.join(",");
  }).join("\n");
  var blob = Utilities.newBlob(csvContent, "text/csv", "report.csv");
  var file = DriveApp.createFile(blob);
  return file.getUrl();
}

function exportReportToTSV(reportType) {
  var data = getReportDataArray(reportType);
  var tsvContent = data.map(function(row) {
    return row.join("\t");
  }).join("\n");
  var blob = Utilities.newBlob(tsvContent, "text/tab-separated-values", "report.tsv");
  var file = DriveApp.createFile(blob);
  return file.getUrl();
}



function setSelectedEmployee(empId) {
  PropertiesService.getUserProperties().setProperty('selectedEmpId', empId);
  return "従業員 " + empId + " が選択されました。";
}


/**
 * 管理者設定シートにメールアドレスを保存
 */
function setOwnerEmail(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("管理者設定");

  if (!sheet) {
    sheet = ss.insertSheet("管理者設定");
    sheet.appendRow(["オーナーのメールアドレス", email]);
  } else {
    sheet.getRange(1, 2).setValue(email);  // Set the email in the second column
  }

  return "メールアドレスを更新しました: " + email;
}


/**
 * 現在設定されているオーナーのメールアドレスを取得
 */
function getOwnerEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("管理者設定");

  if (!sheet) return "";

  var email = sheet.getRange(1, 2).getValue();
  return email || "";
}



/**
 * オーナーのパスワードを設定
 */
function setOwnerPassword(password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("管理者設定");

  if (!sheet) {
    sheet = ss.insertSheet("管理者設定");
    sheet.appendRow(["オーナーのパスワード", password]);
  } else {
    sheet.getRange(1, 2).setValue(password);
  }

  return "パスワードを更新しました。";
}

/**
 * 現在設定されているオーナーのパスワードを取得
 */
function getOwnerPassword() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("管理者設定");

  if (!sheet) return "";

  var password = sheet.getRange(1, 2).getValue();
  return password || "";
}


/**
 * パスワードを確認する関数
 */
function verifyPassword(inputPassword) {
  var storedPassword = getOwnerPassword();  // 現在設定されているオーナーのパスワードを取得
  if (inputPassword === storedPassword) {
    return { success: true };  // パスワード一致
  } else {
    return { success: false }; // パスワード不一致
  }
}



/**
 * 【新規】有給休暇申請をシート「有給申請」に保存する関数
 */
function submitPaidVacationRequest(form) {
  var empId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!empId) {
    return "エラー: 従業員が選択されていません。";
  }
  var targetDate = form.target_date;
  if (!targetDate) {
    return "エラー: 対象日が未入力です。";
  }
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("有給申請");
  if (!sheet) {
    sheet = ss.insertSheet("有給申請");
    // ヘッダー：RequestID, 従業員ID, 対象日, 提出日時
    sheet.appendRow(["RequestID", "従業員ID", "対象日", "提出日時"]);
  }
  var requestId = "REQ-" + new Date().getTime();
  var submitTime = new Date();
  sheet.appendRow([requestId, empId, targetDate, submitTime]);
  return "有給休暇申請を送信しました。";
}

/**
 * 【新規】有給休暇申請一覧を取得する関数
 */
function getPaidVacationRequests() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("有給申請");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var requests = [];
  // ヘッダー行をスキップ
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // row: [RequestID, 従業員ID, 対象日, 提出日時]
    var empName = getEmployeeNameById(row[1]);
    requests.push({
      requestId: row[0],
      empId: row[1],
      empName: empName,
      targetDate: row[2],
      submitTime: row[3]
    });
  }
  return requests;
}

/**
 * 【新規】有給休暇申請を受諾する関数
 */
function approvePaidVacation(requestId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("有給申請");
  if (!sheet) return "有給申請シートが見つかりません。";
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  var requestData = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == requestId) {
      targetRow = i + 1;
      requestData = data[i];
      break;
    }
  }
  if (targetRow == -1) return "該当する申請が見つかりません。";
  
  var empId = requestData[1];
  var targetDate = requestData[2];
  
  // 有給休暇として打刻履歴シートに登録
  var workSheet = ss.getSheetByName("打刻履歴");
  if (!workSheet) return "打刻履歴シートが見つかりません。";
  var newRow = workSheet.getLastRow() + 1;
  var timestamp = new Date(targetDate + "T00:00:00");
  workSheet.getRange(newRow, 1).setValue(empId);
  workSheet.getRange(newRow, 2).setValue("有給休暇");
  workSheet.getRange(newRow, 3).setValue(timestamp);
  
  // 申請レコードを削除
  sheet.deleteRow(targetRow);
  return "有給休暇申請が承認されました。";
}

/**
 * 【新規】有給休暇申請を拒否する関数
 */
function rejectPaidVacation(requestId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("有給申請");
  if (!sheet) return "有給申請シートが見つかりません。";
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == requestId) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow == -1) return "該当する申請が見つかりません。";
  
  sheet.deleteRow(targetRow);
  return "有給休暇申請が拒否されました。";
}



/**
 * 【新規】有給休暇申請を受諾する関数
 */
function approvePaidVacation(requestId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("有給申請");
  if (!sheet) return JSON.stringify({ success: false, message: "有給申請シートが見つかりません。" });
  var data = sheet.getDataRange().getValues();
  var targetRow = -1;
  var requestData = null;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == requestId) {
      targetRow = i + 1;
      requestData = data[i];
      break;
    }
  }
  if (targetRow == -1) return JSON.stringify({ success: false, message: "該当する申請が見つかりません。" });
  
  var empId = requestData[1];
  var targetDate = requestData[2]; // 例："2025-02-15"
  
  // 有給休暇として打刻履歴シートに登録
  var workSheet = ss.getSheetByName("打刻履歴");
  if (!workSheet) return JSON.stringify({ success: false, message: "打刻履歴シートが見つかりません。" });
  var newRow = workSheet.getLastRow() + 1;
  // ※タイムは 00:00:00 として登録（必要に応じて変更）
  var timestamp = new Date(targetDate + "T00:00:00");
  workSheet.getRange(newRow, 1).setValue(empId);
  workSheet.getRange(newRow, 2).setValue("有給休暇");
  workSheet.getRange(newRow, 3).setValue(timestamp);
  
  // 申請レコードを削除
  sheet.deleteRow(targetRow);
  
  // 対象日の月を "YYYY-MM" 形式に変換
  var dateObj = new Date(targetDate);
  var monthStr = Utilities.formatDate(dateObj, "Asia/Tokyo", "yyyy-MM");
  
  return JSON.stringify({ success: true, message: "有給休暇申請が承認されました。", empId: empId, month: monthStr });
}




/**
 * タイムレコーダーのレコードを削除する関数
 * @param {string} datetimeStr 「yyyy-MM-dd HH:mm」形式の日時文字列
 * @param {string} type レコードの種別（例："出勤", "退勤", "有給休暇"）
 * @return {string} 結果メッセージ
 */
function deleteTimeClockRecord(datetimeStr, type) {
  var empId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  if (!empId) return "エラー: 従業員が選択されていません。";

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("打刻履歴");
  if (!sheet) return "エラー: 打刻履歴シートが見つかりません。";
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return "エラー: 削除対象のレコードがありません。";

  // 2行目以降の全データを取得
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var targetRow = null;
  for (var i = 0; i < data.length; i++) {
    var rowEmpId = data[i][0].toString();
    var rowType = data[i][1];
    var rowTimestamp = data[i][2];
    // 対象従業員かつ種別が一致しているか確認
    if (rowEmpId === empId && rowType === type) {
      // シート上の日時を "yyyy-MM-dd HH:mm" 形式に変換
      var rowDateTimeStr = Utilities.formatDate(new Date(rowTimestamp), "Asia/Tokyo", "yyyy-MM-dd HH:mm");
      if (rowDateTimeStr === datetimeStr) {
        targetRow = i + 2;  // データは2行目から始まっているため
        break;
      }
    }
  }

  if (targetRow != null) {
    sheet.deleteRow(targetRow);
    return "レコードを削除しました。";
  } else {
    return "レコードが見つかりませんでした。";
  }
}

