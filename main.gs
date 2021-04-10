function onOpen(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.addMenu("スクリプト", [{name: "振休申請書作成", functionName: "main"}]);
}

function test() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const querySheet = ss.getSheetByName("QUERY");
  const sh = "2104"
  const d = 13
  const query = `=query('${sh}'!B10:J40, "select * where B = ${d}")`;
  querySheet.getRange('A2').setFormula(query);
  Logger.log(ss.getSheetByName("QUERY").getDataRange().getValues());
}

/**
 * @return "_"区切りのタイムスタンプ
 */
function getTimeStamp() {
  const now = new Date();
  const year = now.getFullYear().toString();
  const month = (now.getMonth() + 1).toString().padStart(2, "0");
  const date = now.getDate().toString().padStart(2, "0");
  const hour = now.getHours().toString().padStart(2, "0");
  const minute = now.getMinutes().toString().padStart(2, "0");
  return `${year}_${month}_${date}_${hour}_${minute}`;
}

const searchCompensationDaysOff = (sheet) => {
  // 隠し列から振替元と振替先を取得
  const searchTarget = sheet.getRange(10, 17, 31, 2).getValues();
  const origin = [];
  const destination = [];
  for (let t of searchTarget) {
    if (t[0] !== '' & t[1] !== '') {
      origin.push([Utilities.formatDate(t[0], "JST", "yyyy年MM月dd日")]);
      destination.push([Utilities.formatDate(t[1], "JST", "yyyy年MM月dd日")]);
    }
  }

  const result = {
    createFlag: origin.length > 0 && destination.length > 0,
    origin,
    destination
  };
  return result;
}

/**
 * @param {String} SpreadsheetID
 * @param {String} シート名
 * @param {String} GoogleDriveのフォルダID
 */
const createPDF = (ssid, sheetName, folderId) => {
  const spreadSheet = SpreadsheetApp.openById(ssid);
  const targetSheet = spreadSheet.getSheetByName(sheetName);
  const token = ScriptApp.getOAuthToken();
  const opts = {
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "true",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "false",  // シート名をPDF上部に表示するか
    printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false",  // 固定行の表示有無
    scale: 4,
    gid:          targetSheet.getSheetId()   // シートIDを指定
  };

  const url = "https://docs.google.com/spreadsheets/d/SSID/".replace("SSID", spreadSheet.getId()) + "export?";
  
  const urlExt = [];
  for (let option in opts) {
    urlExt.push(option + "=" + opts[option]);
  }
  const options = urlExt.join("&");

  const response = UrlFetchApp.fetch(url + options, {headers: {'Authorization': 'Bearer ' +  token}});
  const timeStamp = getTimeStamp();
  const fileName = spreadSheet.getName() + timeStamp;

  return DriveApp.getFolderById(folderId).createFile(response).setName(fileName + ".pdf");
}

const updateAppForm = (result) => {
  // 更新対象を取得
  const appForm = SpreadsheetApp.openById(APPFORM_ID).getSheetByName("振替休日申請書");

  // 更新内容を取得
  const rowsNum = result.origin.length;
  for (let i = 0; i < 5 - rowsNum; i++) {
    result.origin.push(['']);
    result.destination.push(['']);
  }

  // 更新処理
  appForm.getRange(11, 4, 5, 1).setValues(result.origin);
  appForm.getRange(11, 14, 5, 1).setValues(result.destination);
  appForm.getRange(4, 15).setValue("TODO");

  // PDF出力
  createPDF(APPFORM_ID, APPFORMSHEET_NAME, FOLDER_ID);
}

/**
 * エントリーポイント
 */
const main = () => {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet  = ss.getActiveSheet();
  const confirmation = ui.alert(
    '確認',
    `シート：${targetSheet.getName()}　の振休申請を作成しますか？`,
    ui.ButtonSet.YES_NO_CANCEL
  );
  if (confirmation === ui.Button.NO || confirmation === ui.Button.CANCEL) {
    return;
  }

  const result = searchCompensationDaysOff(targetSheet);
  if (result.createFlag) {
    updateAppForm(result);
  } else {
    ui.alert("振替休日を取得していないため、処理を終了しました。");
    return;
  }
}