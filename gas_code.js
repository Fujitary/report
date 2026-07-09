// ============================================================
// 活動日報 参加者データ記録スクリプト
// Google Apps Script に貼り付けて「ウェブアプリとして導入」する
// ============================================================

// ▼ スプレッドシートIDをここに入力（URLの /d/〇〇〇/ の部分）
const SPREADSHEET_ID = 'ここにスプレッドシートIDを入力';

// ▼ シート名
const SHEET_NAME = '参加者記録';

// ============================================================
// CORS ヘッダーを付けてレスポンスを返す共通関数
// ============================================================
function corsResponse(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// OPTIONS プリフライト対応（GETで代用）
// ============================================================
function doGet(e) {
  // action=ping でアプリの死活確認に使える
  return corsResponse({ status: 'ok', message: '活動日報 GAS is running' });
}

// ============================================================
// POST：データをスプレッドシートに書き込む
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // シートがなければ新規作成してヘッダーを追加
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        '記録日時', '活動日', '曜日', '開始時間', '終了時間',
        '記録者', '場所', '天候',
        '未就学児', '小学生', '中学生', '高校生', 'こども合計',
        '学校関係', '地域住民', '大学生', 'その他大人', '大人合計',
        '合計'
      ];
      sheet.appendRow(headers);

      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#2563a8');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    const now = new Date();
    const row = [
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
      data.date        || '',
      data.day         || '',
      data.timeStart   || '',
      data.timeEnd     || '',
      data.recorder    || '',
      data.place       || '',
      data.weather     || '',
      Number(data.preschool   || 0),
      Number(data.elementary  || 0),
      Number(data.junior      || 0),
      Number(data.senior      || 0),
      Number(data.kidsTotal   || 0),
      Number(data.school      || 0),
      Number(data.community   || 0),
      Number(data.univ        || 0),
      Number(data.otherAdult  || 0),
      Number(data.adultsTotal || 0),
      Number(data.grandTotal  || 0),
    ];

    sheet.appendRow(row);

    // 偶数行に薄い背景色
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, row.length).setBackground('#f0f4ff');
    }

    return corsResponse({ status: 'ok', row: lastRow });

  } catch (err) {
    return corsResponse({ status: 'error', message: err.message });
  }
}
