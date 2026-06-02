// ============================================================
// 活動日報 参加者データ記録スクリプト
// Google Apps Script に貼り付けて「ウェブアプリとして導入」する
// ============================================================

// ▼ スプレッドシートIDをここに入力（URLの /d/〇〇〇/ の部分）
const SPREADSHEET_ID = 'ここにスプレッドシートIDを入力';

// ▼ シート名（変更する場合はここを書き換え）
const SHEET_NAME = '参加者記録';

// ============================================================
// POST リクエストを受け取ってスプレッドシートに書き込む
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet   = ss.getSheetByName(SHEET_NAME);

    // シートがなければ新規作成してヘッダーを追加
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        '記録日時',
        '活動日',
        '曜日',
        '開始時間',
        '終了時間',
        '記録者',
        '場所',
        '天候',
        // こども
        '未就学児',
        '小学生',
        '中学生',
        '高校生',
        'こども合計',
        // 大人
        '学校関係',
        '地域住民',
        '大学生',
        'その他大人',
        '大人合計',
        // 総計
        '合計',
      ]);

      // ヘッダー行を太字・背景色設定
      const headerRange = sheet.getRange(1, 1, 1, 20);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#2563a8');
      headerRange.setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }

    // 書き込む行データ
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
      Number(data.preschool  || 0),
      Number(data.elementary || 0),
      Number(data.junior     || 0),
      Number(data.senior     || 0),
      Number(data.kidsTotal  || 0),
      Number(data.school     || 0),
      Number(data.community  || 0),
      Number(data.univ       || 0),
      Number(data.otherAdult || 0),
      Number(data.adultsTotal|| 0),
      Number(data.grandTotal || 0),
    ];

    sheet.appendRow(row);

    // 最終行のスタイル（交互カラー）
    const lastRow = sheet.getLastRow();
    if (lastRow % 2 === 0) {
      sheet.getRange(lastRow, 1, 1, 20).setBackground('#f0f4ff');
    }

    // 列幅を自動調整（初回のみ重いので最初の数回だけ実行）
    if (lastRow <= 5) {
      sheet.autoResizeColumns(1, 20);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok', row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET リクエスト（テスト用・ブラウザで直接開いて確認できる）
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '活動日報 GAS is running' }))
    .setMimeType(ContentService.MimeType.JSON);
}
