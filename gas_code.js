// ============================================================
// 活動日報 クラウド同期スクリプト（Google Apps Script）
// スプレッドシート ＋ Googleドライブ（写真）
//
// 【設定手順】
// 1. スプレッドシートを作成し、URLの /d/〇〇〇/ をコピー
// 2. Googleドライブに写真保存用フォルダを作成し、URLの末尾IDをコピー
// 3. 下の2つの定数に貼り付ける
// 4. デプロイ → 新しいデプロイ → ウェブアプリ
//    実行するユーザー：自分  /  アクセスできるユーザー：全員
// ============================================================

const SPREADSHEET_ID = 'ここにスプレッドシートIDを入力';
const PHOTO_FOLDER_ID = 'ここに写真保存用フォルダIDを入力';
const SHEET_NAME = '活動日報';

// 列定義（この順序でシートに書き込まれる）
const COLUMNS = [
  'ID', '更新日時', '活動日', '曜日', '開始時間', '終了時間',
  '記録者', '場所', '天候',
  '未就学児', '小学生', '中学生', '高校生', 'こども合計',
  '学校関係', '地域住民', '大学生', 'その他大人', '大人合計', '合計',
  '実施内容', '所感・気づき・課題',
  'キャプション1', 'キャプション2', '写真1_ID', '写真2_ID', '写真リンク'
];

// ============================================================
// エントリポイント
// ============================================================
function doGet(e) {
  return json({ status: 'ok', message: '活動日報 GAS 稼働中' });
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const action = req.action || 'save';

    switch (action) {
      case 'list':   return json(listReports());
      case 'get':    return json(getReport(req.id));
      case 'save':   return json(saveReport(req.data));
      case 'delete': return json(deleteReport(req.id));
      default:       return json({ status: 'error', message: '不明なaction: ' + action });
    }
  } catch (err) {
    return json({ status: 'error', message: err.message });
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// シート取得（なければ作成）
// ============================================================
function getSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(COLUMNS);
    const h = sheet.getRange(1, 1, 1, COLUMNS.length);
    h.setFontWeight('bold').setBackground('#2563a8').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(21, 300);
    sheet.setColumnWidth(22, 300);
  }
  return sheet;
}

// 行データ → オブジェクト
function rowToObj(row) {
  const o = {};
  COLUMNS.forEach((c, i) => { o[c] = row[i]; });
  return {
    id:         o['ID'],
    updatedAt:  o['更新日時'],
    date:       formatDateCell(o['活動日']),
    day:        o['曜日'],
    timeStart:  formatTimeCell(o['開始時間']),
    timeEnd:    formatTimeCell(o['終了時間']),
    recorder:   o['記録者'],
    place:      o['場所'],
    weather:    o['天候'],
    preschool:  Number(o['未就学児']) || 0,
    elementary: Number(o['小学生'])   || 0,
    junior:     Number(o['中学生'])   || 0,
    senior:     Number(o['高校生'])   || 0,
    school:     Number(o['学校関係']) || 0,
    community:  Number(o['地域住民']) || 0,
    univ:       Number(o['大学生'])   || 0,
    other:      Number(o['その他大人'])|| 0,
    total:      Number(o['合計'])     || 0,
    content:    o['実施内容'] || '',
    impression: o['所感・気づき・課題'] || '',
    cap0:       o['キャプション1'] || '',
    cap1:       o['キャプション2'] || '',
    photo0Id:   o['写真1_ID'] || '',
    photo1Id:   o['写真2_ID'] || '',
  };
}

// 日付セルを yyyy-MM-dd に正規化（スプシで手編集されても壊れない）
function formatDateCell(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  const s = String(v).trim().replace(/\//g, '-');
  return s;
}

// 時刻セルを HH:mm に正規化
function formatTimeCell(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'HH:mm');
  return String(v).trim();
}

// ============================================================
// 一覧取得（写真本体は含めず軽量に）
// ============================================================
function listReports() {
  const sheet = getSheet();
  const last = sheet.getLastRow();
  if (last < 2) return { status: 'ok', items: [] };

  const rows = sheet.getRange(2, 1, last - 1, COLUMNS.length).getValues();
  const items = rows
    .filter(r => r[0]) // IDがある行のみ
    .map(rowToObj)
    .map(o => ({
      id: o.id, date: o.date, day: o.day,
      recorder: o.recorder, place: o.place,
      total: o.total,
      content: String(o.content).slice(0, 80),
      hasPhoto: !!(o.photo0Id || o.photo1Id),
      updatedAt: o.updatedAt instanceof Date
        ? Utilities.formatDate(o.updatedAt, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm')
        : String(o.updatedAt || ''),
    }))
    .sort((a, b) => String(b.date).localeCompare(String(a.date)));

  return { status: 'ok', items: items };
}

// ============================================================
// 1件取得（写真をbase64で返す）
// ============================================================
function getReport(id) {
  const sheet = getSheet();
  const idx = findRowIndexById(sheet, id);
  if (idx < 0) return { status: 'error', message: '該当データがありません' };

  const row = sheet.getRange(idx, 1, 1, COLUMNS.length).getValues()[0];
  const o = rowToObj(row);

  o.photo0 = o.photo0Id ? driveFileToDataUrl(o.photo0Id) : null;
  o.photo1 = o.photo1Id ? driveFileToDataUrl(o.photo1Id) : null;

  return { status: 'ok', data: o };
}

function driveFileToDataUrl(fileId) {
  try {
    const blob = DriveApp.getFileById(fileId).getBlob();
    return 'data:' + blob.getContentType() + ';base64,' +
           Utilities.base64Encode(blob.getBytes());
  } catch (e) {
    return null;
  }
}

// ============================================================
// 保存（IDがあれば上書き、なければ新規）
// ============================================================
function saveReport(d) {
  const sheet = getSheet();
  const isNew = !d.id;
  const id = d.id || ('rep_' + new Date().getTime());

  // 既存行があれば古い写真IDを取得
  let oldPhoto0 = '', oldPhoto1 = '';
  let rowIdx = -1;
  if (!isNew) {
    rowIdx = findRowIndexById(sheet, id);
    if (rowIdx > 0) {
      const old = sheet.getRange(rowIdx, 1, 1, COLUMNS.length).getValues()[0];
      const oo = rowToObj(old);
      oldPhoto0 = oo.photo0Id;
      oldPhoto1 = oo.photo1Id;
    }
  }

  // 写真の処理
  // d.photo0 が dataURL → 新規アップロード（古いのは削除）
  // d.photo0 === 'KEEP'  → 既存を維持
  // d.photo0 が null     → 削除
  const p0 = handlePhoto(d.photo0, oldPhoto0, id + '_1');
  const p1 = handlePhoto(d.photo1, oldPhoto1, id + '_2');

  const kids   = num(d.preschool) + num(d.elementary) + num(d.junior) + num(d.senior);
  const adults = num(d.school) + num(d.community) + num(d.univ) + num(d.other);

  const photoLinks = [p0, p1].filter(Boolean)
    .map(fid => 'https://drive.google.com/file/d/' + fid + '/view').join('\n');

  const row = [
    id,
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),
    d.date || '', d.day || '',
    d.timeStart || '', d.timeEnd || '',
    d.recorder || '', d.place || '', d.weather || '',
    num(d.preschool), num(d.elementary), num(d.junior), num(d.senior), kids,
    num(d.school), num(d.community), num(d.univ), num(d.other), adults,
    kids + adults,
    d.content || '', d.impression || '',
    d.cap0 || '', d.cap1 || '',
    p0 || '', p1 || '',
    photoLinks
  ];

  if (rowIdx > 0) {
    // 上書き
    sheet.getRange(rowIdx, 1, 1, row.length).setValues([row]);
  } else {
    // 新規追加
    sheet.appendRow(row);
    const lr = sheet.getLastRow();
    if (lr % 2 === 0) {
      sheet.getRange(lr, 1, 1, row.length).setBackground('#f0f4ff');
    }
  }

  return { status: 'ok', id: id, isNew: rowIdx < 0 };
}

function num(v) { return Number(v) || 0; }

// 写真処理：新規アップロード / 維持 / 削除
function handlePhoto(value, oldFileId, baseName) {
  // 変更なし
  if (value === 'KEEP') return oldFileId;

  // 削除された
  if (!value) {
    if (oldFileId) trashFile(oldFileId);
    return '';
  }

  // dataURL → アップロード
  if (typeof value === 'string' && value.indexOf('data:') === 0) {
    if (oldFileId) trashFile(oldFileId);
    try {
      const parts = value.split(',');
      const mime = parts[0].match(/data:([^;]+)/)[1];
      const bytes = Utilities.base64Decode(parts[1]);
      const ext = mime.indexOf('png') >= 0 ? 'png' : 'jpg';
      const blob = Utilities.newBlob(bytes, mime, baseName + '.' + ext);
      const folder = DriveApp.getFolderById(PHOTO_FOLDER_ID);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return file.getId();
    } catch (e) {
      return oldFileId || '';
    }
  }

  return oldFileId || '';
}

function trashFile(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
}

// ============================================================
// 削除
// ============================================================
function deleteReport(id) {
  const sheet = getSheet();
  const idx = findRowIndexById(sheet, id);
  if (idx < 0) return { status: 'error', message: '該当データがありません' };

  // 写真も削除
  const row = sheet.getRange(idx, 1, 1, COLUMNS.length).getValues()[0];
  const o = rowToObj(row);
  if (o.photo0Id) trashFile(o.photo0Id);
  if (o.photo1Id) trashFile(o.photo1Id);

  sheet.deleteRow(idx);
  return { status: 'ok' };
}

// ============================================================
// ID から行番号を検索
// ============================================================
function findRowIndexById(sheet, id) {
  if (!id) return -1;
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}
