// ============================================================
// 環境2課 中期経営計画ダッシュボード — データ保存スクリプト v2
// 更新内容:
//   ⑥ パスワードハッシュをスプレッドシートで管理
//   ⑧ 保存のたびにバックアップ履歴を記録（最新30件）
// ============================================================

const SHEET_JSON     = 'ダッシュボードデータ';
const SHEET_OT       = '残業実績';
const SHEET_LEAVE    = '有休取得';
const SHEET_BACKUP   = 'バックアップ履歴';
const SHEET_SETTINGS = '設定';
const SHEET_SALES    = '売上実績';
const MONTHS = ['4月','5月','6月','7月','8月','9月','10月','11月','12月','1月','2月','3月'];
const MAX_BACKUPS = 30;
// デフォルトパスワード: 3150 の SHA-256ハッシュ
const DEFAULT_PIN_HASH = '4d364fbb3786fc31157cc1e2a2671aac0e36348ae9d6b4ba4459ee883c240fe8';

/* ─── GET: データ取得 / PIN照合 ─── */
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'verify_pin') {
    return verifyPin_(e.parameter.hash || '');
  }
  if (action === 'change_pin') {
    return changePin_(e.parameter.oldHash || '', e.parameter.newHash || '');
  }

  try {
    const sheet = getOrCreate_(SHEET_JSON);
    const json  = sheet.getRange('A1').getValue();
    return ok_(json || '{}');
  } catch (err) {
    return ok_(JSON.stringify({ error: err.message }));
  }
}

/* ─── POST: データ保存 ─── */
function doPost(e) {
  try {
    const json = e.postData.contents;
    const data = JSON.parse(json);

    const sheet = getOrCreate_(SHEET_JSON);
    sheet.getRange('A1').setValue(json);
    sheet.getRange('B1').setValue(
      Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
    );
    sheet.getRange('C1').setValue('最終保存日時');

    saveBackup_(json);
    writeOvertimeSheet_(data);
    writeLeaveSheet_(data);
    writeSalesSheet_(data);
    writeSalesTargetSheet_(data);

    return ok_(JSON.stringify({ status: 'saved' }));
  } catch (err) {
    return ok_(JSON.stringify({ error: err.message }));
  }
}

/* ─── ⑥ PIN照合 ─── */
function verifyPin_(inputHash) {
  try {
    const stored = getStoredPinHash_();
    return ok_(JSON.stringify({ valid: stored === inputHash }));
  } catch (err) {
    return ok_(JSON.stringify({ valid: false, error: err.message }));
  }
}

/* ─── ⑥ PIN変更 ─── */
function changePin_(oldHash, newHash) {
  try {
    if (getStoredPinHash_() !== oldHash)
      return ok_(JSON.stringify({ success: false, error: '現在のパスワードが違います' }));
    if (!newHash || newHash.length !== 64)
      return ok_(JSON.stringify({ success: false, error: '新しいパスワードが無効です' }));
    const s = getOrCreate_(SHEET_SETTINGS);
    s.getRange('A1').setValue(newHash);
    s.getRange('B1').setValue('パスワードハッシュ（SHA-256）');
    s.getRange('A2').setValue(Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'));
    s.getRange('B2').setValue('最終変更日時');
    return ok_(JSON.stringify({ success: true }));
  } catch (err) {
    return ok_(JSON.stringify({ success: false, error: err.message }));
  }
}

/* ─── ⑥ 保存済みPINハッシュを取得 ─── */
function getStoredPinHash_() {
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_SETTINGS);
  if (!s) return DEFAULT_PIN_HASH;
  return s.getRange('A1').getValue() || DEFAULT_PIN_HASH;
}

/* ─── ⑧ バックアップ履歴（最新30件保持） ─── */
function saveBackup_(json) {
  try {
    const sheet = getOrCreate_(SHEET_BACKUP);
    if (sheet.getRange('A1').getValue() !== '保存日時')
      sheet.getRange('A1:B1').setValues([['保存日時','データ（JSON）']]).setFontWeight('bold');
    const ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    sheet.insertRowAfter(1);
    sheet.getRange(2,1,1,2).setValues([[ts, json]]);
    const last = sheet.getLastRow();
    if (last > MAX_BACKUPS + 1)
      sheet.deleteRows(MAX_BACKUPS + 2, last - MAX_BACKUPS - 1);
  } catch(e) { /* バックアップ失敗はメイン処理を止めない */ }
}

/* ─── 残業実績シート ─── */
function writeOvertimeSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_OT) || ss.insertSheet(SHEET_OT);
  sheet.clearContents();
  const OT25  = [239,148,245.5,234.5,229,283.75,276.5,351.25,366.5,289,297.7,401];
  const OT26T = OT25.map(v => Math.round(v*.50*100)/100);
  const OT27T = OT25.map(v => Math.round(v*.40*100)/100);
  const OT28T = OT25.map(v => Math.round(v*.30*100)/100);
  const header = ['月','2025年度実績(h)','2026実績(h)','2027実績(h)','2028実績(h)',
    '2026目標(h)','2027目標(h)','2028目標(h)'];
  const rows = MONTHS.map((m,i) => [m, OT25[i],
    data.actualData && data.actualData['2026'] ? (data.actualData['2026'][i] ?? '') : '',
    data.actualData && data.actualData['2027'] ? (data.actualData['2027'][i] ?? '') : '',
    data.actualData && data.actualData['2028'] ? (data.actualData['2028'][i] ?? '') : '',
    OT26T[i], OT27T[i], OT28T[i]]);
  const toSum = a => a.filter(v=>v!==''&&v!=null).reduce((s,v)=>s+v,0);
  rows.push(['合計',toSum(OT25),
    toSum((data.actualData||{})['2026']||[]),toSum((data.actualData||{})['2027']||[]),
    toSum((data.actualData||{})['2028']||[]),toSum(OT26T),toSum(OT27T),toSum(OT28T)]);
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ─── 有休取得シート ─── */
function writeLeaveSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_LEAVE) || ss.insertSheet(SHEET_LEAVE);
  sheet.clearContents();
  const header = ['氏名','グループ','付与日数','取得合計(日)','取得率(%)','残日数',
    ...MONTHS.map(m=>m+'取得')];
  const members = data.members || [];
  const rows = members.map(m => {
    const tm    = m.takenMonths || new Array(12).fill(0);
    const taken = tm.reduce((a,v)=>a+(v||0),0);
    const rate  = m.days>0?Math.round(taken/m.days*100):0;
    return [m.name,m.group||'',m.days,taken,rate,Math.max(0,m.days-taken),...tm];
  });
  if (!rows.length) return;
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  for (let r=2; r<=rows.length+1; r++) {
    const rate = sheet.getRange(r,5).getValue();
    sheet.getRange(r,5).setBackground(rate>=80?'#dcfce7':rate>=50?'#fef3c7':'#fee2e2');
  }
  sheet.setFrozenRows(1);
}

/* ─── 売上実績シート ─── */
function writeSalesSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_SALES) || ss.insertSheet(SHEET_SALES);
  sheet.clearContents();
  const header = ['年度', '月', 'ユニット2(万円)', 'ユニット3(万円)', '合算(万円)'];
  const rows = [];
  const sd = data.salesData || {};
  ['2025','2026','2027','2028'].forEach(yr => {
    const u2 = sd[yr]?.unit2 || new Array(12).fill(null);
    const u3 = sd[yr]?.unit3 || new Array(12).fill(null);
    MONTHS.forEach((m, i) => {
      const v2 = u2[i] ?? '', v3 = u3[i] ?? '';
      const tot = (u2[i] != null && u3[i] != null) ? u2[i] + u3[i] : '';
      rows.push([yr + '年度', m, v2, v3, tot]);
    });
  });
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  if(rows.length) sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ─── 売上目標シート ─── */
function writeSalesTargetSheet_(data) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName('売上目標') || ss.insertSheet('売上目標');
  sheet.clearContents();
  const header = ['年度','ユニット2目標(万円)','ユニット3目標(万円)','合算目標(万円)'];
  const st = data.salesTargets || {};
  const rows = ['2025','2026','2027','2028'].map(yr => [
    yr+'年度',
    st[yr]?.unit2 ?? '',
    st[yr]?.unit3 ?? '',
    st[yr]?.total ?? ''
  ]);
  sheet.getRange(1,1,1,header.length).setValues([header]).setFontWeight('bold').setBackground('#f1f5f9');
  sheet.getRange(2,1,rows.length,header.length).setValues(rows);
  sheet.setFrozenRows(1);
}

/* ─── ヘルパー ─── */
function getOrCreate_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function ok_(text) {
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}
