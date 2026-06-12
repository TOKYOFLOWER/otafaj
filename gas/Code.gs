/**
 * 花ロス管理 API (Google Apps Script)
 *
 * ■ セットアップ
 * 1. このコードをGASプロジェクトに貼り付け
 * 2. スプレッドシートにバインドされていない(スタンドアロン)場合は、
 *    プロジェクトの設定 → スクリプト プロパティ に以下を追加:
 *      SPREADSHEET_ID = 1myDnj2NsEQjRi7WzWAda1QeYk74IFQmLbvyS3BO25Ms
 *    (バインド型なら設定不要。getActive() が使われます)
 * 3. デプロイ → 新しいデプロイ → ウェブアプリ
 *    実行ユーザー: 自分 / アクセス: 全員 → /exec のURLを index.html 側の config.js へ
 *
 * ■ API
 *   ?action=data&days=15&source=both   … 直近N日の仕入データ一括取得(ドロップダウン用)
 *   ?action=record&...                 … ロス記録(「ロス管理」シートへ追記)
 *   ?action=summary&from&to&groupBy    … 廃棄本数・金額の集計
 */

// ===== シート構成(市場ごとに列が異なる) =====
//
// FAJ (タブ名: 2606 など YYMM)
//   A:seriDate B:khKubun C:khKubunName D:nykBan E:hanbaiName F:hanbaiKubunname
//   G:sanchiName H:seisanName I:shohinName J:kikakuName K:tanka L:kingaku
//   ※数量は 金額÷単価 で逆算
//
// 大田花き (タブ名: OTA2606 など)
//   A:日付 B:No C:登録No D:品目・品種 E:等階級 F:入数 G:口数 H:本数
//   I:単価 J:金額 K:競台 L:競人 M:産地名/記事 N:買参人
//   ※産地名/記事は「産地 出荷者名」形式 → 最初の空白で分割

const SOURCES = {
  FAJ: {
    label: 'FAJ',
    tabName: function (yymm) { return yymm; },
    parse: function (r) {
      const tanka = num(r[10]);            // K
      const kingaku = num(r[11]);          // L
      if (!tanka && !kingaku) return null;
      const name = String(r[8] || '').trim();      // I shohinName
      if (!name || name === 'shohinName') return null;
      const split = splitFajName(name);
      return {
        date: parseDate(r[0]),
        sanchi: String(r[6] || '').trim(),         // G sanchiName
        shukka: String(r[7] || '').trim(),         // H seisanName
        hinmoku: split.hinmoku,
        hinshu: split.hinshu,
        kikaku: String(r[9] || '').trim(),         // J kikakuName
        tanka: tanka,
        suryo: tanka > 0 ? Math.round(kingaku / tanka) : 0,
        kingaku: kingaku
      };
    }
  },
  OTA: {
    label: '大田花き',
    tabName: function (yymm) { return 'OTA' + yymm; },
    parse: function (r) {
      const tanka = num(r[8]);             // I 単価
      const kingaku = num(r[9]);           // J 金額
      if (!tanka && !kingaku) return null;
      const name = String(r[3] || '').trim();      // D 品目・品種
      if (!name || name === '品目・品種') return null;
      const split = splitOtaName(name);
      const sanchiKiji = String(r[12] || '').trim(); // M 産地名/記事
      const sp = sanchiKiji.search(/[\s\u3000]/);
      return {
        date: parseDate(r[0]),
        sanchi: sp > 0 ? sanchiKiji.slice(0, sp) : sanchiKiji,
        shukka: sp > 0 ? sanchiKiji.slice(sp + 1).trim() : '',
        hinmoku: split.hinmoku,
        hinshu: split.hinshu,
        kikaku: String(r[4] || '').trim(),          // E 等階級
        tanka: tanka,
        suryo: num(r[7]),                           // H 本数
        kingaku: kingaku
      };
    }
  }
};

// 大田花き D列「品目■■品種」を分割(■■は色表示プレースホルダ)
function splitOtaName(name) {
  const i = name.indexOf('■■');
  if (i >= 0) {
    return { hinmoku: name.slice(0, i).trim(), hinshu: name.slice(i + 2).trim() };
  }
  return { hinmoku: name, hinshu: '' };
}

// FAJ shohinName「品目・色“品種名”」を分割
// 例: バラ・赤“サムライ０８” → 品目=バラ・赤  品種=サムライ０８
//     ダリア“艶舞”         → 品目=ダリア    品種=艶舞
//     バラ・赤             → 品目=バラ・赤  品種=""
//     カーネーション       → 品目=カーネーション 品種=""
//     キク（デコラ・タイプ）→ 品目=キク（デコラ・タイプ） 品種=""
function splitFajName(name) {
  // “...” の引用符を品種名の目印として使う
  // 引用符の前が品目（色を含む）、引用符の中身が品種
  const q = name.search(/[“”"]/);
  if (q > 0) {
    return {
      hinmoku: name.slice(0, q).trim(),
      hinshu:  name.slice(q + 1).replace(/[“””]\s*$/, '').trim()
    };
  }
  return { hinmoku: name.trim(), hinshu: '' };
}

const LOSS_SHEET_NAME = 'ロス管理';
const LOSS_HEADERS = ['記録日時', '廃棄日', '市場', '仕入タブ', '産地', '出荷者', '品目', '品種', '規格', '単価', '廃棄本数', '廃棄金額', 'メモ'];
const MAX_ROWS = 3000;

// ===== エントリポイント =====
function doGet(e) {
  let result;
  try {
    const p = (e && e.parameter) || {};
    switch (p.action) {
      case 'data':    result = actionData(p);    break;
      case 'record':  result = actionRecord(p);  break;
      case 'summary': result = actionSummary(p); break;
      default:        result = { ok: false, error: 'unknown action: ' + p.action };
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSS() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

// ===== 直近N日の仕入データ一括取得 =====
function actionData(p) {
  const days = Math.min(Number(p.days) || 15, 62);
  const source = String(p.source || 'both');
  const ss = getSS();

  const to = new Date();
  const from = new Date(to.getTime() - days * 24 * 3600 * 1000);
  from.setHours(0, 0, 0, 0);

  const yymms = yymmRange(from, to);
  const keys = (source === 'both') ? Object.keys(SOURCES) : [source];

  const rows = [];
  keys.forEach(function (key) {
    const src = SOURCES[key];
    if (!src) return;
    yymms.forEach(function (yymm) {
      const sh = ss.getSheetByName(src.tabName(yymm));
      if (!sh) return;
      const data = sh.getDataRange().getValues();
      for (let i = 0; i < data.length; i++) {
        const item = src.parse(data[i]);
        if (!item || !item.date) continue;
        if (item.date < from || item.date > to) continue;
        rows.push({
          source: key,
          sourceLabel: src.label,
          tab: src.tabName(yymm),
          date: Utilities.formatDate(item.date, 'Asia/Tokyo', 'yyyy-MM-dd'),
          sanchi: item.sanchi, shukka: item.shukka,
          hinmoku: item.hinmoku, hinshu: item.hinshu, kikaku: item.kikaku,
          tanka: item.tanka, suryo: item.suryo, kingaku: item.kingaku
        });
        if (rows.length >= MAX_ROWS) break;
      }
    });
  });

  rows.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  return {
    ok: true,
    from: Utilities.formatDate(from, 'Asia/Tokyo', 'yyyy-MM-dd'),
    to: Utilities.formatDate(to, 'Asia/Tokyo', 'yyyy-MM-dd'),
    count: rows.length,
    rows: rows
  };
}

// ===== ロス記録 =====
function actionRecord(p) {
  const honsu = Number(p.honsu);
  const tanka = Number(p.tanka);
  if (!honsu || honsu <= 0) return { ok: false, error: '廃棄本数が不正です' };
  if (isNaN(tanka)) return { ok: false, error: '単価が不正です' };

  const sh = getLossSheet();
  const kingaku = Math.round(tanka * honsu);
  const haikiDate = p.date ? p.date : Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

  sh.appendRow([
    new Date(),
    haikiDate,
    String(p.sourceLabel || p.source || ''),
    String(p.tab || ''),
    String(p.sanchi || ''),
    String(p.shukka || ''),
    String(p.hinmoku || ''),
    String(p.hinshu || ''),
    String(p.kikaku || ''),
    tanka,
    honsu,
    kingaku,
    String(p.memo || '')
  ]);
  return { ok: true, kingaku: kingaku, haikiDate: haikiDate };
}

// ===== 集計 =====
function actionSummary(p) {
  const sh = getLossSheet();
  const data = sh.getDataRange().getValues();
  const from = p.from ? new Date(p.from + 'T00:00:00+09:00') : null;
  const to   = p.to   ? new Date(p.to   + 'T23:59:59+09:00') : null;
  const groupBy = p.groupBy || 'hinmoku';   // hinmoku / sanchi / month

  let totalHonsu = 0, totalKingaku = 0, count = 0;
  const groups = {};

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const d = parseDate(r[1]);
    if (!d) continue;
    if (from && d < from) continue;
    if (to && d > to) continue;

    const honsu = Number(r[10]) || 0;
    const kingaku = Number(r[11]) || 0;
    totalHonsu += honsu;
    totalKingaku += kingaku;
    count++;

    let key;
    if (groupBy === 'sanchi') key = String(r[4] || '不明');
    else if (groupBy === 'month') key = Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM');
    else key = String(r[6] || '不明');

    if (!groups[key]) groups[key] = { key: key, honsu: 0, kingaku: 0, count: 0 };
    groups[key].honsu += honsu;
    groups[key].kingaku += kingaku;
    groups[key].count++;
  }

  const list = Object.keys(groups).map(function (k) { return groups[k]; })
    .sort(function (a, b) { return b.kingaku - a.kingaku; });

  return { ok: true, total: { honsu: totalHonsu, kingaku: totalKingaku, count: count }, groups: list };
}

// ===== ユーティリティ =====
function getLossSheet() {
  const ss = getSS();
  let sh = ss.getSheetByName(LOSS_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(LOSS_SHEET_NAME);
    sh.appendRow(LOSS_HEADERS);
    sh.setFrozenRows(1);
    return sh;
  }
  // 旧フォーマット(品種列なし)なら品目の後ろに品種列を自動挿入
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  if (headers.indexOf('品種') === -1) {
    const pos = headers.indexOf('品目');
    if (pos >= 0) {
      sh.insertColumnAfter(pos + 1);
      sh.getRange(1, pos + 2).setValue('品種');
    }
  }
  return sh;
}

// from〜to をカバーする YYMM のリスト (例: ['2605','2606'])
function yymmRange(from, to) {
  const list = [];
  let d = new Date(from.getFullYear(), from.getMonth(), 1);
  while (d <= to) {
    const yy = ('0' + (d.getFullYear() % 100)).slice(-2);
    const mm = ('0' + (d.getMonth() + 1)).slice(-2);
    list.push(yy + mm);
    d = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  }
  return list;
}

function num(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v || '').replace(/[,，\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function parseDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === 'string' && v) {
    const d = new Date(v.replace(/\//g, '-'));
    if (!isNaN(d)) return d;
    const d2 = new Date(v);
    if (!isNaN(d2)) return d2;
  }
  return null;
}
