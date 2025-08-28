/**
 * 叫藥自動化（Apps Script 版本）
 * 從【結果】分頁抓到每家廠商的藥品資訊（可是一列多個品名的自由文字），
 * 再參照【常見量對照】分頁（商品｜規格｜常見叫藥數量｜廠商），
 * 輸出「{廠商}想訂{商品 規格? 常見叫藥數量}、…」到【訂單文字】分頁。
 *
 * ⚙️ 參數：
 *   RESULT_SHEET  -> 來源分頁（預設：'結果'）
 *   MAP_SHEET     -> 常見量對照分頁（預設：'常見量對照'）
 *   OUTPUT_SHEET  -> 輸出分頁（預設：'訂單文字'）
 *   INCLUDE_SPEC  -> 是否在輸出文字中包含規格（true/false）
 *   DEFAULT_UNIT  -> 當常見叫藥數量沒有單位時補上的預設單位（預設：'盒'）
 */

const RESULT_SHEET = '結果';
const MAP_SHEET = '常見量對照'; // 欄位必須包含：商品｜規格（可空）｜常見叫藥數量｜廠商（可空）
const OUTPUT_SHEET = '訂單文字';
const INCLUDE_SPEC = false;     // 若想只輸出「商品 常見叫藥數量」→ 改為 false
const DEFAULT_UNIT = '盒';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('叫藥自動化')
    .addItem('生成訂單文字', 'generateVendorOrderLines')
    .addToUi();
}

/** 主要流程 */
function generateVendorOrderLines() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); // 官方建議用法存取活頁簿
  const resultSheet = ss.getSheetByName(RESULT_SHEET);
  const mapSheet = ss.getSheetByName(MAP_SHEET);

  if (!resultSheet) throw new Error(`找不到分頁：${RESULT_SHEET}`);
  if (!mapSheet) throw new Error(`找不到分頁：${MAP_SHEET}（請建立並放欄位：商品｜規格｜常見叫藥數量｜廠商）`);

  // 讀結果分頁與對照分頁
  const resultRows = readResultSheet_(resultSheet);
  const mapRows = readMapSheet_(mapSheet);

  // 依廠商彙整一段「大文字」，再以商品關鍵字去比對
  const vendorText = buildVendorBigText_(resultRows);
  const lines = makeLinesFromMap_(vendorText, mapRows);

  // 輸出到【訂單文字】
  writeOutput_(ss, lines);
}

/** 讀【結果】分頁：自動抓欄位（廠商、商品／品名／藥品名稱、規格、或「藥品資訊」） */
function readResultSheet_(sheet) {
  const values = sheet.getDataRange().getValues(); // 一次讀完整表，效能佳
  if (!values.length) return [];

  const headers = values[0].map(h => String(h).trim());
  const idx = colIndexFinder_(headers, {
    vendor: ['廠商', '供應商', '製造商', '廠牌'],
    name: ['商品名', '商品名稱', '品名', '藥品名稱', '藥品名', '名稱', '品項'],
    spec: ['規格', '含量', '劑量', '規格含量', '包裝', 'Strength'],
    info: ['藥品資訊', '藥品資訊(商品+規格)', '資訊']
  });

  // 將每列轉為 {vendor, name, spec, info, rawText}
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const vendor = normText_(pickCell_(row, headers, idx.vendor));
    const name = normText_(pickCell_(row, headers, idx.name));
    const spec = normText_(pickCell_(row, headers, idx.spec));
    const info = normText_(pickCell_(row, headers, idx.info));

    // rawText：供關鍵字搜尋（把安全庫存類備註移除）
    const rawText = [name, spec, info].filter(Boolean).join(' ').trim();
    if (!vendor && !rawText) continue;
    rows.push({ vendor, name, spec, info, rawText });
  }
  return rows;
}

/** 讀【常見量對照】分頁：商品｜規格（可空）｜常見叫藥數量｜廠商（可空） */
function readMapSheet_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error(`對照分頁 ${MAP_SHEET} 是空的`);

  const headers = values[0].map(h => String(h).trim());
  const need = ['商品', '常見叫藥數量']; // 規格、廠商可空
  need.forEach(k => { if (!headers.includes(k)) throw new Error(`${MAP_SHEET} 缺少欄位：${k}`); });

  const get = name => headers.indexOf(name);
  const iName = get('商品');
  const iSpec = headers.indexOf('規格');
  const iQty  = get('常見叫藥數量');
  const iVen  = headers.indexOf('廠商');

  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const v = values[r];
    const name = normText_(v[iName]);
    const spec = normText_(iSpec >= 0 ? v[iSpec] : '');
    const qty  = normText_(v[iQty]);
    const vendor = normText_(iVen >= 0 ? v[iVen] : '');
    if (!name) continue;

    rows.push({
      name,
      nameNorm: name.toLowerCase(),
      spec,
      specNorm: (spec || '').toLowerCase(),
      qty,
      vendor,
      vendorNorm: (vendor || '').toLowerCase()
    });
  }
  return rows;
}

/** 將結果列依廠商彙整成大段文字（用來做 substring 關鍵字比對） */
function buildVendorBigText_(rows) {
  const map = {};
  rows.forEach(({ vendor, rawText }) => {
    const key = vendor || ''; // 允許空廠商
    map[key] = (map[key] || '') + ' ' + (rawText || '');
  });
  // 正規化：去除安全庫存備註、壓小寫
  Object.keys(map).forEach(k => {
    map[k] = normText_(map[k]).toLowerCase();
  });
  return map; // { '美時': '…', '保瑞': '…', '': '…' }
}

/** 依每家廠商的大文字，從 mapRows 找出出現過的商品，產生訂單句子 */
function makeLinesFromMap_(vendorText, mapRows) {
  const lines = []; // [[vendor, "美時想訂Cimidona 30盒、Mesyrel 10盒"], …]
  const vendors = Object.keys(vendorText);

  vendors.forEach(vendorKey => {
    const big = vendorText[vendorKey]; // 此廠商在【結果】分頁所出現的所有文字
    // 限縮：若 mapRows 有廠商欄位，就先只看同廠商；否則全表
    const candidates = mapRows.filter(r => !r.vendorNorm || r.vendorNorm === vendorKey.toLowerCase());
    const items = [];
    const seen = new Set();

    candidates.forEach(r => {
      if (!r.nameNorm) return;
      if (!big.includes(r.nameNorm)) return; // 關鍵字出現才算此廠商本次要訂
      const qty = ensureQtyUnit_(r.qty);
      if (!qty) return; // 沒常見量就跳過
      const disp = INCLUDE_SPEC && r.spec ? `${r.name} ${r.spec} ${qty}` : `${r.name} ${qty}`;
      if (!seen.has(disp)) {
        seen.add(disp);
        items.push(disp);
      }
    });

    if (items.length) {
      const vendorDisp = vendorKey || '（未指定廠商）';
      lines.push([vendorDisp, `${vendorDisp}想訂` + items.join('、')]);
    }
  });

  return lines.sort((a, b) => a[0].localeCompare(b[0], 'zh-Hant'));
}

/** 寫出到【訂單文字】分頁 */
function writeOutput_(ss, lines) {
  const sheet = ss.getSheetByName(OUTPUT_SHEET) || ss.insertSheet(OUTPUT_SHEET);
  sheet.clear(); // 先清空
  const out = [['廠商', '訂單文字']].concat(lines);
  sheet.getRange(1, 1, out.length, out[0].length).setValues(out); // 一次寫入二維陣列
}

/* ----------------------- 小工具 ------------------------ */

/** 在標題列中找出對應欄位索引（用多個候選名） */
function colIndexFinder_(headers, dict) {
  const norm = s => String(s || '').replace(/\s+/g, '').toLowerCase();
  const findOne = keys => {
    const set = new Set(headers.map(h => norm(h)));
    for (const k of keys) {
      const key = norm(k);
      if (set.has(key)) return headers.findIndex(h => norm(h) === key);
    }
    // 模糊包含
    for (let i = 0; i < headers.length; i++) {
      const h = norm(headers[i]);
      if (keys.some(k => h.includes(norm(k)))) return i;
    }
    return -1;
  };
  return {
    vendor: findOne(dict.vendor || []),
    name:   findOne(dict.name || []),
    spec:   findOne(dict.spec || []),
    info:   findOne(dict.info || []),
  };
}

/** 取出一格文字（容錯） */
function pickCell_(row, headers, idx) {
  if (idx < 0) return '';
  const v = row[idx];
  return v == null ? '' : v;
}

/** 正規化字串 + 去除「（兩倍安全庫存: …）」與「（安全庫存: …）」備註 */
function normText_(s) {
  let t = (s == null) ? '' : String(s);
  t = t.replace(/\s*[（(]兩倍安全庫存[:：][^)）]*[)）]\s*/g, ' ');
  t = t.replace(/\s*[（(]安全庫存[:：][^)）]*[)）]\s*/g, ' ');
  return t.trim().replace(/\s+/g, ' ');
}

/** 若常見量沒有單位就補 DEFAULT_UNIT；已有單位（盒/粒/錠…）則維持 */
function ensureQtyUnit_(s) {
  const x = (s == null) ? '' : String(s).trim();
  if (!x || x.toLowerCase() === 'nan' || x.toLowerCase() === 'none') return '';
  if (/(盒|粒|顆|錠|瓶|支|包|條)\b/.test(x)) return x;
  if (/\d$/.test(x)) return x + DEFAULT_UNIT;
  return x;
}
