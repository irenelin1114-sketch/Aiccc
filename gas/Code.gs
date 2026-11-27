/**
 * 煎心鍋貼｜線上訂單後台（Google Apps Script + Google Sheet）
 *
 * 功能
 *  1. doPost：接收網站送出的訂單，驗證資料、以「商品價格」工作表重新計算金額，
 *     產生訂單編號，寫入「訂單」與「訂單明細」工作表。
 *  2. 新訂單 Email 通知店家；顧客有填 Email 時寄送訂單確認信。
 *  3. 在「訂單」工作表把「狀態」改成「可取餐／已出貨／已取消」時，自動 Email 通知顧客。
 *  4. 每天晚上寄出當日營業摘要（各門市訂單數與營業額）。
 *
 * 安裝步驟請見同資料夾的 README.md。
 */

/* ======================= 設定 ======================= */
var CONFIG = {
  TZ: 'Asia/Taipei',
  // 接收新訂單通知的 Email（多個請用逗號分隔）；留空則寄給部署這個腳本的帳號
  NOTIFY_EMAIL: '',
  SHOP_NAME: '煎心鍋貼 JIAN XIN',
  FREE_SHIPPING: 1500,
  SHIPPING_FEE: 160,
  DAILY_REPORT_HOUR: 22, // 每日摘要寄送時間（點）
  SHEETS: { ORDERS: '訂單', ITEMS: '訂單明細', PRICES: '商品價格' }
};

var ORDER_TYPES = { takeout: '外帶', dinein: '內用', delivery: '宅配' };
var STATUSES = ['新訂單', '製作中', '可取餐', '已出貨', '已完成', '已取消'];
var STORES = {
  flagship: '旗艦店（台北市大安區○○路一段 100 號）',
  xinyi: '信義店（台北市信義區○○路五段 50 號 B1）',
  nangang: '南港店（台北市南港區○○路 20 號）',
  banqiao: '板橋店（新北市板橋區○○路 88 號）'
};

// 「商品價格」工作表的初始資料（之後請直接在工作表修改價格與上下架）
var DEFAULT_PRICES = [
  ['p1', '招牌高麗菜鍋貼', 70, '否', '是'], ['p2', '韭黃鮮蝦鍋貼', 120, '否', '是'],
  ['p3', '麻辣起司鍋貼', 95, '否', '是'], ['p4', '玉米鮮肉鍋貼', 75, '否', '是'],
  ['d1', '經典韭菜水餃', 70, '否', '是'], ['d2', '鮮蝦干貝水餃', 140, '否', '是'],
  ['d3', '紅油抄手', 85, '否', '是'], ['d4', '香菇素水餃', 70, '否', '是'],
  ['s1', '酸辣湯', 45, '否', '是'], ['s2', '鮮肉貢丸湯', 40, '否', '是'],
  ['s3', '古早味紅茶', 30, '否', '是'], ['s4', '無糖豆漿', 35, '否', '是'],
  ['x1', '涼拌小黃瓜', 40, '否', '是'], ['x2', '麻辣鴨血豆腐', 60, '否', '是'],
  ['x3', '皮蛋豆腐', 50, '否', '是'], ['x4', '滷味拼盤', 70, '否', '是'],
  ['g1', '家庭常備組', 450, '是', '是'], ['g2', '招財鍋貼禮盒', 980, '是', '是'],
  ['g3', '團圓節慶組', 1680, '是', '是']
];

var ORDER_HEADERS = ['訂單編號', '建立時間', '狀態', '取餐方式', '姓名', '電話', 'Email', '門市', '取餐/到店時間',
  '宅配地址', '備註', '品項', '小計', '運費', '應付金額', '最後通知'];
var ITEM_HEADERS = ['訂單編號', '建立時間', '商品代碼', '商品名稱', '單價', '數量', '金額'];
var PRICE_HEADERS = ['商品代碼', '商品名稱', '價格', '冷凍商品(可宅配)', '上架中'];

/* ======================= 初始化（第一次請手動執行） ======================= */
function setup() {
  var ss = SpreadsheetApp.getActive();
  var orders = ensureSheet_(ss, CONFIG.SHEETS.ORDERS, ORDER_HEADERS);
  ensureSheet_(ss, CONFIG.SHEETS.ITEMS, ITEM_HEADERS);
  var prices = ensureSheet_(ss, CONFIG.SHEETS.PRICES, PRICE_HEADERS);
  if (prices.getLastRow() < 2) prices.getRange(2, 1, DEFAULT_PRICES.length, 5).setValues(DEFAULT_PRICES);

  // 狀態欄下拉選單 + 顏色
  var statusCol = ORDER_HEADERS.indexOf('狀態') + 1;
  var statusRange = orders.getRange(2, statusCol, orders.getMaxRows() - 1, 1);
  statusRange.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).build());
  var colors = { '新訂單': '#FDE7D3', '製作中': '#FFF3C4', '可取餐': '#D6F5E3', '已出貨': '#D6E8F5', '已完成': '#E8E8E8', '已取消': '#F8D7DA' };
  orders.setConditionalFormatRules(Object.keys(colors).map(function (s) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(s).setBackground(colors[s]).setRanges([statusRange]).build();
  }));
  orders.getRange('B:B').setNumberFormat('yyyy/mm/dd hh:mm');
  orders.getRange('M:O').setNumberFormat('"NT$"#,##0');

  // 觸發器：狀態變更通知、每日摘要（重複執行 setup 不會重複建立）
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['onStatusEdit', 'sendDailyReport'].indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onStatusEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendDailyReport').timeBased().everyDays(1).atHour(CONFIG.DAILY_REPORT_HOUR).inTimezone(CONFIG.TZ).create();

  SpreadsheetApp.getActive().toast('設定完成！接著請部署為網頁應用程式。', CONFIG.SHOP_NAME, 8);
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#3E2A1E').setFontColor('#F7F1E8');
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ======================= 網站 API ======================= */
function doGet() {
  return json_({ ok: true, service: CONFIG.SHOP_NAME + ' 訂單 API' });
}

function doPost(e) {
  try {
    var data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (data.website) return json_({ ok: false, error: '訂單送出失敗' }); // 防機器人欄位
    var order = buildOrder_(data);

    // 同一支電話 30 秒內不重複下單（避免重複按送出）
    var cache = CacheService.getScriptCache();
    var dupKey = 'dup-' + order.phone;
    if (cache.get(dupKey)) return json_({ ok: false, error: '訂單已送出，請稍候 30 秒再試' });

    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      order.no = nextOrderNo_(order.createdAt);
      saveOrder_(order);
    } finally {
      lock.releaseLock();
    }
    cache.put(dupKey, '1', 30);

    try { notifyShop_(order); } catch (err) { console.error('店家通知失敗', err); }
    try { if (order.email) sendCustomerMail_(order, '新訂單'); } catch (err) { console.error('顧客確認信失敗', err); }

    return json_({ ok: true, orderNo: order.no, subtotal: order.subtotal, shipping: order.shipping, total: order.total });
  } catch (err) {
    if (!(err && err.userMessage)) console.error(err);
    return json_({ ok: false, error: err && err.userMessage ? err.userMessage : '系統忙碌中，請稍後再試或直接致電門市' });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function userError_(msg) { var e = new Error(msg); e.userMessage = msg; return e; }

/* ======================= 訂單驗證與計算 ======================= */
function buildOrder_(d) {
  var type = String(d.type || '');
  if (!ORDER_TYPES[type]) throw userError_('取餐方式不正確');
  var name = clean_(d.name, 30);
  if (!name) throw userError_('請填寫姓名');
  var phone = String(d.phone || '').replace(/[-\s]/g, '');
  if (!/^09\d{8}$/.test(phone)) throw userError_('手機號碼格式不正確');
  var email = clean_(d.email, 100);
  if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw userError_('Email 格式不正確');

  var prices = loadPrices_();
  if (!Array.isArray(d.items) || !d.items.length || d.items.length > 50) throw userError_('購物車是空的');
  var items = d.items.map(function (it) {
    var p = prices[String(it.id)];
    var qty = Math.floor(Number(it.qty));
    if (!p || !p.active) throw userError_('部分商品已下架，請重新整理頁面');
    if (!(qty >= 1 && qty <= 99)) throw userError_('商品數量不正確');
    if (type === 'delivery' && !p.frozen) throw userError_('宅配僅限冷凍禮盒');
    return { id: p.id, name: p.name, price: p.price, qty: qty, amount: p.price * qty };
  });

  var store = '', time = '', address = '';
  if (type === 'delivery') {
    address = clean_(d.address, 120);
    if (address.length < 8) throw userError_('請填寫完整的宅配地址');
  } else {
    if (!STORES[d.storeId]) throw userError_('請選擇門市');
    store = STORES[d.storeId];
    time = absoluteTime_(clean_(d.time, 20));
    if (!time) throw userError_('請選擇時間');
  }

  var subtotal = items.reduce(function (s, it) { return s + it.amount; }, 0);
  var shipping = type === 'delivery' && subtotal < CONFIG.FREE_SHIPPING ? CONFIG.SHIPPING_FEE : 0;
  return {
    createdAt: new Date(), type: type, typeLabel: ORDER_TYPES[type], name: name, phone: phone, email: email,
    store: store, time: time, address: address, note: clean_(d.note, 200),
    items: items, subtotal: subtotal, shipping: shipping, total: subtotal + shipping
  };
}

// 把網站傳來的「今天 12:00／明天 12:00」換成實際日期，例如「11/28（五）12:00」
function absoluteTime_(t) {
  var m = /^(今天|明天)\s*(\d{1,2}:\d{2})$/.exec(t);
  if (!m) return t;
  var d = new Date(Date.now() + (m[1] === '明天' ? 86400000 : 0));
  var wd = '日一二三四五六'.charAt(Number(Utilities.formatDate(d, CONFIG.TZ, 'u')) % 7);
  return Utilities.formatDate(d, CONFIG.TZ, 'MM/dd') + '（' + wd + '）' + m[2];
}

function loadPrices_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.PRICES);
  var rows = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 5).getValues();
  var map = {};
  rows.forEach(function (r) {
    if (!r[0]) return;
    map[String(r[0])] = { id: String(r[0]), name: String(r[1]), price: Number(r[2]), frozen: r[3] === '是', active: r[4] !== '否' };
  });
  return map;
}

// 去除控制字元與前後空白、限制長度
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max);
}

// 寫入試算表時，避免以 = + - @ 開頭的文字被當成公式執行
function cell_(v) {
  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v;
}

function nextOrderNo_(date) {
  var day = Utilities.formatDate(date, CONFIG.TZ, 'yyMMdd');
  var props = PropertiesService.getScriptProperties();
  var key = 'seq-' + day;
  var n = Number(props.getProperty(key) || 0) + 1;
  props.setProperty(key, String(n));
  return 'JX' + day + '-' + ('000' + n).slice(-4);
}

function saveOrder_(o) {
  var ss = SpreadsheetApp.getActive();
  var summary = o.items.map(function (it) { return it.name + ' × ' + it.qty; }).join('、');
  ss.getSheetByName(CONFIG.SHEETS.ORDERS).appendRow([
    o.no, o.createdAt, '新訂單', o.typeLabel, o.name, "'" + o.phone, o.email, o.store, o.time,
    o.address, o.note, summary, o.subtotal, o.shipping, o.total, ''
  ].map(cell_));
  var rows = o.items.map(function (it) { return [o.no, o.createdAt, it.id, it.name, it.price, it.qty, it.amount].map(cell_); });
  var items = ss.getSheetByName(CONFIG.SHEETS.ITEMS);
  items.getRange(items.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/* ======================= Email 通知 ======================= */
function shopEmail_() { return CONFIG.NOTIFY_EMAIL || Session.getEffectiveUser().getEmail(); }

function itemsHtml_(o) {
  return '<table style="border-collapse:collapse;width:100%;font-size:14px">' + o.items.map(function (it) {
    return '<tr><td style="padding:4px 0">' + esc_(it.name) + ' × ' + it.qty + '</td><td style="text-align:right">NT$' + it.amount.toLocaleString() + '</td></tr>';
  }).join('') +
  '<tr><td style="padding-top:8px;border-top:1px solid #ddd">運費</td><td style="text-align:right;padding-top:8px;border-top:1px solid #ddd">' + (o.shipping ? 'NT$' + o.shipping : '免運') + '</td></tr>' +
  '<tr><td><b>應付金額</b></td><td style="text-align:right"><b>NT$' + o.total.toLocaleString() + '</b></td></tr></table>';
}

function esc_(s) { return String(s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

function notifyShop_(o) {
  var where = o.type === 'delivery' ? '宅配地址：' + esc_(o.address) : '門市：' + esc_(o.store) + '<br>時間：' + esc_(o.time);
  MailApp.sendEmail({
    to: shopEmail_(),
    subject: '【新訂單】' + o.no + '｜' + o.typeLabel + '｜NT$' + o.total.toLocaleString(),
    htmlBody: '<p><b>' + esc_(o.name) + '</b>（' + o.phone + (o.email ? '，' + esc_(o.email) : '') + '）</p>' +
      '<p>取餐方式：' + o.typeLabel + '<br>' + where + '</p>' +
      (o.note ? '<p>備註：' + esc_(o.note) + '</p>' : '') + itemsHtml_(o) +
      '<p style="color:#888;font-size:12px">在 Google Sheet「訂單」工作表更新狀態即可通知顧客。<br>' + SpreadsheetApp.getActive().getUrl() + '</p>'
  });
}

var STATUS_MAIL = {
  '新訂單': { subject: '訂單確認', body: '感謝您的訂購！我們已收到您的訂單，將盡快為您準備。' },
  '可取餐': { subject: '餐點已完成，可以取餐囉', body: '您的餐點已經煎好了！請於預約時間到門市取餐，取餐時請告知訂單編號。' },
  '已出貨': { subject: '禮盒已出貨', body: '您訂購的冷凍禮盒已交由黑貓低溫宅配寄出，預計 1–2 天內送達。' },
  '已取消': { subject: '訂單已取消', body: '您的訂單已取消。如有疑問，請直接來電門市，造成不便敬請見諒。' }
};

function sendCustomerMail_(o, status) {
  var m = STATUS_MAIL[status];
  if (!m || !o.email) return;
  MailApp.sendEmail({
    to: o.email,
    name: CONFIG.SHOP_NAME,
    subject: '【' + CONFIG.SHOP_NAME + '】' + m.subject + '（' + o.no + '）',
    htmlBody: '<p>' + esc_(o.name) + ' 您好：</p><p>' + m.body + '</p>' +
      '<p>訂單編號：<b>' + o.no + '</b><br>取餐方式：' + o.typeLabel +
      (o.type === 'delivery' ? '<br>寄送地址：' + esc_(o.address) : '<br>門市：' + esc_(o.store) + '<br>時間：' + esc_(o.time)) + '</p>' +
      (o.items ? itemsHtml_(o) : '<p>應付金額：NT$' + Number(o.total).toLocaleString() + '</p>') +
      '<p style="color:#888;font-size:12px">此信件由系統自動寄出，請勿直接回覆。</p>'
  });
}

/* ======================= 狀態變更 → 通知顧客（可安裝觸發器） ======================= */
function onStatusEdit(e) {
  var sh = e.range.getSheet();
  if (sh.getName() !== CONFIG.SHEETS.ORDERS) return;
  var statusCol = ORDER_HEADERS.indexOf('狀態') + 1;
  if (e.range.getColumn() > statusCol || e.range.getLastColumn() < statusCol || e.range.getRow() < 2) return;
  for (var r = e.range.getRow(); r <= e.range.getLastRow(); r++) {
    var row = sh.getRange(r, 1, 1, ORDER_HEADERS.length).getValues()[0];
    var get = function (h) { return row[ORDER_HEADERS.indexOf(h)]; };
    var status = get('狀態');
    if (!STATUS_MAIL[status] || status === '新訂單' || !get('Email')) continue;
    if (String(get('最後通知')).indexOf(status) === 0) continue; // 同一狀態不重複寄
    var typeKey = Object.keys(ORDER_TYPES).filter(function (k) { return ORDER_TYPES[k] === get('取餐方式'); })[0];
    sendCustomerMail_({
      no: get('訂單編號'), name: get('姓名'), email: get('Email'), type: typeKey, typeLabel: get('取餐方式'),
      store: get('門市'), time: get('取餐/到店時間'), address: get('宅配地址'), total: get('應付金額')
    }, status);
    sh.getRange(r, ORDER_HEADERS.indexOf('最後通知') + 1)
      .setValue(status + ' 已通知 ' + Utilities.formatDate(new Date(), CONFIG.TZ, 'MM/dd HH:mm'));
  }
}

/* ======================= 每日營業摘要 ======================= */
function sendDailyReport() {
  var sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEETS.ORDERS);
  if (sh.getLastRow() < 2) return;
  var today = Utilities.formatDate(new Date(), CONFIG.TZ, 'yyyy/MM/dd');
  var rows = sh.getRange(2, 1, sh.getLastRow() - 1, ORDER_HEADERS.length).getValues().filter(function (r) {
    return r[1] instanceof Date && Utilities.formatDate(r[1], CONFIG.TZ, 'yyyy/MM/dd') === today && r[2] !== '已取消';
  });
  var byGroup = {}, total = 0;
  rows.forEach(function (r) {
    var key = r[3] === '宅配' ? '宅配' : String(r[7]).split('（')[0];
    byGroup[key] = byGroup[key] || { n: 0, sum: 0 };
    byGroup[key].n++; byGroup[key].sum += Number(r[14]) || 0; total += Number(r[14]) || 0;
  });
  var lines = Object.keys(byGroup).map(function (k) {
    return '<tr><td style="padding:4px 12px 4px 0">' + k + '</td><td>' + byGroup[k].n + ' 筆</td><td style="text-align:right;padding-left:12px">NT$' + byGroup[k].sum.toLocaleString() + '</td></tr>';
  }).join('');
  MailApp.sendEmail({
    to: shopEmail_(),
    subject: '【每日摘要】' + today + '｜' + rows.length + ' 筆訂單｜NT$' + total.toLocaleString(),
    htmlBody: rows.length ? '<table style="font-size:14px">' + lines + '</table><p><b>合計：' + rows.length + ' 筆，NT$' + total.toLocaleString() + '</b>（不含已取消）</p>' : '<p>今天沒有線上訂單。</p>'
  });
}
