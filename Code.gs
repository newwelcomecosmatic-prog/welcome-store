/**
 * Welcome Cosmetic Store - Inventory (Google Sheet + Apps Script)
 * Sheet ke andar: Extensions > Apps Script mein yeh file paste karein.
 *
 * Products sheet : Barcode | Product | Purchase Price | Sale Price | Stock
 * Transactions   : Date | Type | Bill No | Barcode | Product | Qty | Price | Total | Customer | Cost Price | Profit
 */
const SHOP_NAME = 'Welcome Cosmetic Store';
const SHOP_ADDRESS = 'Main Bazar Begumkot Shahdara Lahore';
const FOLDER_NAME = 'Welcome Cosmetic Store - PDFs';
const TZ = 'Asia/Karachi';

// Theme (#e8bcf0)
const C_THEME = '#e8bcf0';
const C_DARK = '#6b2d7b';
const C_HEADTXT = '#3d1347';

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// App ka password: installable app ki Settings mein yahi dalna hai. Zaroor badal dein.
const API_KEY = 'welcome123';

/** Browser mein /exec link kholein to "chal rahi hai" likha aana chahiye */
function doGet() {
  return json_({ ok: true, data: SHOP_NAME + ' API chal rahi hai' });
}

/** Installable app (PWA) yahan JSON bhej kar kaam karwati hai */
function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    if (String(req.key || '') !== API_KEY) throw new Error('Password ghalat hai');
    const p = req.payload || {};
    let data;
    switch (req.action) {
      case 'getProducts': data = getProducts(); break;
      case 'addProduct': data = addProduct(p); break;
      case 'saveTransaction': data = saveTransaction(p); break;
      case 'getStockReport': data = getStockReport(); break;
      case 'getSalesReport': data = getSalesReport(p.from, p.to, !!p.makePdf); break;
      default: throw new Error('Unknown action');
    }
    return json_({ ok: true, data: data });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * Editor se ek baar Run karein (pehli dafa ya update ke baad dobara).
 * Purani sheet ho to "Price" column ko "Sale Price" bana kar "Purchase Price" column add kar deta hai.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let p = ss.getSheetByName('Products');
  if (!p) p = ss.insertSheet('Products');
  if (p.getLastRow() === 0) {
    p.appendRow(['Barcode', 'Product', 'Purchase Price', 'Sale Price', 'Stock']);
    p.setFrozenRows(1);
  } else if (String(p.getRange('C1').getValue()).trim() === 'Price') {
    p.insertColumnAfter(2);
    p.getRange('C1').setValue('Purchase Price');
    p.getRange('D1').setValue('Sale Price');
  }
  p.getRange('A1:E1').setBackground(C_THEME).setFontColor(C_HEADTXT).setFontWeight('bold');
  p.getRange('A:A').setNumberFormat('@');

  let t = ss.getSheetByName('Transactions');
  if (!t) t = ss.insertSheet('Transactions');
  if (t.getLastRow() === 0) {
    t.appendRow(['Date', 'Type', 'Bill No', 'Barcode', 'Product', 'Qty', 'Price', 'Total', 'Customer', 'Cost Price', 'Profit']);
    t.setFrozenRows(1);
  } else if (!String(t.getRange('J1').getValue()).trim()) {
    t.getRange('J1:K1').setValues([['Cost Price', 'Profit']]);
  }
  t.getRange('A1:K1').setBackground(C_THEME).setFontColor(C_HEADTXT).setFontWeight('bold');
  t.getRange('A2:A').setNumberFormat('dd-MMM-yyyy hh:mm AM/PM');
  t.getRange('D:D').setNumberFormat('@');

  getFolder_();
}

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" sheet nahi mili. setup() Run karein.');
  return sh;
}

function readProducts_() {
  const sh = sheet_('Products');
  const n = sh.getLastRow() - 1;
  if (n < 1) return { sh: sh, rows: [] };
  const v = sh.getRange(2, 1, n, 5).getValues();
  const rows = [];
  v.forEach(function (r, i) {
    const code = String(r[0]).trim();
    if (!code) return;
    rows.push({
      row: i + 2,
      barcode: code,
      name: String(r[1]),
      purchase: Number(r[2]) || 0,
      sale: Number(r[3]) || 0,
      stock: Number(r[4]) || 0
    });
  });
  return { sh: sh, rows: rows };
}

function getProducts() {
  return readProducts_().rows.map(function (r) {
    return { barcode: r.barcode, name: r.name, purchase: r.purchase, sale: r.sale, stock: r.stock };
  });
}

/** Naya product: barcode, name, purchase, sale, stock (opening, optional) */
function addProduct(p) {
  const code = String(p.barcode || '').trim();
  const name = String(p.name || '').trim();
  const purchase = Number(p.purchase) || 0;
  const sale = Number(p.sale) || 0;
  const stock = Math.max(0, Math.floor(Number(p.stock) || 0));
  if (!code || !name) throw new Error('Barcode aur naam zaroori hain');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const d = readProducts_();
    if (d.rows.some(function (r) { return r.barcode === code; })) throw new Error('Yeh barcode pehle se maujood hai');
    const row = d.sh.getLastRow() + 1;
    d.sh.getRange(row, 1, 1, 5).setValues([[code, name, purchase, sale, stock]]);
    return { barcode: code, name: name, purchase: purchase, sale: sale, stock: stock };
  } finally {
    lock.releaseLock();
  }
}

/**
 * type 'SALE' : stock minus, item.price = sale price
 * type 'IN'   : stock plus,  item.price = purchase price (product ka purchase price update ho jata hai)
 */
function saveTransaction(p) {
  if (!p || !p.items || !p.items.length) throw new Error('Cart khali hai');
  const type = p.type === 'IN' ? 'IN' : 'SALE';
  const customer = String(p.customer || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let result;
  try {
    const d = readProducts_();
    const map = {};
    d.rows.forEach(function (r) { map[r.barcode] = r; });

    const agg = {};
    const order = [];
    p.items.forEach(function (it) {
      const code = String(it.barcode).trim();
      const qty = Math.floor(Number(it.qty));
      const prod = map[code];
      if (!prod) throw new Error('Product nahi mila: ' + code);
      if (!(qty > 0)) throw new Error('Qty ghalat hai: ' + prod.name);
      if (!agg[code]) {
        const given = Number(it.price);
        const price = given >= 0 ? given : (type === 'SALE' ? prod.sale : prod.purchase);
        agg[code] = { prod: prod, qty: 0, price: price };
        order.push(code);
      }
      agg[code].qty += qty;
    });

    const lines = order.map(function (c) { return agg[c]; });
    lines.forEach(function (l) {
      if (type === 'SALE' && l.qty > l.prod.stock) {
        throw new Error(l.prod.name + ' ka stock kam hai (available ' + l.prod.stock + ')');
      }
    });

    const props = PropertiesService.getScriptProperties();
    const key = type === 'SALE' ? 'billSale' : 'billIn';
    const n = Number(props.getProperty(key) || 0) + 1;
    props.setProperty(key, String(n));
    const billNo = (type === 'SALE' ? 'B-' : 'IN-') + Utilities.formatString('%05d', n);
    const now = new Date();
    const dateStr = Utilities.formatDate(now, TZ, 'dd-MMM-yyyy hh:mm a');

    const vals = [];
    lines.forEach(function (l) {
      const prod = l.prod;
      if (type === 'SALE') {
        d.sh.getRange(prod.row, 5).setValue(prod.stock - l.qty);
        const cost = prod.purchase;
        vals.push([now, type, billNo, prod.barcode, prod.name, l.qty, l.price, l.qty * l.price, customer, cost, (l.price - cost) * l.qty]);
      } else {
        d.sh.getRange(prod.row, 5).setValue(prod.stock + l.qty);
        if (l.price !== prod.purchase) d.sh.getRange(prod.row, 3).setValue(l.price);
        vals.push([now, type, billNo, prod.barcode, prod.name, l.qty, l.price, l.qty * l.price, customer, l.price, '']);
      }
    });
    const tr = sheet_('Transactions');
    tr.getRange(tr.getLastRow() + 1, 1, vals.length, 11).setValues(vals);
    SpreadsheetApp.flush();

    result = { billNo: billNo, type: type, lines: lines, dateStr: dateStr, customer: customer };
  } finally {
    lock.releaseLock();
  }

  // PDF lock ke bahar; PDF fail ho to bhi stock save rehta hai
  const out = { ok: true, billNo: result.billNo };
  try {
    out.pdf = makeBillPdf_(result);
  } catch (e) {
    out.pdfError = String(e.message || e);
  }
  return out;
}

function makeBillPdf_(r) {
  const meta = ['Bill No: ' + r.billNo, 'Date: ' + r.dateStr];
  if (r.customer) meta.push('Customer: ' + r.customer);
  if (r.type === 'SALE') {
    let total = 0;
    const rows = r.lines.map(function (l, i) {
      total += l.qty * l.price;
      return [String(i + 1), l.prod.name, String(l.qty), fmt_(l.price), fmt_(l.qty * l.price)];
    });
    return buildPdf_(r.billNo + '.pdf', 'BILL', meta,
      ['#', 'Product', 'Qty', 'Price', 'Total'], rows, [30, 245, 50, 90, 100], [2, 3, 4],
      'Grand Total: Rs. ' + fmt_(total));
  }
  let cost = 0;
  const rowsIn = r.lines.map(function (l, i) {
    cost += l.qty * l.price;
    return [String(i + 1), l.prod.barcode, l.prod.name, String(l.qty), fmt_(l.price)];
  });
  return buildPdf_(r.billNo + '.pdf', 'STOCK IN RECEIPT', meta,
    ['#', 'Barcode', 'Product', 'Qty', 'Purchase Price'], rowsIn, [30, 120, 225, 50, 90], [3, 4],
    'Total Purchase: Rs. ' + fmt_(cost));
}

/** Stock report PDF: har product ka available stock */
function getStockReport() {
  const rows = readProducts_().rows.slice().sort(function (a, b) { return a.name.localeCompare(b.name); });
  let units = 0, costVal = 0, saleVal = 0;
  const data = rows.map(function (r, i) {
    units += r.stock;
    costVal += r.stock * r.purchase;
    saleVal += r.stock * r.sale;
    return [String(i + 1), r.barcode, r.name, fmt_(r.purchase), fmt_(r.sale), String(r.stock)];
  });
  const now = Utilities.formatDate(new Date(), TZ, 'dd-MMM-yyyy hh:mm a');
  const stamp = Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmm');
  return buildPdf_('Stock_Report_' + stamp + '.pdf', 'STOCK REPORT',
    ['Date: ' + now,
     'Total products: ' + rows.length + '   |   Total units: ' + units,
     'Stock value (purchase): Rs. ' + fmt_(costVal) + '   |   Stock value (sale): Rs. ' + fmt_(saleVal)],
    ['#', 'Barcode', 'Product', 'Purchase', 'Sale', 'Available'], data, [25, 100, 200, 65, 60, 65], [],
    '');
}

/**
 * Sale report (date range) with profit.
 * from / to: 'yyyy-MM-dd'. makePdf false = sirf summary.
 */
function getSalesReport(from, to, makePdf) {
  from = String(from || '');
  to = String(to || '');
  if (!from || !to) throw new Error('From aur To date chunein');
  if (from > to) { const t = from; from = to; to = t; }

  const prods = {};
  readProducts_().rows.forEach(function (r) { prods[r.barcode] = r; });

  const tr = sheet_('Transactions');
  const n = tr.getLastRow() - 1;
  const vals = n < 1 ? [] : tr.getRange(2, 1, n, 11).getValues();

  const lines = [];
  const bills = {};
  let sales = 0, cost = 0, units = 0;
  vals.forEach(function (r) {
    if (String(r[1]) !== 'SALE') return;
    const day = dayKey_(r[0]);
    if (!day || day < from || day > to) return;
    const qty = Number(r[5]) || 0;
    const total = Number(r[7]) || 0;
    let unitCost;
    if (r[9] === '' || r[9] === null) {
      const pr = prods[String(r[3]).trim()];
      unitCost = pr ? pr.purchase : 0;
    } else {
      unitCost = Number(r[9]) || 0;
    }
    const lineCost = unitCost * qty;
    sales += total;
    cost += lineCost;
    units += qty;
    bills[String(r[2])] = 1;
    lines.push({ day: day, bill: String(r[2]), name: String(r[4]), qty: qty, total: total, profit: total - lineCost });
  });

  const summary = {
    from: from, to: to,
    bills: Object.keys(bills).length, units: units,
    sales: sales, cost: cost, profit: sales - cost
  };
  const out = { summary: summary };
  if (makePdf) out.pdf = makeSalesPdf_(summary, lines);
  return out;
}

function makeSalesPdf_(s, lines) {
  const meta = [
    'Period: ' + prettyDay_(s.from, true) + '  to  ' + prettyDay_(s.to, true),
    'Total bills: ' + s.bills + '   |   Units sold: ' + s.units,
    'Total sale: Rs. ' + fmt_(s.sales) + '   |   Total cost: Rs. ' + fmt_(s.cost)
  ];
  let rows = lines.map(function (l, i) {
    return [String(i + 1), prettyDay_(l.day, false), l.bill, l.name, String(l.qty), fmt_(l.total), fmt_(l.profit)];
  });
  if (!rows.length) rows = [['', '', '', 'Is muddat mein koi sale nahi', '', '', '']];
  return buildPdf_('Sale_Report_' + s.from + '_to_' + s.to + '.pdf', 'SALE REPORT', meta,
    ['#', 'Date', 'Bill', 'Product', 'Qty', 'Sale', 'Profit'], rows,
    [25, 65, 55, 190, 30, 70, 80], [4, 5, 6],
    'Net Profit: Rs. ' + fmt_(s.profit));
}

/** Cell ki date ko 'yyyy-MM-dd' mein badalta hai (Date object ya purani text date dono) */
function dayKey_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = String(v).match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})/);
  if (!m) return '';
  const mi = MONTHS.map(function (x) { return x.toLowerCase(); }).indexOf(m[2].toLowerCase());
  if (mi < 0) return '';
  return m[3] + '-' + ('0' + (mi + 1)).slice(-2) + '-' + ('0' + m[1]).slice(-2);
}

function prettyDay_(key, fullYear) {
  const p = key.split('-');
  return p[2] + '-' + MONTHS[Number(p[1]) - 1] + '-' + (fullYear ? p[0] : p[0].slice(2));
}

function fmt_(n) {
  return Number(n).toLocaleString('en-US', { maximumFractionDigits: 2 });
}

function getFolder_() {
  const it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

/** A4 PDF: header = shop name, footer = address (har page par). */
function buildPdf_(fileName, title, metaLines, header, rows, widths, rightCols, bottomText) {
  const doc = DocumentApp.create('tmp_' + fileName);
  const body = doc.getBody();
  body.setPageWidth(595.28);
  body.setPageHeight(841.89);
  body.setMarginTop(60);
  body.setMarginBottom(60);
  body.setMarginLeft(40);
  body.setMarginRight(40);

  const hp = doc.addHeader().getChild(0).asParagraph();
  hp.setText(SHOP_NAME);
  hp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  hp.setBold(true);
  hp.setFontSize(22);
  hp.setForegroundColor(C_DARK);

  const fp = doc.addFooter().getChild(0).asParagraph();
  fp.setText('Address: ' + SHOP_ADDRESS);
  fp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  fp.setFontSize(9);
  fp.setForegroundColor('#555555');

  const tp = body.getChild(0).asParagraph();
  tp.setText(title);
  tp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  tp.setBold(true);
  tp.setFontSize(14);
  tp.setForegroundColor(C_DARK);

  metaLines.forEach(function (m) {
    const mp = body.appendParagraph(m);
    mp.setFontSize(10);
    mp.setBold(false);
    mp.setForegroundColor('#222222');
    mp.setSpacingAfter(0);
  });
  body.appendParagraph('');

  const table = body.appendTable([header].concat(rows));
  const attrs = {};
  attrs[DocumentApp.Attribute.FONT_SIZE] = 9;
  table.setAttributes(attrs);
  table.setBorderColor('#c9a6d1');
  table.setBorderWidth(0.5);
  widths.forEach(function (w, i) { table.setColumnWidth(i, w); });

  const hr = table.getRow(0);
  for (let c = 0; c < header.length; c++) {
    const cell = hr.getCell(c);
    cell.setBackgroundColor(C_THEME);
    cell.editAsText().setForegroundColor(C_HEADTXT).setBold(true);
  }
  // bohat bari table mein alignment skip (PDF jaldi banti hai)
  if (rightCols.length && table.getNumRows() <= 150) {
    for (let r = 0; r < table.getNumRows(); r++) {
      rightCols.forEach(function (c) {
        table.getRow(r).getCell(c).getChild(0).asParagraph()
          .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
      });
    }
  }

  if (bottomText) {
    body.appendParagraph('');
    const bp = body.appendParagraph(bottomText);
    bp.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    bp.setBold(true);
    bp.setFontSize(12);
    bp.setForegroundColor(C_DARK);
  }

  doc.saveAndClose();
  const tmp = DriveApp.getFileById(doc.getId());
  const pdf = tmp.getAs(MimeType.PDF).setName(fileName);
  const out = getFolder_().createFile(pdf);
  out.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  tmp.setTrashed(true);
  return {
    name: fileName,
    url: out.getUrl(),
    downloadUrl: 'https://drive.google.com/uc?export=download&id=' + out.getId()
  };
}
