/**
 * วิธีติดตั้ง (ทำครั้งเดียว):
 * 1. ไปที่ https://sheets.google.com สร้าง Google Sheet ใหม่ (ไฟล์ว่างๆ ก็ได้ ชื่ออะไรก็ได้)
 * 2. เมนู ส่วนขยาย (Extensions) > Apps Script
 * 3. ลบโค้ดเดิมทั้งหมดในไฟล์ Code.gs แล้ววางโค้ดทั้งหมดในไฟล์นี้แทน
 * 4. กด Deploy (ปุ่มสีน้ำเงินมุมขวาบน) > New deployment
 * 5. เลือก Type เป็น "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. กด Deploy แล้วอนุญาตสิทธิ์ (Authorize access) ตามที่ระบบขอ (สคริปต์นี้ต้องใช้สิทธิ์เข้าถึง Google Drive
 *    ด้วย เพื่อเก็บไฟล์รูปภาพสินค้าในโฟลเดอร์ "POS_ProductImages")
 * 7. คัดลอก "Web app URL" ที่ได้ (ลงท้ายด้วย /exec)
 * 8. นำ URL ไปวางในหน้า "ตั้งค่า" ของเว็บแอพ POS แล้วกด "บันทึกและซิงค์"
 *
 * หมายเหตุ: ทุกครั้งที่แก้โค้ดนี้ ต้องกด Deploy > Manage deployments > แก้ไข (ไอคอนดินสอ) > Version: New > Deploy ใหม่
 */

function doGet(e) {
  return handle(e);
}

function doPost(e) {
  return handle(e);
}

function handle(e) {
  var result;
  try {
    var action = (e.parameter && e.parameter.action) || '';
    var body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
      if (!action) action = body.action;
    }

    if (action === 'load') {
      result = {
        products: readSheet('Products'),
        sales: readSalesWithItems(),
        stockIns: readSheet('StockIns')
      };
    } else if (action === 'saveProducts') {
      writeProductsSheet(body.products || []);
      result = { status: 'ok' };
    } else if (action === 'saveSales') {
      writeSalesSheet(body.sales || []);
      result = { status: 'ok' };
    } else if (action === 'saveStockIns') {
      writeStockInsSheet(body.stockIns || []);
      result = { status: 'ok' };
    } else if (action === 'uploadImage') {
      result = { status: 'ok', url: uploadImageToDrive(body.imageData, body.filename) };
    } else if (action === 'ping') {
      result = { status: 'ok', message: 'เชื่อมต่อสำเร็จ' };
    } else {
      result = { error: 'unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// โครงสร้างฐานข้อมูล (แต่ละชีต = 1 ตาราง)
//
//  Products   (ตารางหลักสินค้า)
//    id (PK), name, barcode (UNIQUE), price, cost, stock, image,
//    category (ชื่อหมวดหมู่/กลุ่มสินค้า แบบข้อความอิสระ, ใส่ว่างได้), unit, minStock,
//    createdAt, updatedAt
//
//  Sales      (หัวบิลขาย - 1 แถวต่อ 1 บิล)
//    id (PK), billNo, datetime, total, received, change, paymentMethod
//
//  SaleItems  (รายการสินค้าในบิล - normalize ออกจาก Sales เดิม)
//    id (PK), saleId (FK -> Sales.id), productId (FK -> Products.id),
//    productName, price, cost, qty, subtotal
//    -> 1 Sales มีได้หลาย SaleItems (one-to-many)
//
//  StockIns   (ประวัติการรับสินค้าเข้าสต๊อก)
//    id (PK), datetime, productId (FK -> Products.id), productName,
//    qty, cost, totalCost, note
//
// เหตุผลที่แยก SaleItems ออกจาก Sales (เดิมเก็บเป็น JSON string ในคอลัมน์ items):
//  - เปิดใช้ Pivot Table / SUMIF / QUERY บน Google Sheets เพื่อดูยอดขายรายสินค้าได้ตรงๆ
//  - เชื่อมกับ Products ผ่าน productId เพื่อดึงชื่อ/ต้นทุนล่าสุดได้ ไม่ผูกกับข้อความที่ก็อปปี้มา
//  - แก้ปัญหาแถวข้อมูลใหญ่เกินไปเมื่อบิลมีสินค้าเยอะ
// ---------------------------------------------------------------------------

var HEADERS = {
  Products: ['id', 'name', 'barcode', 'price', 'cost', 'stock', 'image', 'category', 'unit', 'minStock', 'createdAt', 'updatedAt'],
  Sales: ['id', 'billNo', 'datetime', 'total', 'received', 'change', 'paymentMethod'],
  SaleItems: ['id', 'saleId', 'productId', 'productName', 'price', 'cost', 'qty', 'subtotal'],
  StockIns: ['id', 'datetime', 'productId', 'productName', 'qty', 'cost', 'totalCost', 'note']
};

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getImageFolder() {
  var folderName = 'POS_ProductImages';
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function uploadImageToDrive(dataUrl, filename) {
  var match = dataUrl.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.*)$/);
  if (!match) throw new Error('รูปแบบข้อมูลรูปภาพไม่ถูกต้อง');
  var mimeType = match[1];
  var base64Data = match[2];
  var bytes = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(bytes, mimeType, filename || ('product-' + Date.now()));
  var folder = getImageFolder();
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://lh3.googleusercontent.com/d/' + file.getId();
}

function ensureSheet(name) {
  var ss = getSS();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
  }
  return sheet;
}

function readSheet(name) {
  var sheet = ensureSheet(name);
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = values.slice(1);
  return rows
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function readSalesWithItems() {
  var sales = readSheet('Sales');
  var items = readSheet('SaleItems');
  var itemsBySaleId = {};
  items.forEach(function (it) {
    if (!itemsBySaleId[it.saleId]) itemsBySaleId[it.saleId] = [];
    itemsBySaleId[it.saleId].push({
      productId: it.productId,
      name: it.productName,
      price: it.price,
      cost: it.cost,
      qty: it.qty
    });
  });
  return sales.map(function (s) {
    return {
      id: s.id,
      billNo: s.billNo,
      datetime: s.datetime,
      items: itemsBySaleId[s.id] || [],
      total: s.total,
      received: s.received,
      change: s.change,
      paymentMethod: s.paymentMethod || 'cash'
    };
  });
}

function writeProductsSheet(products) {
  var headers = HEADERS.Products;
  var sheet = ensureSheet('Products');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (products.length === 0) return;
  var rows = products.map(function (p) {
    return headers.map(function (h) { return p[h] !== undefined ? p[h] : ''; });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

function writeSalesSheet(sales) {
  var headers = HEADERS.Sales;
  var sheet = ensureSheet('Sales');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  var itemHeaders = HEADERS.SaleItems;
  var itemSheet = ensureSheet('SaleItems');
  itemSheet.clearContents();
  itemSheet.getRange(1, 1, 1, itemHeaders.length).setValues([itemHeaders]);

  if (sales.length === 0) return;

  var rows = sales.map(function (s) {
    return [s.id, s.billNo, s.datetime, s.total, s.received, s.change, s.paymentMethod || 'cash'];
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);

  var itemRows = [];
  sales.forEach(function (s) {
    (s.items || []).forEach(function (it) {
      var qty = Number(it.qty) || 0;
      var price = Number(it.price) || 0;
      itemRows.push([
        s.id + '-' + it.productId,
        s.id,
        it.productId,
        it.name,
        price,
        it.cost !== undefined ? it.cost : '',
        qty,
        qty * price
      ]);
    });
  });
  if (itemRows.length > 0) {
    itemSheet.getRange(2, 1, itemRows.length, itemHeaders.length).setValues(itemRows);
  }
}

function writeStockInsSheet(stockIns) {
  var headers = HEADERS.StockIns;
  var sheet = ensureSheet('StockIns');
  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (stockIns.length === 0) return;
  var rows = stockIns.map(function (r) {
    return headers.map(function (h) { return r[h] !== undefined ? r[h] : ''; });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}
