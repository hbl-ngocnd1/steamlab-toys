/**
 * STEAMLab — Backend Google Apps Script
 * =============================================================
 * Một Web App nhỏ phục vụ 2 việc cho site tĩnh (GitHub Pages):
 *   1) Đếm lượt truy cập   → ?action=hit  (mặc định)  → tăng + ghi log.
 *   2) Trả dữ liệu sản phẩm → ?action=products         → đọc sheet "products".
 *
 * CÁCH CÀI ĐẶT / CẬP NHẬT
 * ------------------------
 * 1. Mở Google Sheet → Tiện ích mở rộng (Extensions) → Apps Script.
 * 2. Dán TOÀN BỘ file này vào, lưu.
 * 3. Chạy hàm  seedProducts  MỘT LẦN (chọn hàm ở thanh trên → Run) để tạo
 *    sheet "products" và đổ sẵn 12 sản phẩm. Cấp quyền khi được hỏi.
 * 4. Triển khai: Deploy → Manage deployments → (deployment hiện có) → Edit
 *    → Version: New version → Deploy.  (Giữ nguyên URL /exec cũ.)
 *    Lần đầu thì: New deployment → Web app → Execute as: Me, Access: Anyone.
 * 5. URL /exec dán vào COUNTER_API trong index.html, guide.html và products.js.
 *
 * SHEET "products" — cột (hàng 1 là header):
 *   id | name | price | originalPrice | rating | sold | ageRange | steam | guideNote | link | order | active
 *   - price/originalPrice: số nguyên VND (vd 65000). % giảm tính ở client.
 *   - sold: chuỗi hiển thị (vd "4,1k").  - steam: danh mục cách nhau bởi dấu phẩy.
 *   - active: TRUE/FALSE để ẩn/hiện.    - order: thứ tự sắp xếp.
 */

var COUNTER_SHEET = 'counter';
var LOG_SHEET = 'log';
var PRODUCTS_SHEET = 'products';
var PRODUCTS_CACHE_KEY = 'products_json';
var PRODUCTS_CACHE_TTL = 300; // giây — cache phía server để khỏi đọc Sheet mỗi request

function doGet(e) { return route_(e); }
function doPost(e) { return route_(e); }

function route_(e) {
  var action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action) : 'hit';
  if (action === 'products') return getProducts_();
  return handleHit_(e); // mặc định: đếm lượt truy cập
}

/* ----------------------- Đếm lượt truy cập ----------------------- */
function handleHit_(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // tránh hai lượt ghi đè lên nhau
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var counter = ss.getSheetByName(COUNTER_SHEET);
    if (!counter) {
      counter = ss.insertSheet(COUNTER_SHEET);
      counter.getRange('A1').setValue('Tổng lượt truy cập');
      counter.getRange('A2').setValue(0);
    }
    var total = Number(counter.getRange('A2').getValue()) || 0;
    total += 1;
    counter.getRange('A2').setValue(total);

    var log = ss.getSheetByName(LOG_SHEET);
    if (!log) {
      log = ss.insertSheet(LOG_SHEET);
      log.appendRow(['Thời gian', 'Trang', 'Tham chiếu']);
    }
    var page = (e && e.parameter && e.parameter.page) ? e.parameter.page : '';
    var ref = (e && e.parameter && e.parameter.ref) ? e.parameter.ref : '';
    log.appendRow([new Date(), page, ref]);

    return json_({ ok: true, total: total });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ----------------------- Dữ liệu sản phẩm ----------------------- */
function getProducts_() {
  try {
    var cache = CacheService.getScriptCache();
    var hit = cache.get(PRODUCTS_CACHE_KEY);
    if (hit) {
      return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(PRODUCTS_SHEET);
    if (!sh) return json_({ ok: false, error: 'Chưa có sheet "products". Hãy chạy seedProducts().' });

    var values = sh.getDataRange().getValues();
    if (values.length < 2) return json_({ ok: true, products: [] });

    var header = values[0].map(function (h) { return String(h).trim(); });
    var products = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (!row[0] && !row[1]) continue; // bỏ hàng trống
      var obj = {};
      for (var c = 0; c < header.length; c++) {
        if (header[c]) obj[header[c]] = row[c];
      }
      // chuẩn hóa kiểu dữ liệu
      obj.price = Number(obj.price) || 0;
      obj.originalPrice = Number(obj.originalPrice) || 0;
      obj.rating = Number(obj.rating) || 0;
      obj.order = Number(obj.order) || 0;
      obj.active = (obj.active === true || String(obj.active).toUpperCase() === 'TRUE');
      products.push(obj);
    }

    var payload = JSON.stringify({ ok: true, products: products, cachedAt: new Date().toISOString() });
    cache.put(PRODUCTS_CACHE_KEY, payload, PRODUCTS_CACHE_TTL);
    return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/**
 * Xoá cache sản phẩm phía server (gọi sau khi sửa Sheet nếu muốn cập nhật ngay,
 * thay vì chờ TTL hết hạn). Có thể chạy thủ công trong editor.
 */
function clearProductsCache() {
  CacheService.getScriptCache().remove(PRODUCTS_CACHE_KEY);
}

/**
 * Trigger đơn giản: tự xoá cache khi sheet "products" bị chỉnh sửa, để dữ liệu
 * mới hiện ra gần như tức thì (không phải chờ hết TTL 300s).
 * Chạy tự động sau khi lưu code — KHÔNG cần deploy lại.
 */
function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === PRODUCTS_SHEET) {
      clearProductsCache();
    }
  } catch (err) { /* bỏ qua */ }
}

/* ----------------------- Seed dữ liệu ban đầu ----------------------- */
/**
 * Tạo sheet "products" (nếu chưa có) và ghi header + 12 sản phẩm ban đầu.
 * Chạy MỘT LẦN trong editor. Chạy lại sẽ ghi đè toàn bộ dữ liệu hiện có.
 */
function seedProducts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PRODUCTS_SHEET);
  if (!sh) sh = ss.insertSheet(PRODUCTS_SHEET);
  sh.clear();

  var header = ['id', 'name', 'price', 'originalPrice', 'rating', 'sold', 'ageRange', 'steam', 'guideNote', 'link', 'order', 'active'];
  var rows = [
    ["p1", "Bộ Đất Nặn Hình Học Sáng Tạo", 65000, 89000, 4.6, "4,1k", "3–5", "Nghệ thuật,Toán học", "Có hướng dẫn Tiếng Việt", "guide.html", 1, true],
    ["p2", "Bộ Vẽ Sáng Tạo Màu Nước", 89000, 120000, 4.6, "3,4k", "3–5", "Nghệ thuật", "Có hướng dẫn Tiếng Việt", "guide.html", 2, true],
    ["p3", "Khối Gỗ Toán Học Đếm Số", 75000, 99000, 4.8, "2,7k", "3–5", "Toán học", "Có hướng dẫn Tiếng Việt", "guide.html", 3, true],
    ["p4", "Bộ Xếp Hình Bánh Răng Kỹ Thuật", 145000, 189000, 4.7, "2,2k", "6–8", "Kỹ nghệ,Toán học", "Có hướng dẫn Tiếng Việt", "guide.html", 4, true],
    ["p5", "Robot Lập Trình Mạch Điện Mini", 189000, 259000, 4.8, "2,1k", "6–8", "Khoa học,Công nghệ,Kỹ nghệ,Toán học", "Có hướng dẫn Tiếng Việt", "guide.html", 5, true],
    ["p6", "Bộ Thí Nghiệm Núi Lửa Phun Trào", 120000, 165000, 4.7, "1,9k", "6–8", "Khoa học", "Có hướng dẫn Tiếng Việt", "guide.html", 6, true],
    ["p7", "Mô Hình Hệ Mặt Trời Lắp Ráp", 159000, 210000, 4.7, "1,3k", "6–8", "Khoa học,Kỹ nghệ", "Có hướng dẫn Tiếng Việt", "guide.html", 7, true],
    ["p8", "Xe Đua Năng Lượng Mặt Trời", 245000, 310000, 4.8, "1,1k", "9–12", "Khoa học,Kỹ nghệ,Công nghệ", "Có hướng dẫn Tiếng Việt", "guide.html", 8, true],
    ["p9", "Bộ Kính Hiển Vi Khám Phá Vi Sinh", 250000, 320000, 4.9, "860", "9–12", "Khoa học", "Có hướng dẫn Tiếng Việt", "guide.html", 9, true],
    ["p10", "Đàn Phím Điện Tử Mini", 199000, 259000, 4.5, "770", "3–5", "Nghệ thuật,Công nghệ", "Có hướng dẫn Tiếng Việt", "guide.html", 10, true],
    ["p11", "Bảng Mạch Điện Tử Vui Nhộn", 280000, 350000, 4.9, "690", "9–12", "Công nghệ,Kỹ nghệ", "Có hướng dẫn Tiếng Việt", "guide.html", 11, true],
    ["p12", "Cánh Tay Robot Thủy Lực", 320000, 399000, 4.9, "540", "9–12", "Kỹ nghệ,Công nghệ", "Có hướng dẫn Tiếng Việt", "guide.html", 12, true]
  ];

  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sh.getRange(2, 1, rows.length, header.length).setValues(rows);
  sh.setFrozenRows(1);
  clearProductsCache();
}

/* ----------------------- Tiện ích ----------------------- */
function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
