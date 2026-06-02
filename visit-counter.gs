/**
 * STEAMLab — Bộ đếm lượt truy cập (Google Apps Script backend)
 * =============================================================
 *
 * Script này đóng vai trò "server nhỏ": mỗi lần trang web gọi tới, nó tăng
 * số lượt truy cập lưu trong Google Sheet và trả về tổng số dưới dạng JSON.
 *
 * CÁCH CÀI ĐẶT
 * ------------
 * 1. Tạo một Google Sheet mới tại https://sheets.google.com
 * 2. Trong Sheet, mở menu  Tiện ích mở rộng (Extensions) → Apps Script.
 * 3. Xoá hết code mẫu, dán TOÀN BỘ nội dung file này vào.
 * 4. Bấm  Triển khai (Deploy) → Tùy chọn triển khai mới (New deployment).
 *      - Loại (Type):            Ứng dụng web (Web app)
 *      - Thực thi với tư cách:   Tôi (Me)
 *      - Ai có quyền truy cập:   Bất kỳ ai (Anyone)
 *    → Bấm Triển khai, cấp quyền khi được hỏi.
 * 5. Sao chép  "URL ứng dụng web"  (dạng https://script.google.com/macros/s/XXXX/exec)
 * 6. Dán URL đó vào hằng số  COUNTER_API  trong index.html và guide.html.
 *
 * Mỗi lượt truy cập được ghi thêm 1 dòng (thời gian + trang) vào sheet "log",
 * và tổng số được lưu ở ô A2 của sheet "counter" — bạn xem trực tiếp trong file.
 */

// Tên hai sheet dùng để lưu dữ liệu (tự tạo nếu chưa có).
var COUNTER_SHEET = 'counter';
var LOG_SHEET = 'log';

function doGet(e) {
  return handleHit_(e);
}

function doPost(e) {
  return handleHit_(e);
}

function handleHit_(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000); // tránh hai lượt ghi đè lên nhau
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // --- sheet đếm tổng ---
    var counter = ss.getSheetByName(COUNTER_SHEET);
    if (!counter) {
      counter = ss.insertSheet(COUNTER_SHEET);
      counter.getRange('A1').setValue('Tổng lượt truy cập');
      counter.getRange('A2').setValue(0);
    }
    var total = Number(counter.getRange('A2').getValue()) || 0;
    total += 1;
    counter.getRange('A2').setValue(total);

    // --- sheet log chi tiết ---
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

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
