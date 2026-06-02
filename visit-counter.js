/**
 * STEAMLab — Badge đếm lượt truy cập (tách khỏi HTML để bật CSP script-src 'self').
 * Gọi Apps Script Web App rồi hiển thị tổng lượt; mọi lỗi đều ẩn badge.
 */
(function () {
  // Dán URL "Ứng dụng web" của Apps Script vào đây (xem hướng dẫn trong backend.gs)
  var COUNTER_API = 'https://script.google.com/macros/s/AKfycbwhfnMbzfO25rnW8H2yKqHT6cwiYNzi9LL-p3ng-71IWkTsDb_9DEeNiBNrngF_oREPwQ/exec';
  var elCount = document.getElementById('visit-count');
  var elBox = document.getElementById('visit-counter');

  if (!COUNTER_API) {
    if (elBox) elBox.style.display = 'none'; // chưa cấu hình thì ẩn badge
    console.warn('[visit-counter] Chưa đặt COUNTER_API — xem hướng dẫn trong backend.gs');
    return;
  }

  var url = COUNTER_API
    + (COUNTER_API.indexOf('?') === -1 ? '?' : '&')
    + 'page=' + encodeURIComponent(location.pathname)
    + '&ref=' + encodeURIComponent(document.referrer || '');

  fetch(url, { method: 'GET' })
    .then(function (r) { return r.json(); })
    .then(function (data) {
      if (data && data.ok && elCount) {
        elCount.textContent = Number(data.total).toLocaleString('vi-VN');
      } else if (elBox) {
        elBox.style.display = 'none';
      }
    })
    .catch(function () { if (elBox) elBox.style.display = 'none'; });
})();
