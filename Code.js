// Hàm hiển thị giao diện Web App
function doGet() {
  var template = HtmlService.createTemplateFromFile('Index');
  template.url = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('Phiếu Điều Tra Thông Tin Hành Chính')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Hàm tiếp nhận dữ liệu từ Form gửi lên và tách ghi vào 2 Sheet
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');

    var spreadsheetId = '1gj70N3TTJUvAZxU_C0f_TN3HxTuwBCw6r80it2g1nQM';
    var ss = SpreadsheetApp.openById(spreadsheetId);

    var sheet1 = ss.getSheetByName('Thông tin hộ gia đình');
    if (!sheet1) {
      sheet1 = ss.insertSheet('Thông tin hộ gia đình');
    }

    if (sheet1.getLastRow() === 0) {
      sheet1.appendRow(['Mã Hộ', 'Chủ Hộ', 'Địa Chỉ', 'Khu Vực', 'SĐT Hộ', 'Tổng Số Nhân Khẩu']);
    }

    var maHoTuDong = "HO_" + new Date().getTime();

    sheet1.appendRow([
      "'" + maHoTuDong,
      data.chuHo || '',
      data.diaChi || '',
      data.khuVuc || '',
      "'" + (data.sdtHo || ''),
      data.tongNhanKhau || 0
    ]);

    var sheet2 = ss.getSheetByName('Thông tin thành viên trong hộ');
    if (!sheet2) {
      sheet2 = ss.insertSheet('Thông tin thành viên trong hộ');
    }

    if (sheet2.getLastRow() === 0) {
      sheet2.appendRow([
        'Mã Hộ', 'STT Thành Viên', 'Họ và Tên', 'Hộ khẩu', 'Thời gian tạm trú', 'Giới Tính', 'Ngày Sinh',
        'CCCD/Định Danh', 'SĐT Thành Viên', 'Nghề Nghiệp/Nơi Làm Việc',
        'Nhu Cầu Khám Sức Khỏe', 'Nhu Cầu Khám Sàng Lọc', 'Nhóm Đối Tượng', 'Đăng Ký Khám Tại Nhà'
      ]);
    }

    var members = Array.isArray(data.members) ? data.members : [];
    var memberRows = [];

    for (var i = 0; i < members.length; i++) {
      var m = members[i] || {};
      memberRows.push([
        "'" + maHoHoTuDong,
        m.stt || (i + 1),
        m.hoTen || '',
        m.hoKhau || '',
        m.thoiGianTamTru || '',
        m.gioiTinh || '',
        m.ngaySinh || '',
        "'" + (m.cccd || ''),
        "'" + (m.sdtCaNhan || ''),
        m.ngheNghiep || '',
        m.khamSucKhoe || 'Không đăng ký',
        m.khamSangLoc || '',
        m.nhomDoiTuong || '',
        m.khamTaiNha || ''
      ]);
    }

    if (memberRows.length > 0) {
      sheet2.getRange(sheet2.getLastRow() + 1, 1, memberRows.length, memberRows[0].length).setValues(memberRows);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Ghi dữ liệu thành công!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
