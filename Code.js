// Hàm hiển thị giao diện Web App
function doGet() {
  // Tạo template từ file Index
  var template = HtmlService.createTemplateFromFile('index');
  
  // Tự động lấy URL của Web App gán vào biến 'url'
  template.url = ScriptApp.getService().getUrl(); 
  
  return template.evaluate()
    .setTitle('Phiếu Điều Tra Thông Tin Hành Chính')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Hàm tiếp nhận dữ liệu từ Form gửi lên và tách ghi vào 2 Sheet
function doPost(e) {
  try {
    // Đọc dữ liệu JSON gửi từ trình duyệt
    var data = JSON.parse(e.postData.contents);
    
    // LIÊN KẾT TRỰC TIẾP TỚI FILE GOOGLE SHEET QUA ID
    var spreadsheetId = "1gj70N3TTJUvAZxU_C0f_TN3HxTuwBCw6r80it2g1nQM";
    var ss = SpreadsheetApp.openById(spreadsheetId);
    
    // ---------------------------------------------------------------
    // XỬ LÝ SHEET 1: Thông tin hộ gia đình
    // ---------------------------------------------------------------
    var sheet1 = ss.getSheetByName("Thông tin hộ gia đình");
    if (!sheet1) {
      sheet1 = ss.insertSheet("Thông tin hộ gia đình");
    }
    // Nếu trang tính chưa có dữ liệu, tự động tạo hàng tiêu đề (Header)
    if (sheet1.getLastRow() === 0) {
      sheet1.appendRow(["Mã Hộ", "Chủ Hộ", "Địa Chỉ", "SĐT Hộ", "Tổng Số Nhân Khẩu"]);
    }
    
    // Tự động tính Mã hộ tăng dần từ 1 dựa trên số dòng hiện tại của Sheet 1
    var maHoTuDong = sheet1.getLastRow();
    
    // Ghi thông tin Hộ gia đình vào Sheet 1
    sheet1.appendRow([
      maHoTuDong,
      data.chuHo,
      data.diaChi,
      "'" + data.sdtHo, // Thêm dấu nháy đơn để giữ định dạng chuỗi số điện thoại
      data.tongNhanKhau
    ]);
    
    // ---------------------------------------------------------------
    // XỬ LÝ SHEET 2: Thông tin thành viên trong hộ
    // ---------------------------------------------------------------
    var sheet2 = ss.getSheetByName("Thông tin thành viên trong hộ");
    if (!sheet2) {
      sheet2 = ss.insertSheet("Thông tin thành viên trong hộ");
    }
    // Nếu trang tính chưa có dữ liệu, tự động tạo hàng tiêu đề (Header)
    if (sheet2.getLastRow() === 0) {
      sheet2.appendRow([
        "Mã Hộ", "STT Thành Viên", "Họ và Tên", "Hộ khẩu", "Giới Tính", "Ngày Sinh", 
        "CCCD/Định Danh", "SĐT Thành Viên", "Nghề Nghiệp/Nơi Làm Việc", 
        "Nhu Cầu Khám Sức Khỏe", "Nhu Cầu Khám Sàng Lọc", "Nhóm Đối Tượng", "Đăng Ký Khám Tại Nhà"
      ]);
    }
    
    // Lặp qua danh sách thành viên trong hộ để ghi vào từng dòng ở Sheet 2
    var members = data.members;
    for (var i = 0; i < members.length; i++) {
      var m = members[i];
      sheet2.appendRow([
        maHoTuDong, // Cột đầu tiên là Mã hộ (Dùng chung mã hộ vừa tạo ở Sheet 1)
        m.stt,
        m.hoTen,
		m.hoKhau,
        m.gioiTinh == "1" ? "Nam" : "Nữ", // Chuyển đổi mã số sang text cho dễ đọc
        m.ngaySinh,
        "'" + m.cccd,       // Thêm dấu nháy đơn tránh mất số 0 ở đầu CCCD
        "'" + m.sdtCaNhan,  // Thêm dấu nháy đơn tránh mất số 0 ở đầu SĐT
        m.ngheNghiep,
        m.khamSucKhoe || "Không đăng ký",
        m.khamSangLoc,      // Chuỗi các bệnh lý cách nhau bằng dấu phẩy
        m.nhomDoiTuong,
        m.khamTaiNha
      ]);
    }
    
    // Trả về phản hồi thành công cho trình duyệt
    return ContentService.createTextOutput(JSON.stringify({ "status": "success", "message": "Ghi dữ liệu thành công!" }))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (error) {
    // Trả về lỗi nếu có sự cố xảy ra
    return ContentService.createTextOutput(JSON.stringify({ "status": "error", "message": error.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}