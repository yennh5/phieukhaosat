// Hàm hiển thị giao diện Web App
function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  template.url = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('Phiếu Điều Tra Thông Tin Hành Chính')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Hàm tiếp nhận dữ liệu từ Form gửi lên và ghi vào sheet TDP
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents || '{}');

    var spreadsheetId = '1gj70N3TTJUvAZxU_C0f_TN3HxTuwBCw6r80it2g1nQM';
    var ss = SpreadsheetApp.openById(spreadsheetId);

    var sheet = ss.getSheetByName('TDP');
    if (!sheet) {
      sheet = ss.insertSheet('TDP');
    }

    var HEADERS = [
      'STT',
      'Họ và tên',
      'Chủ hộ',
      'Nam',
      'Nữ',
      'Ngày tháng năm sinh',
      'CCCD/Số định danh',
      'Thường trú hoặc tạm trú từ 12 tháng trở lên',
      'Tạm trú dưới 12 tháng',
      'Số điện thoại',
      'Nhóm 1 - <6 tuổi',
      'Nhóm 1 - 6-18 tuổi',
      'Nhóm 1 - >18 tuổi',
      'Nhóm 2 - Học ở phường Long Biên <6 tuổi',
      'Nhóm 2 - Học ở phường Long Biên 6-18 tuổi',
      'Nhóm 2 - Học ở ngoài phường Long Biên <6 tuổi',
      'Nhóm 2 - Học ở ngoài phường Long Biên 6-18 tuổi',
      'Nhóm 3',
      'Nhóm 4',
      'Nhóm 5 - <6 tuổi',
      'Nhóm 5 - 6-18 tuổi',
      'Nhóm 5 - >18 tuổi',
      'Từ tháng 1/2026 đến nay đã từng được KSK miễn phí',
      'Khám sức khỏe định kỳ',
      'Tăng huyết áp',
      'Đái tháo đường típ 2',
      'Hen phế quản',
      'Phổi tắc nghẽn mạn tính',
      'Ung thư vú',
      'Ung thư cổ tử cung',
      'Ung thư khoang miệng',
      'Ung thư đại trực tràng',
      'Ung thư tuyến tiền liệt',
      'Rối loạn trầm cảm',
      'Rối loạn lo âu',
      'Rối loạn tâm thần do rượu'
    ];
    var NUM_COLS = HEADERS.length; // 36 cột (A:AJ)

    // Khởi tạo cấu trúc sheet nếu trống
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, NUM_COLS)
        .setValues([HEADERS])
        .setFontWeight('bold')
        .setBorder(true, true, true, true, true, true);
      // Tạo dòng "Tổng số" ban đầu
      var tongSoInit = [];
      for (var k = 0; k < NUM_COLS; k++) tongSoInit.push('');
      tongSoInit[0] = 'Tổng số';
      tongSoInit[1] = '0';
      sheet.getRange(2, 1, 1, NUM_COLS).setValues([tongSoInit]).setFontWeight('bold');
      // Tạo dòng chữ ký/footer
      var footerRow = [];
      for (var k = 0; k < NUM_COLS; k++) footerRow.push('');
      footerRow[0] = 'Người lập biểu';
      footerRow[NUM_COLS - 1] = 'Xác nhận của Trưởng thôn/TDP';
      sheet.getRange(3, 1, 1, NUM_COLS).setValues([footerRow]);
    }

    var members = Array.isArray(data.members) ? data.members : [];
    if (members.length === 0) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: 'success', message: 'Không có thành viên nào để ghi.' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Tìm vị trí dòng "Tổng số" (chỉ số 1-based trong sheet)
    var allColA = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    var tongSoSheetRow = -1;
    for (var i = 1; i < allColA.length; i++) {
      if (allColA[i][0] === 'Tổng số') {
        tongSoSheetRow = i + 1; // chuyển sang 1-based
        break;
      }
    }

    // Vị trí chèn dòng mới: trước dòng "Tổng số", hoặc cuối sheet nếu không tìm thấy
    var insertAt;
    if (tongSoSheetRow > 0) {
      insertAt = tongSoSheetRow;
    } else {
      insertAt = sheet.getLastRow() + 1;
    }

    // Xây dựng các dòng dữ liệu thành viên mới
    var newRows = [];
    for (var i = 0; i < members.length; i++) {
      newRows.push(buildMemberRow(members[i] || {}, data.sdtHo || '', 0));
    }

    // Chèn dòng mới vào sheet
    if (tongSoSheetRow > 0) {
      sheet.insertRowsBefore(insertAt, newRows.length);
    }
    var dataRange = sheet.getRange(insertAt, 1, newRows.length, NUM_COLS);
    dataRange.setValues(newRows);
    dataRange.setBorder(true, true, true, true, true, true);

    // Đánh lại số thứ tự STT toàn sheet
    renumberSTT(sheet);

    // Cập nhật công thức dòng "Tổng số"
    updateTongSo(sheet, NUM_COLS);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', message: 'Ghi dữ liệu thành công!' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Xây dựng mảng giá trị cho một dòng thành viên
function buildMemberRow(m, sdtHoFallback, stt) {
  return [
    stt,
    m.hoTen || '',
    m.laChuHo === 'Có' ? 'x' : '',
    m.gioiTinh === 'Nam' ? 'x' : '',
    m.gioiTinh === 'Nữ' ? 'x' : '',
    formatDateVN(m.ngaySinh),
    "'" + (m.cccd || ''),
    m.cuTru === 'thuong_tru_hoac_tam_tru_12_thang' ? 'x' : '',
    m.cuTru === 'tam_tru_duoi_12_thang' ? 'x' : '',
    "'" + (m.sdtCaNhan || sdtHoFallback || ''),
    m.nhomChinh === 'nhom1' && m.nhom1ChiTiet === 'duoi_6_tuoi' ? 'x' : '',
    m.nhomChinh === 'nhom1' && m.nhom1ChiTiet === '6_18_tuoi' ? 'x' : '',
    m.nhomChinh === 'nhom1' && m.nhom1ChiTiet === 'tren_18_tuoi' ? 'x' : '',
    m.nhomChinh === 'nhom2' && m.nhom2ChiTiet === 'hoc_trong_phuong_duoi_6' ? 'x' : '',
    m.nhomChinh === 'nhom2' && m.nhom2ChiTiet === 'hoc_trong_phuong_6_18' ? 'x' : '',
    m.nhomChinh === 'nhom2' && m.nhom2ChiTiet === 'hoc_ngoai_phuong_duoi_6' ? 'x' : '',
    m.nhomChinh === 'nhom2' && m.nhom2ChiTiet === 'hoc_ngoai_phuong_6_18' ? 'x' : '',
    m.nhomChinh === 'nhom3' ? 'x' : '',
    m.nhomChinh === 'nhom4' ? 'x' : '',
    m.nhomChinh === 'nhom5' && m.nhom5ChiTiet === 'duoi_6_tuoi' ? 'x' : '',
    m.nhomChinh === 'nhom5' && m.nhom5ChiTiet === '6_18_tuoi' ? 'x' : '',
    m.nhomChinh === 'nhom5' && m.nhom5ChiTiet === 'tren_18_tuoi' ? 'x' : '',
    m.daKskMienPhi === 'Có' ? 'x' : '',
    m.khamSucKhoeDinhKy === 'Có' ? 'x' : '',
    m.khamSangLoc === '1. Tăng huyết áp' ? 'x' : '',
    m.khamSangLoc === '2. Đái tháo đường típ 2' ? 'x' : '',
    m.khamSangLoc === '3. Hen phế quản' ? 'x' : '',
    m.khamSangLoc === '4. Phổi tắc nghẽn mạn tính' ? 'x' : '',
    m.khamSangLoc === '5. Ung thư vú' ? 'x' : '',
    m.khamSangLoc === '6. Ung thư cổ tử cung' ? 'x' : '',
    m.khamSangLoc === '7. Ung thư khoang miệng' ? 'x' : '',
    m.khamSangLoc === '8. Ung thư đại trực tràng' ? 'x' : '',
    m.khamSangLoc === '9. Ung thư tuyến tiền liệt' ? 'x' : '',
    m.khamSangLoc === '10. Rối loạn trầm cảm' ? 'x' : '',
    m.khamSangLoc === '11. Rối loạn lo âu' ? 'x' : '',
    m.khamSangLoc === '12. Rối loạn tâm thần do rượu' ? 'x' : ''
  ];
}

// Đánh lại số STT cho tất cả các dòng dữ liệu trong sheet (dựa trên cột B không rỗng)
function renumberSTT(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var stt = 1;
  for (var i = 0; i < values.length; i++) {
    var colA = String(values[i][0]);
    var colB = values[i][1];
    // Bỏ qua dòng "Tổng số", dòng footer và dòng trống
    if (colA === 'Tổng số') continue;
    if (colB === '' || colB === null || colB === undefined) continue;
    sheet.getRange(i + 2, 1).setValue(stt);
    stt++;
  }
}

// Cập nhật công thức đếm ở dòng "Tổng số" theo phạm vi dữ liệu thực tế
function updateTongSo(sheet, numCols) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  var colAValues = sheet.getRange(1, 1, lastRow, 1).getValues();
  var tongSoSheetRow = -1;
  for (var i = 1; i < colAValues.length; i++) {
    if (colAValues[i][0] === 'Tổng số') {
      tongSoSheetRow = i + 1; // 1-based
      break;
    }
  }

  if (tongSoSheetRow < 0) return;

  // Công thức đếm số ô có dữ liệu trong cột B từ dòng 2 đến dòng liền trước "Tổng số"
  var lastDataRow = tongSoSheetRow - 1;
  if (lastDataRow < 2) {
    sheet.getRange(tongSoSheetRow, 2).setValue('0');
  } else {
    sheet.getRange(tongSoSheetRow, 2).setFormula('=COUNTA(B2:B' + lastDataRow + ')');
  }
}

// Định dạng ngày tháng năm theo kiểu Việt Nam (d/m/yyyy)
function formatDateVN(dateStr) {
  if (!dateStr) return '';
  var d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  var day = d.getDate();
  var month = d.getMonth() + 1;
  var year = d.getFullYear();
  return day + '/' + month + '/' + year;
}
