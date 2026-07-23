function doGet() {
  var template = HtmlService.createTemplateFromFile('index');
  template.url = ScriptApp.getService().getUrl();

  return template.evaluate()
    .setTitle('Phiếu Điều Tra Khám Sức Khỏe')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    var raw = (e && e.postData && e.postData.contents) || '{}';
    var data = JSON.parse(raw);

    validatePayload(data);

    var spreadsheetId = '1gj70N3TTJUvAZxU_C0f_TN3HxTuwBCw6r80it2g1nQM';
    var ss = SpreadsheetApp.openById(spreadsheetId);
    if (!ss) {
      throw new Error('Không thể mở file Google Sheet.');
    }

    var sheet = ss.getSheetByName('TDP');
    if (!sheet) {
      sheet = ss.insertSheet('TDP');
    }

    setupSheetLayout(sheet, data);

    var members = Array.isArray(data.members) ? data.members : [];
    if (members.length === 0) {
      return jsonOutput({
        status: 'success',
        message: 'Không có thành viên nào để ghi.'
      });
    }

    var totalRow = findTotalRow(sheet);
    if (totalRow < 0) {
      throw new Error('Không tìm thấy dòng "Tổng số".');
    }

    var insertAt = totalRow;
    sheet.insertRowsBefore(insertAt, members.length);

    var newRows = [];
    for (var i = 0; i < members.length; i++) {
      newRows.push(buildMemberRow(members[i] || {}, i + 1));
    }

    var dataRange = sheet.getRange(insertAt, 1, newRows.length, 36);
    dataRange.setValues(newRows);
    formatDataRows(dataRange);

    renumberSTT(sheet);
    updateTongSo(sheet);
    restoreFooterArea(sheet);

    SpreadsheetApp.flush();

    return jsonOutput({
      status: 'success',
      message: 'Ghi dữ liệu thành công!'
    });

  } catch (error) {
    return jsonOutput({
      status: 'error',
      message: error && error.message ? error.message : String(error)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (e2) {}
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function validatePayload(data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Dữ liệu không hợp lệ.');
  }

  var requiredFields = ['tenPhuong', 'toDanPho', 'toSo', 'chuHo', 'diaChi', 'sdtHo', 'tongNhanKhau'];
  for (var i = 0; i < requiredFields.length; i++) {
    var key = requiredFields[i];
    if (!String(data[key] || '').trim()) {
      throw new Error('Thiếu thông tin bắt buộc: ' + key);
    }
  }

  if (!Array.isArray(data.members) || data.members.length === 0) {
    throw new Error('Danh sách thành viên trống.');
  }

  var tongNhanKhau = Number(data.tongNhanKhau);
  if (!tongNhanKhau || tongNhanKhau < 1) {
    throw new Error('Tổng số nhân khẩu không hợp lệ.');
  }

  if (data.members.length !== tongNhanKhau) {
    throw new Error('Số thành viên không khớp với tổng số nhân khẩu.');
  }

  for (var j = 0; j < data.members.length; j++) {
    var m = data.members[j] || {};

    if (!String(m.hoTen || '').trim()) {
      throw new Error('Thiếu họ tên ở thành viên thứ ' + (j + 1));
    }
    if (!String(m.gioiTinh || '').trim()) {
      throw new Error('Thiếu giới tính ở thành viên thứ ' + (j + 1));
    }
    if (!String(m.ngaySinh || '').trim()) {
      throw new Error('Thiếu ngày sinh ở thành viên thứ ' + (j + 1));
    }
    if (!String(m.cuTru || '').trim()) {
      throw new Error('Thiếu thông tin cư trú ở thành viên thứ ' + (j + 1));
    }
    if (!String(m.nhomChinh || '').trim()) {
      throw new Error('Thiếu nhóm đối tượng ở thành viên thứ ' + (j + 1));
    }

    if (m.nhomChinh === 'nhom1' && !String(m.nhom1ChiTiet || '').trim()) {
      throw new Error('Thiếu chi tiết nhóm 1 ở thành viên thứ ' + (j + 1));
    }
    if (m.nhomChinh === 'nhom2' && !String(m.nhom2ChiTiet || '').trim()) {
      throw new Error('Thiếu chi tiết nhóm 2 ở thành viên thứ ' + (j + 1));
    }
    if (m.nhomChinh === 'nhom5' && !String(m.nhom5ChiTiet || '').trim()) {
      throw new Error('Thiếu chi tiết nhóm 5 ở thành viên thứ ' + (j + 1));
    }

    if (m.khamSucKhoeDinhKy === 'Không' && !String(m.khamSangLoc || '').trim()) {
      throw new Error('Phải chọn khám sàng lọc ở thành viên thứ ' + (j + 1));
    }

    if (m.khamSucKhoeDinhKy === 'Có' && String(m.khamSangLoc || '').trim()) {
      throw new Error('Không được chọn khám sàng lọc khi đã đăng ký khám định kỳ ở thành viên thứ ' + (j + 1));
    }
  }
}

function setupSheetLayout(sheet, data) {
  var numCols = 36;

  ensureMinRows(sheet, 20);
  ensureMinColumns(sheet, numCols);

  var totalRow = findTotalRow(sheet);
  if (sheet.getLastRow() > 0 && totalRow > 0) {
    restoreFooterArea(sheet);
    return;
  }

  try {
    sheet.clear();
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (e) {}

  var widths = {
    1: 30, 2: 140, 3: 40, 4: 35, 5: 35, 6: 80, 7: 90, 8: 70, 9: 55, 10: 80,
    11: 40, 12: 40, 13: 40, 14: 48, 15: 48, 16: 48, 17: 48, 18: 45, 19: 45,
    20: 40, 21: 40, 22: 40, 23: 45, 24: 45, 25: 45, 26: 45, 27: 45, 28: 45,
    29: 45, 30: 45, 31: 45, 32: 45, 33: 45, 34: 45, 35: 45, 36: 45
  };

  for (var c = 1; c <= numCols; c++) {
    sheet.setColumnWidth(c, widths[c] || 45);
  }

  sheet.getRange(1, 1, 20, numCols)
    .setFontFamily('Times New Roman')
    .setVerticalAlignment('middle');

  var tenPhuong = String(data.tenPhuong || 'UBND PHƯỜNG LONG BIÊN').trim();
  sheet.getRange('B1:F1').merge().setValue(tenPhuong)
    .setFontWeight('normal')
    .setHorizontalAlignment('center');

  var toDanPho = data.toDanPho ? 'TỔ DÂN PHỐ ' + data.toDanPho : 'TỔ DÂN PHỐ............................';
  sheet.getRange('B2:F2').merge().setValue(toDanPho)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sheet.getRange(3, 6, 1, 26).merge().setValue(
    'BIỂU THU THẬP THÔNG TIN NGƯỜI DÂN PHƯỜNG LONG BIÊN PHỤC VỤ CÔNG TÁC KHÁM SỨC KHỎE ĐỊNH KỲ, KHÁM SÀNG LỌC MIỄN PHÍ NĂM 2026'
  ).setFontWeight('bold')
   .setHorizontalAlignment('center')
   .setWrap(true)
   .setFontSize(13);

  var toSoText = data.toSo ? 'Tổ số: ' + data.toSo : 'Tổ số: ....................';
  sheet.getRange('B4:F4').merge().setValue(toSoText)
    .setHorizontalAlignment('left');

  buildComplexHeader(sheet);

  sheet.getRange(16, 1, 1, numCols).setBorder(true, true, true, true, true, true);
  sheet.getRange('A16:B16').merge().setValue('Tổng số')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  restoreFooterArea(sheet);

  sheet.getRange(5, 1, 11, numCols).setBorder(true, true, true, true, true, true);

  sheet.setRowHeight(1, 28);
  sheet.setRowHeight(2, 28);
  sheet.setRowHeight(3, 38);
  sheet.setRowHeight(4, 28);
  sheet.setRowHeight(5, 32);
  sheet.setRowHeight(6, 70);
  sheet.setRowHeight(7, 70);
  sheet.setRowHeight(8, 40);
}

function buildComplexHeader(sheet) {
  mergeAndSet(sheet, 'A5:A8', 'S\nT\nT');
  mergeAndSet(sheet, 'B5:B8', 'Họ và tên');
  mergeAndSet(sheet, 'C5:C8', 'Chủ\nhộ');
  mergeAndSet(sheet, 'D5:E6', 'Giới tính');
  mergeAndSet(sheet, 'D7:D8', 'Nam');
  mergeAndSet(sheet, 'E7:E8', 'Nữ');
  mergeAndSet(sheet, 'F5:F8', 'Ngày tháng\nnăm sinh');
  mergeAndSet(sheet, 'G5:G8', 'CCCD/\nSố định danh');

  mergeAndSet(sheet, 'H5:I6', 'Thông tin cư trú\n(theo VNeID)');
  mergeAndSet(sheet, 'H7:H8', 'Thường trú hoặc\ntạm trú từ\n12 tháng trở lên');
  mergeAndSet(sheet, 'I7:I8', 'Tạm trú\ndưới 12 tháng');

  mergeAndSet(sheet, 'J5:J8', 'Số điện thoại\n(đối với trẻ em\ndưới 18 tuổi\nghi SDT của\nngười giám hộ)');

  mergeAndSet(sheet, 'K5:W5', 'Đối tượng\n(Đánh dấu X vào đúng đối tượng, đối với nhóm 1 có thể lựa chọn nhiều phương án)');
  mergeAndSet(sheet, 'K6:M7', 'Nhóm 1 (người cao tuổi; người khuyết tật; hộ nghèo, cận nghèo; người mắc bệnh mạn tính)');
  mergeAndSet(sheet, 'K8:K8', '<6 tuổi');
  mergeAndSet(sheet, 'L8:L8', '6-18 tuổi');
  mergeAndSet(sheet, 'M8:M8', '>18 tuổi');

  mergeAndSet(sheet, 'N6:Q7', 'Nhóm 2 (học sinh từ mầm non đến THPT)');

  sheet.getRange('N7:O7').merge().setValue('Học ở phường\nLong Biên')
    .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('P7:Q7').merge().setValue('Học ở ngoài\nphường Long Biên')
    .setWrap(true).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.getRange('N8').setValue('<6 tuổi');
  sheet.getRange('O8').setValue('6-18\ntuổi');
  sheet.getRange('P8').setValue('<6 tuổi');
  sheet.getRange('Q8').setValue('6-18\ntuổi');

  mergeAndSet(sheet, 'R6:R8', 'Nhóm 3\n\nNgười lao động trong công ty,\ncơ quan, xí nghiệp,\nnhà máy...');
  mergeAndSet(sheet, 'S6:S8', 'Nhóm 4\n\nLực lượng\nvũ trang');
  mergeAndSet(sheet, 'T6:V7', 'Nhóm 5\n(những người không thuộc nhóm 1, 2, 3, 4 như lao động tự do, trẻ em không đi học...)');
  mergeAndSet(sheet, 'T8:T8', '<6 tuổi');
  mergeAndSet(sheet, 'U8:U8', '6-18 tuổi');
  mergeAndSet(sheet, 'V8:V8', '>18 tuổi');
  mergeAndSet(sheet, 'W6:W8', 'Từ tháng\n1/2026\ndến nay\nđã từng\nđược KSK\nmiễn phí');

  mergeAndSet(sheet, 'X5:AJ5', 'Đăng Ký Khám\n(Đánh dấu X vào ô tương ứng, mỗi người chỉ được chọn 1 loại khám)');
  mergeAndSet(sheet, 'X6:X8', 'Khám\nsức khỏe\nđịnh kỳ');
  mergeAndSet(sheet, 'Y6:AJ6', 'Khám sàng lọc');

  var screeningHeaders = [
    'Tăng\nhuyết\náp',
    'Đái\ntháo\nđường\ntíp 2',
    'Hen\nphế\nquản',
    'Phổi tắc\nnghẽn\nmạn tính',
    'Ung\nthư\nvú',
    'Ung\nthư\ncổ tử\ncung',
    'Ung thư\nkhoang\nmiệng',
    'Ung\nthư đại\ntrực\ntràng',
    'Ung thư\ntuyến\ntiền liệt',
    'Rối\nloạn\ntrầm\ncảm',
    'Rối\nloạn\nlo âu',
    'Rối\nloạn\ntâm thần\ndo rượu'
  ];

  for (var i = 0; i < screeningHeaders.length; i++) {
    sheet.getRange(7, 25 + i, 2, 1).merge().setValue(screeningHeaders[i])
      .setWrap(true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
  }

  sheet.getRange(5, 1, 4, 36)
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
}

function mergeAndSet(sheet, a1, value) {
  var range = sheet.getRange(a1);
  range.merge();
  range.setValue(value);
  range.setWrap(true);
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
}

function buildMemberRow(m, stt) {
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
    "'" + (m.sdtCaNhan || ''),
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

function formatDataRows(range) {
  range
    .setFontFamily('Times New Roman')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(true, true, true, true, true, true);

  range.offset(0, 1, range.getNumRows(), 1).setHorizontalAlignment('left');
}

function renumberSTT(sheet) {
  var totalRow = findTotalRow(sheet);
  if (totalRow < 0 || totalRow <= 9) return;

  var stt = 1;
  for (var row = 9; row < totalRow; row++) {
    var name = sheet.getRange(row, 2).getValue();
    if (name !== '') {
      sheet.getRange(row, 1).setValue(stt);
      stt++;
    } else {
      sheet.getRange(row, 1).setValue('');
    }
  }
}

function updateTongSo(sheet) {
  var totalRow = findTotalRow(sheet);
  if (totalRow < 0) return;

  var lastDataRow = totalRow - 1;
  if (lastDataRow < 9) {
    sheet.getRange(totalRow, 2).setValue('0');
  } else {
    sheet.getRange(totalRow, 2).setFormula('=COUNTA(B9:B' + lastDataRow + ')');
  }
}

function findTotalRow(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return -1;

  var values = sheet.getRange(1, 1, lastRow, 2).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var colA = String(values[i][0] || '').trim();
    var colB = String(values[i][1] || '').trim();
    if (colA === 'Tổng số' || colB === 'Tổng số') {
      return i + 1;
    }
  }
  return -1;
}

function restoreFooterArea(sheet) {
  var totalRow = findTotalRow(sheet);
  if (totalRow < 0) return;

  var footerStart = totalRow + 2;
  ensureMinRows(sheet, footerStart + 4);
  ensureMinColumns(sheet, 36);

  sheet.getRange(footerStart, 23, 1, 5).merge().setValue('Ngày     tháng     năm 2026')
    .setHorizontalAlignment('center')
    .setFontFamily('Times New Roman')
    .setFontStyle('italic');

  sheet.getRange(footerStart + 1, 12, 1, 4).merge().setValue('Người điều tra')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontFamily('Times New Roman');

  sheet.getRange(footerStart + 1, 23, 1, 7).merge().setValue('Người nhận phiếu điều tra')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setFontFamily('Times New Roman');
}

function ensureMinRows(sheet, minRows) {
  var current = sheet.getMaxRows();
  if (current < minRows) {
    sheet.insertRowsAfter(current, minRows - current);
  }
}

function ensureMinColumns(sheet, minCols) {
  var current = sheet.getMaxColumns();
  if (current < minCols) {
    sheet.insertColumnsAfter(current, minCols - current);
  }
}

function formatDateVN(dateStr) {
  if (!dateStr) return '';
  var d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  var day = d.getDate();
  var month = d.getMonth() + 1;
  var year = d.getFullYear();
  return day + '/' + month + '/' + year;
}
