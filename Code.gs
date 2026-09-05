/* ============================================================================
   HỆ THỐNG QUẢN LÝ BÁO CÁO HAK - Code.gs (BẢN HOÀN THIỆN v3)
   ============================================================================
   MỤC LỤC:
   1. CẤU HÌNH & KHỞI TẠO SHEET
   2. ĐĂNG NHẬP & PHÂN QUYỀN
   3. QUẢN LÝ TÀI KHOẢN
   4. DANH MỤC HỒ SƠ CON + KIỂM TRA ĐÚNG MẪU
   5. HẠN NỘP BÁO CÁO
   6. NỘP HỒ SƠ (UPLOAD) - chặn nộp trùng khi đang chờ duyệt + giới hạn cứng file
   7. QUY TRÌNH DUYỆT 2 CẤP: KTTH DUYỆT -> QUẢN LÝ PHÊ DUYỆT
   8. BÁO CÁO THIẾU HỒ SƠ (dựa trên cột "Là bản mới nhất")
   9. NHẮC HẠN NỘP TỰ ĐỘNG
   10. DASHBOARD TỔNG QUAN (dành cho Admin)
   11. MIGRATION - chạy 1 lần cho dữ liệu cũ
   ============================================================================

   CẤU TRÚC SHEET (Google Sheet ID = SHEET_DB_ID):

   [PHAN_QUYEN]  Email | Đơn vị | Quyền Báo Cáo | Vai trò | Trạng thái

   [DATA_FILE]   (20 cột)
    0 ID Giao dịch        1 Thời gian nộp     2 Đơn vị            3 Loại Báo Cáo
    4 Năm                 5 Kỳ Báo Cáo        6 Tên File          7 Link File
    8 Email người nộp     9 Mã Hồ Sơ          10 Trạng thái mẫu   11 Hạn nộp
    12 Trạng thái hạn     13 Trạng thái duyệt 14 KTTH duyệt(email) 15 Ngày KTTH duyệt
    16 Quản lý duyệt(email) 17 Ngày phê duyệt  18 Ghi chú duyệt   19 Là bản mới nhất

    Trạng thái duyệt: "Chờ KTTH duyệt" -> "Chờ Quản lý phê duyệt" -> "Đã phê duyệt"
                      (hoặc "KTTH từ chối..." / "Quản lý từ chối..." để nộp lại)
    Là bản mới nhất: "Có" / "Không" - mỗi tổ hợp (đơn vị+loại báo cáo+năm+kỳ+mã hồ
                     sơ) chỉ có DUY NHẤT 1 dòng "Có" tại 1 thời điểm, dùng để biết
                     dòng nào là bản hiện hành khi có nhiều lượt nộp lại.

   [DANH_MUC_HOSO]  Loại Báo Cáo | Mã Hồ Sơ | Tên Hồ Sơ | Regex Mẫu Tên File | Bắt buộc

   [KE_HOACH]  Đơn vị | Loại Báo Cáo | Ngày đến hạn (tháng kế tiếp) | Email cảnh báo
   ============================================================================ */

const ROOT_FOLDER_ID = '10ZLlISdP7mN3jTViMYlgHgcYig5yEIuP';
const SHEET_DB_ID     = '11NJxTJXpe9_YVoDGnXsJGxJEigZxPVbcmfodiLR25jU';

const TRANG_THAI_TU_CHOI = ["KTTH từ chối - Yêu cầu nộp lại", "Quản lý từ chối - Yêu cầu nộp lại"];
const TRANG_THAI_CHO_XU_LY = ["Chờ KTTH duyệt", "Chờ Quản lý phê duyệt"];

// Giới hạn cứng khi nộp file - vượt quá sẽ TỪ CHỐI hẳn, không cho nộp
const DINH_DANG_CHO_PHEP = ['.pdf', '.xls', '.xlsx', '.doc', '.docx', '.xml', '.csv', '.zip', '.jpg', '.jpeg', '.png'];
const GIOI_HAN_MB_MOI_FILE = 15;
const GIOI_HAN_MB_TONG_1_LUOT = 30;


/* ============================================================================
   1. CẤU HÌNH & KHỞI TẠO SHEET
   ============================================================================ */

function doGet() {
  kiemTraVaTaoSheetTuDong();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Hệ Thống Quản Lý Báo Cáo HAK')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function kiemTraVaTaoSheetTuDong() {
  var ss = SpreadsheetApp.openById(SHEET_DB_ID);

  var sheetPhanQuyen = ss.getSheetByName('PHAN_QUYEN');
  if (!sheetPhanQuyen) {
    sheetPhanQuyen = ss.insertSheet('PHAN_QUYEN');
    sheetPhanQuyen.appendRow(['Email', 'Đơn vị', 'Quyền Báo Cáo', 'Vai trò', 'Trạng thái']);
    sheetPhanQuyen.appendRow(['admin@gmail.com', 'ALL', 'ALL', 'Admin', 'Hoạt động']);
  }

  var sheetDataFile = ss.getSheetByName('DATA_FILE');
  if (!sheetDataFile) {
    sheetDataFile = ss.insertSheet('DATA_FILE');
    sheetDataFile.appendRow([
      'ID Giao dịch', 'Thời gian nộp', 'Đơn vị', 'Loại Báo Cáo', 'Năm', 'Kỳ Báo Cáo',
      'Tên File', 'Link File', 'Email người nộp', 'Mã Hồ Sơ', 'Trạng thái mẫu',
      'Hạn nộp', 'Trạng thái hạn', 'Trạng thái duyệt', 'KTTH duyệt (email)',
      'Ngày KTTH duyệt', 'Quản lý phê duyệt (email)', 'Ngày phê duyệt', 'Ghi chú duyệt',
      'Là bản mới nhất'
    ]);
    sheetDataFile.setFrozenRows(1);
  }
  // NẾU sheet DATA_FILE của bạn đã tồn tại từ trước (19 cột, chưa có "Là bản mới
  // nhất"), hãy tự thêm tiêu đề cột 20 = "Là bản mới nhất", sau đó chạy 1 lần hàm
  // capNhatLaBanMoiNhatChoDuLieuCu() ở mục 11 để gán giá trị cho dữ liệu cũ.

  taoSheetDanhMucHoSo();
  taoSheetKeHoach();
}

function _escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return str.toString()
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function _parseNgayVN(str) {
  if (!str) return null;
  var ngayGio = str.toString().split(" ");
  var dmy = ngayGio[0].split("/");
  if (dmy.length !== 3) return null;
  return new Date(Number(dmy[2]), Number(dmy[1]) - 1, Number(dmy[0]));
}

// Gọi 1 lần (thủ công) để tạo trigger chạy tự động mỗi ngày 7h sáng - mục 9.
function taoTriggerNhacHan() {
  var da_co = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'guiCanhBaoTruocHan';
  });
  if (da_co) return "Trigger đã tồn tại, không tạo thêm.";
  ScriptApp.newTrigger('guiCanhBaoTruocHan').timeBased().everyDays(1).atHour(7).create();
  return "✅ Đã tạo trigger nhắc hạn nộp, chạy hằng ngày lúc 7h sáng.";
}


/* ============================================================================
   2. ĐĂNG NHẬP & PHÂN QUYỀN
   ============================================================================ */

function layEmailHienHanhTuGoogle() {
  try {
    return Session.getActiveUser().getEmail() || "";
  } catch (e) {
    return "";
  }
}

function xacThucEmailThuCong(emailInput) {
  kiemTraVaTaoSheetTuDong();

  if (!emailInput || emailInput.trim() === "") {
    return { success: false, message: "⚠️ Vui lòng nhập địa chỉ Email!" };
  }

  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('PHAN_QUYEN');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][0].toString().toLowerCase().trim() === emailInput.toLowerCase().trim()) {
      var donVi = data[i][1] ? data[i][1].toString().trim() : "";
      var quyenBaoCao = data[i][2] ? data[i][2].toString().trim().toUpperCase() : "ALL";
      var vaiTro = data[i][3] ? data[i][3].toString().trim() : "Nhân viên";
      var trangThai = data[i][4] ? data[i][4].toString().trim() : "Hoạt động";

      if (trangThai === "Khóa") {
        return { success: false, message: "❌ <b>Tài khoản của bạn đã bị khóa!</b> Vui lòng liên hệ cấp quản lý để biết thêm chi tiết." };
      }
      return { success: true, email: emailInput, donVi: donVi, quyenBaoCao: quyenBaoCao, vaiTro: vaiTro };
    }
  }
  return { success: false, message: "❌ <b>Từ chối truy cập!</b> Email (<b>" + _escapeHtml(emailInput) + "</b>) chưa được cấp quyền trên hệ thống." };
}

function _layThongTinPhanQuyen(email) {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('PHAN_QUYEN');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][0].toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      return {
        donVi: data[i][1] ? data[i][1].toString().trim() : "",
        vaiTro: data[i][3] ? data[i][3].toString().trim() : "",
        trangThai: data[i][4] ? data[i][4].toString().trim() : "Hoạt động"
      };
    }
  }
  return null;
}


/* ============================================================================
   3. QUẢN LÝ TÀI KHOẢN
   ============================================================================ */

function capNhatHoacThemPhanQuyen(emailNguoiThucHien, email, donVi, loaiBaoCao, vaiTroMoi) {
  var nguoiThucHien = _layThongTinPhanQuyen(emailNguoiThucHien);
  if (!nguoiThucHien) throw new Error("⛔ Tài khoản thực hiện không hợp lệ.");

  var quyHoacGan = { "Admin": "Quản lý", "Quản lý": "KTTH", "KTTH": "Nhân viên" };
  if (quyHoacGan[nguoiThucHien.vaiTro] !== vaiTroMoi) {
    throw new Error("⛔ Bạn (" + nguoiThucHien.vaiTro + ") không có quyền cấp vai trò [" + vaiTroMoi + "].");
  }

  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('PHAN_QUYEN');
  var data = sheet.getDataRange().getValues();

  if (vaiTroMoi === "KTTH") {
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      if (data[i][1].toString().trim() === donVi && data[i][3].toString().trim() === "KTTH" &&
          data[i][0].toString().toLowerCase().trim() !== email.toLowerCase().trim()) {
        var status = data[i][4] ? data[i][4].toString().trim() : "Hoạt động";
        if (status === "Hoạt động") throw new Error("⛔ Nhà máy [" + donVi + "] đã có KTTH phụ trách đang hoạt động!");
      }
    }
  }

  var daTonTai = false;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][0].toString().toLowerCase().trim() === email.toLowerCase().trim()) {
      sheet.getRange(i + 1, 2).setValue(donVi);
      sheet.getRange(i + 1, 3).setValue(loaiBaoCao);
      sheet.getRange(i + 1, 4).setValue(vaiTroMoi);
      sheet.getRange(i + 1, 5).setValue("Hoạt động");
      daTonTai = true;
      break;
    }
  }
  if (!daTonTai) sheet.appendRow([email.trim().toLowerCase(), donVi, loaiBaoCao, vaiTroMoi, "Hoạt động"]);
  return "✅ Thao tác thành công cho tài khoản: " + email;
}

function thayDoiTrangThaiTaiKhoan(emailNguoiThucHien, emailMucTieu, trangThaiMoi) {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('PHAN_QUYEN');
  var data = sheet.getDataRange().getValues();

  var vaiTroHienTai = "";
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][0].toString().toLowerCase().trim() === emailNguoiThucHien.toLowerCase().trim()) {
      vaiTroHienTai = data[i][3].toString().trim();
      break;
    }
  }

  var targetRow = -1, targetVaiTro = "", targetDonVi = "";
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][0].toString().toLowerCase().trim() === emailMucTieu.toLowerCase().trim()) {
      targetRow = i + 1;
      targetVaiTro = data[i][3].toString().trim();
      targetDonVi = data[i][1].toString().trim();
      break;
    }
  }
  if (targetRow === -1) return "❌ Không tìm thấy tài khoản mục tiêu.";

  if (vaiTroHienTai === "Admin" && targetVaiTro !== "Quản lý") throw new Error("⛔ Lỗi phân cấp.");
  if (vaiTroHienTai === "Quản lý" && targetVaiTro !== "KTTH") throw new Error("⛔ Lỗi phân cấp.");
  if (vaiTroHienTai === "KTTH" && targetVaiTro !== "Nhân viên") throw new Error("⛔ Lỗi phân cấp.");

  sheet.getRange(targetRow, 5).setValue(trangThaiMoi);
  var arrDonViTarget = targetDonVi.split(",").map(function (item) { return item.trim(); });

  if (targetVaiTro === "Quản lý") {
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowVaiTro = data[i][3].toString().trim();
      var rowDonVi = data[i][1].toString().trim();
      if ((rowVaiTro === "KTTH" || rowVaiTro === "Nhân viên") && arrDonViTarget.includes(rowDonVi)) {
        sheet.getRange(i + 1, 5).setValue(trangThaiMoi);
      }
    }
  } else if (targetVaiTro === "KTTH") {
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var rowVaiTro = data[i][3].toString().trim();
      var rowDonVi = data[i][1].toString().trim();
      if (rowVaiTro === "Nhân viên" && rowDonVi === targetDonVi) {
        sheet.getRange(i + 1, 5).setValue(trangThaiMoi);
      }
    }
  }
  return "✅ Đã " + (trangThaiMoi === "Khóa" ? "KHÓA" : "MỞ KHÓA") + " thành công tài khoản: " + emailMucTieu + " và toàn bộ cấp dưới.";
}

function layDanhSachTaiKhoan(emailHienTai) {
  if (!emailHienTai) return "";
  emailHienTai = emailHienTai.toLowerCase().trim();

  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('PHAN_QUYEN');
  var data = sheet.getDataRange().getValues();

  var vaiTroXem = "", donViXem = "";
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][0].toString().toLowerCase().trim() === emailHienTai) {
      vaiTroXem = data[i][3].toString().trim();
      donViXem = data[i][1] ? data[i][1].toString().trim() : "";
      break;
    }
  }

  var html = "<table><tr><th>Email</th><th>Nhà máy / Đơn vị</th><th>Vai trò</th><th>Trạng thái</th><th>Thao tác</th></tr>";
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var emailVal = data[i][0];
    var donViVal = data[i][1];
    var vaiTroVal = data[i][3] ? data[i][3] : "Nhân viên";
    var trangThaiVal = data[i][4] ? data[i][4].toString().trim() : "Hoạt động";

    var hienThi = false;
    if (vaiTroXem === "Admin" && vaiTroVal === "Quản lý") hienThi = true;
    if (vaiTroXem === "Quản lý" && vaiTroVal === "KTTH" && donViXem.includes(donViVal)) hienThi = true;
    if (vaiTroXem === "KTTH" && vaiTroVal === "Nhân viên" && donViXem === donViVal) hienThi = true;

    if (hienThi) {
      var displayStatus = (trangThaiVal === "Khóa") ? "<span class='status-locked'>Bị khóa</span>" : "<span class='status-active'>Hoạt động</span>";
      var btnText = (trangThaiVal === "Khóa") ? "Mở Khóa" : "Khóa";
      var btnClass = (trangThaiVal === "Khóa") ? "btn-unlock" : "btn-lock";
      html += "<tr>";
      html += "<td>" + _escapeHtml(emailVal) + "</td>";
      html += "<td>" + _escapeHtml(donViVal) + "</td>";
      html += "<td><b>" + _escapeHtml(vaiTroVal) + "</b></td>";
      html += "<td>" + displayStatus + "</td>";
      html += "<td><button class='" + btnClass + "' onclick=\"doiTrangThaiTaiKhoan('" + _escapeHtml(emailVal) + "', '" + trangThaiVal + "')\">" + btnText + "</button></td>";
      html += "</tr>";
    }
  }
  html += "</table>";
  return html;
}

function getOrCreateFolder(parentFolder, folderName) {
  var folders = parentFolder.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parentFolder.createFolder(folderName);
}


/* ============================================================================
   4. DANH MỤC HỒ SƠ CON + KIỂM TRA ĐÚNG MẪU
   ============================================================================ */

function taoSheetDanhMucHoSo() {
  var ss = SpreadsheetApp.openById(SHEET_DB_ID);
  var sheet = ss.getSheetByName('DANH_MUC_HOSO');
  if (sheet) return;

  sheet = ss.insertSheet('DANH_MUC_HOSO');
  sheet.appendRow(['Loại Báo Cáo', 'Mã Hồ Sơ', 'Tên Hồ Sơ (hiển thị)', 'Regex Mẫu Tên File', 'Bắt buộc']);

  var seed = [
    ['THUE_GTGT', 'TK_GTGT',   'Tờ khai 01/GTGT (XML nộp eTax)',      '.*_GTGT_TT80-[MQ]\\d{2,6}-L\\d{2}\\.xml$', 'Y'],
    ['THUE_GTGT', 'BK_BANRA',  'Bảng kê hóa đơn bán ra',              '^01-GTGT-Bang-ke-ban-ra.*\\.xls[x]?$',     'Y'],
    ['THUE_GTGT', 'BK_MUAVAO', 'Bảng kê hóa đơn mua vào',             '^01-GTGT-Bang-ke-mua-vao.*\\.xls[x]?$',    'Y'],
    ['THUE_GTGT', 'GIAM_THUE', 'Bảng giảm thuế GTGT theo Nghị quyết', '^01-GTGT-Giam-thue-GTGT.*\\.xls[x]?$',     'N'],

    ['THUE_TNCN', 'TK_TNCN',   'Tờ khai 05/KK-TNCN (XML)',            '.*_KK_TNCN_TT80-[MQ]\\d{2,6}-L\\d{2}\\.xml$', 'Y'],
    ['THUE_TNCN', 'TH_TNCN',   'Bảng tổng hợp thu nhập chịu thuế',    '', 'Y'],

    ['BCTC_NB', 'CDKT',    'Bảng cân đối kế toán',           '', 'Y'],
    ['BCTC_NB', 'KQHDKD',  'Báo cáo kết quả HĐKD',           '', 'Y'],
    ['BCTC_NB', 'LCTT',    'Báo cáo lưu chuyển tiền tệ',     '', 'Y'],
    ['BCTC_NB', 'TMBCTC',  'Thuyết minh BCTC',               '', 'Y'],
    ['BCTC_NB', 'CN_NOIBO','Bảng đối chiếu công nợ nội bộ',  '', 'Y'],

    ['BCTC_THUE', 'CDKT_T',   'Bảng cân đối kế toán (nộp thuế)',       '', 'Y'],
    ['BCTC_THUE', 'KQHDKD_T', 'Báo cáo kết quả HĐKD (nộp thuế)',      '', 'Y'],
    ['BCTC_THUE', 'LCTT_T',   'Báo cáo lưu chuyển tiền tệ (nộp thuế)', '', 'Y'],
    ['BCTC_THUE', 'TMBCTC_T', 'Thuyết minh BCTC (nộp thuế)',          '', 'Y'],
    ['BCTC_THUE', 'TK_QT',    'Tờ khai quyết toán thuế TNDN (XML)',   '.*_TNDN.*\\.xml$', 'Y']
  ];
  seed.forEach(function (r) { sheet.appendRow(r); });
  sheet.setFrozenRows(1);
}

function layDanhSachHoSoTheoLoai(loaiBaoCao) {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DANH_MUC_HOSO');
  var data = sheet.getDataRange().getValues();
  var ketQua = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === loaiBaoCao) {
      ketQua.push({
        maHoSo: data[i][1].toString().trim(),
        tenHoSo: data[i][2].toString().trim(),
        batBuoc: data[i][4].toString().trim().toUpperCase() === 'Y'
      });
    }
  }
  return ketQua;
}

function kiemTraMauFile(tenFileGoc, loaiBaoCao, maHoSo) {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DANH_MUC_HOSO');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === loaiBaoCao && data[i][1].toString().trim() === maHoSo) {
      var regexStr = data[i][3].toString().trim();
      if (!regexStr) return 'Không kiểm tra';
      try {
        var re = new RegExp(regexStr, 'i');
        return re.test(tenFileGoc) ? 'Đúng mẫu' : 'Sai mẫu';
      } catch (e) { return 'Không kiểm tra'; }
    }
  }
  return 'Không kiểm tra';
}


/* ============================================================================
   5. HẠN NỘP BÁO CÁO
   ============================================================================ */

function taoSheetKeHoach() {
  var ss = SpreadsheetApp.openById(SHEET_DB_ID);
  var sheet = ss.getSheetByName('KE_HOACH');
  if (!sheet) {
    sheet = ss.insertSheet('KE_HOACH');
    sheet.appendRow(['Đơn vị', 'Loại Báo Cáo', 'Ngày đến hạn (của tháng kế tiếp)', 'Người nhận cảnh báo (Email, cách nhau bởi dấu phẩy)']);
  }
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) return;

  var dsDonVi = ['HAK_DN', 'CNHAK_QS', 'HAKQN_QC', 'DAIHIEP_DH'];
  var dsLoai = [['THUE_GTGT', 20], ['THUE_TNCN', 20], ['BCTC_NB', 15], ['BCTC_THUE', 30], ['LUONG', 5]];
  dsDonVi.forEach(function (dv) {
    dsLoai.forEach(function (lb) { sheet.appendRow([dv, lb[0], lb[1], '']); });
  });
}

function layCauHinhHanNop(donVi, loaiBaoCao) {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('KE_HOACH');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === donVi && data[i][1].toString().trim() === loaiBaoCao) {
      return {
        ngayDenHan: Number(data[i][2]),
        emailsCanhBao: data[i][3] ? data[i][3].toString().split(',').map(function (e) { return e.trim(); }).filter(Boolean) : []
      };
    }
  }
  return null;
}

function tinhHanNop(kyBaoCao, nam, ngayDenHan) {
  if (!ngayDenHan) return null;
  var namSo = Number(nam);
  var thangCuoiKy;

  var mThang = kyBaoCao.match(/Tháng\s*(\d+)/i);
  var mQuy = kyBaoCao.match(/Quý\s*(\d+)/i);
  if (mThang) thangCuoiKy = Number(mThang[1]);
  else if (mQuy) thangCuoiKy = Number(mQuy[1]) * 3;
  else return null;

  var thangHan = thangCuoiKy + 1;
  var namHan = namSo;
  if (thangHan > 12) { thangHan = 1; namHan += 1; }

  return new Date(namHan, thangHan - 1, ngayDenHan, 23, 59, 59);
}


/* ============================================================================
   6. NỘP HỒ SƠ (UPLOAD)
   ============================================================================ */

// Kiểm tra định dạng + dung lượng - CHẶN CỨNG nếu vi phạm
function _kiemTraFileHopLe(filesData) {
  var tongMB = 0;
  for (var i = 0; i < filesData.length; i++) {
    var f = filesData[i];
    var viTriDauCham = f.name.lastIndexOf(".");
    var ext = viTriDauCham !== -1 ? f.name.substring(viTriDauCham).toLowerCase() : "";
    if (DINH_DANG_CHO_PHEP.indexOf(ext) === -1) {
      return "⛔ Định dạng file <b>" + _escapeHtml(f.name) + "</b> không được phép. Chỉ chấp nhận: " + DINH_DANG_CHO_PHEP.join(", ");
    }
    var mb = (Number(f.size) || (f.data.length * 0.75)) / (1024 * 1024);
    if (mb > GIOI_HAN_MB_MOI_FILE) {
      return "⛔ File <b>" + _escapeHtml(f.name) + "</b> (" + mb.toFixed(1) + "MB) vượt quá giới hạn " + GIOI_HAN_MB_MOI_FILE + "MB/file.";
    }
    tongMB += mb;
  }
  if (tongMB > GIOI_HAN_MB_TONG_1_LUOT) {
    return "⛔ Tổng dung lượng " + tongMB.toFixed(1) + "MB vượt quá giới hạn " + GIOI_HAN_MB_TONG_1_LUOT + "MB/lượt nộp. Vui lòng chia nhỏ thành nhiều lượt.";
  }
  return null;
}

// Tìm trạng thái duyệt của dòng ĐANG LÀ BẢN MỚI NHẤT cho đúng tổ hợp
// (đơn vị + loại báo cáo + năm + kỳ [+ mã hồ sơ nếu có cấu hình danh mục]).
function _kiemTraTruocKhiNop(donVi, loaiBaoCao, nam, kyBaoCao, maHoSo, dbData) {
  for (var i = 1; i < dbData.length; i++) {
    if (dbData[i][2] === donVi && dbData[i][3] === loaiBaoCao &&
        dbData[i][4].toString() === nam.toString() && dbData[i][5] === kyBaoCao &&
        dbData[i][19] === "Có") {
      if (maHoSo && dbData[i][9] !== maHoSo) continue;

      var trangThaiCuoi = dbData[i][13];
      if (trangThaiCuoi === "Đã phê duyệt") {
        return "⛔ <b>TỪ CHỐI:</b> Hồ sơ " + loaiBaoCao + (maHoSo ? " - " + maHoSo : "") + " kỳ <b>" + kyBaoCao +
               "</b> đã được Quản lý phê duyệt, chốt số liệu. Không thể tải đè!";
      }
      if (TRANG_THAI_CHO_XU_LY.indexOf(trangThaiCuoi) !== -1) {
        return "⛔ <b>TỪ CHỐI:</b> Hồ sơ này đang ở trạng thái <b>" + trangThaiCuoi + "</b>, vui lòng chờ xử lý xong (hoặc bị từ chối) trước khi nộp lại.";
      }
    }
  }
  return null;
}

// Đánh dấu các dòng cũ cùng tổ hợp không còn là bản mới nhất nữa
function _danhDauKhongConMoiNhat(sheetDb, dbData, donVi, loaiBaoCao, nam, kyBaoCao, maHoSo) {
  for (var i = 1; i < dbData.length; i++) {
    if (dbData[i][2] === donVi && dbData[i][3] === loaiBaoCao &&
        dbData[i][4].toString() === nam.toString() && dbData[i][5] === kyBaoCao &&
        dbData[i][19] === "Có") {
      if (maHoSo && dbData[i][9] !== maHoSo) continue;
      sheetDb.getRange(i + 1, 20).setValue("Không");
    }
  }
}

function uploadMultipleFiles(emailNguoiGoi, donVi, loaiBaoCao, nam, kyBaoCao, filesData, maHoSo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var loiFile = _kiemTraFileHopLe(filesData);
    if (loiFile) return { success: false, message: loiFile };

    var sheetDb = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
    var dbData = sheetDb.getDataRange().getValues();

    var loiChan = _kiemTraTruocKhiNop(donVi, loaiBaoCao, nam, kyBaoCao, maHoSo, dbData);
    if (loiChan) return { success: false, message: loiChan };

    var sttDem = 1;
    for (var i = 1; i < dbData.length; i++) {
      if (dbData[i][2] === donVi && dbData[i][3] === loaiBaoCao && dbData[i][4].toString() === nam && dbData[i][5] === kyBaoCao) sttDem++;
    }

    _danhDauKhongConMoiNhat(sheetDb, dbData, donVi, loaiBaoCao, nam, kyBaoCao, maHoSo);

    var capHinh = layCauHinhHanNop(donVi, loaiBaoCao);
    var hanNopDate = capHinh ? tinhHanNop(kyBaoCao, nam, capHinh.ngayDenHan) : null;
    var hanNopStr = hanNopDate ? Utilities.formatDate(hanNopDate, "GMT+7", "dd/MM/yyyy") : "";
    var trangThaiHan = hanNopDate ? (new Date() <= hanNopDate ? "Đúng hạn" : "Trễ hạn") : "Không xác định";

    var root = DriveApp.getFolderById(ROOT_FOLDER_ID);
    var targetFolder = getOrCreateFolder(getOrCreateFolder(getOrCreateFolder(getOrCreateFolder(root, donVi), loaiBaoCao), nam), kyBaoCao);
    var kyBaoCaoClean = kyBaoCao.replace(/\s+/g, "_");
    var thoiGian = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss");
    var danhSachTenFile = "";
    var canhBaoSaiMau = [];

    for (var j = 0; j < filesData.length; j++) {
      var fileInfo = filesData[j];
      var blob = Utilities.newBlob(Utilities.base64Decode(fileInfo.data), fileInfo.mimeType, fileInfo.name);

      var originalName = fileInfo.name;
      var fileExtension = "", nameWithoutExt = originalName;
      if (originalName.lastIndexOf(".") !== -1) {
        fileExtension = originalName.substring(originalName.lastIndexOf("."));
        nameWithoutExt = originalName.substring(0, originalName.lastIndexOf("."));
      }

      var trangThaiMau = maHoSo ? kiemTraMauFile(originalName, loaiBaoCao, maHoSo) : 'Không kiểm tra';
      if (trangThaiMau === 'Sai mẫu') canhBaoSaiMau.push(originalName);

      var cleanOriginalName = nameWithoutExt.replace(/[\\/:*?"<>|]/g, "").replace(/\s+/g, "_");
      var sttFormatted = ("0" + sttDem).slice(-2);
      var standardizedFileName = donVi + "_" + loaiBaoCao + "_" + kyBaoCaoClean + "_" + nam + "_" + sttFormatted + "_" + cleanOriginalName + fileExtension;
      blob.setName(standardizedFileName);

      var newFile = targetFolder.createFile(blob);
      var idGiaoDich = "BC_" + Utilities.formatDate(new Date(), "GMT+7", "yyMMdd_HHmmss") + "_" + sttFormatted;

      sheetDb.appendRow([
        idGiaoDich, thoiGian, donVi, loaiBaoCao, nam, kyBaoCao, newFile.getName(), newFile.getUrl(),
        emailNguoiGoi || "", maHoSo || "", trangThaiMau, hanNopStr, trangThaiHan,
        "Chờ KTTH duyệt", "", "", "", "", "", "Có"
      ]);

      danhSachTenFile += "<li>" + standardizedFileName +
        (trangThaiMau === 'Sai mẫu' ? " <span style='color:#c0392b'>(⚠ Sai mẫu)</span>" : "") +
        (trangThaiHan === 'Trễ hạn' ? " <span style='color:#c0392b'>(⚠ Trễ hạn, hạn " + hanNopStr + ")</span>" : "") +
        "</li>";
      sttDem++;
    }

    if (emailNguoiGoi && emailNguoiGoi.trim() !== "") {
      var tieuDeMail = "[HAK] Biên nhận hồ sơ: Báo cáo " + loaiBaoCao + " - Đơn vị " + donVi;
      var noiDungMail = "<h3>XÁC NHẬN NỘP BÁO CÁO THÀNH CÔNG</h3>" +
        "<p>Xin chào <b>" + emailNguoiGoi + "</b>,</p>" +
        "<p>Hồ sơ đã được ghi nhận, đang ở trạng thái <b>Chờ KTTH duyệt</b>" +
        (hanNopStr ? (", hạn nộp: <b>" + hanNopStr + "</b>") : "") + ".</p>" +
        "<p><b>Danh sách tệp tin (" + filesData.length + " tệp):</b></p><ul>" + danhSachTenFile + "</ul>" +
        "<p>Trân trọng,<br><b>Hệ Thống Quản Trị HAK</b></p>";
      try { MailApp.sendEmail({ to: emailNguoiGoi, subject: tieuDeMail, htmlBody: noiDungMail }); } catch (e) { console.log(e); }
    }

    var msg = "✅ Đã tải thành công <b>" + filesData.length + "</b> tệp tin, chuyển sang trạng thái <b>Chờ KTTH duyệt</b>.";
    if (trangThaiHan === 'Trễ hạn') msg += "<br>⚠️ Lưu ý: đã nộp <b>trễ hạn</b> (hạn nộp " + hanNopStr + ").";
    if (canhBaoSaiMau.length > 0) msg += "<br>⚠️ " + canhBaoSaiMau.length + " tệp chưa khớp mẫu quy định, KTTH sẽ kiểm tra kỹ lại.";
    return { success: true, message: msg };

  } catch (error) {
    return { success: false, message: "❌ Lỗi hệ thống: " + error.toString() };
  } finally {
    lock.releaseLock();
  }
}


/* ============================================================================
   7. QUY TRÌNH DUYỆT 2 CẤP: KTTH DUYỆT -> QUẢN LÝ PHÊ DUYỆT
   ============================================================================ */

function ktthDuyetHoSo(emailKTTH, idGiaoDich, ketQua, ghiChu) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var thongTin = _layThongTinPhanQuyen(emailKTTH);
    if (!thongTin || thongTin.vaiTro !== "KTTH") throw new Error("⛔ Chỉ KTTH mới được duyệt bước này.");

    var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idGiaoDich) {
        if (data[i][2].toString().trim() !== thongTin.donVi) throw new Error("⛔ Hồ sơ không thuộc đơn vị bạn phụ trách.");
        if (data[i][13] !== "Chờ KTTH duyệt") throw new Error("⛔ Hồ sơ này không ở trạng thái chờ KTTH duyệt (hiện tại: " + data[i][13] + ").");

        var trangThaiMoi = (ketQua === "Duyệt") ? "Chờ Quản lý phê duyệt" : "KTTH từ chối - Yêu cầu nộp lại";
        var row = i + 1;
        sheet.getRange(row, 14).setValue(trangThaiMoi);
        sheet.getRange(row, 15).setValue(emailKTTH);
        sheet.getRange(row, 16).setValue(Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss"));
        sheet.getRange(row, 19).setValue(ghiChu || "");

        var emailNguoiNop = data[i][8];
        if (emailNguoiNop) _guiMailKetQuaDuyet(emailNguoiNop, data[i][3], data[i][5], "KTTH", trangThaiMoi, ghiChu);
        return "✅ Đã " + (ketQua === "Duyệt" ? "duyệt, chuyển Quản lý phê duyệt" : "từ chối, yêu cầu nộp lại") + " hồ sơ " + idGiaoDich;
      }
    }
    return "❌ Không tìm thấy hồ sơ.";
  } finally {
    lock.releaseLock();
  }
}

function quanLyPheDuyet(emailQuanLy, idGiaoDich, ketQua, ghiChu) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var thongTin = _layThongTinPhanQuyen(emailQuanLy);
    if (!thongTin || thongTin.vaiTro !== "Quản lý") throw new Error("⛔ Chỉ Quản lý mới được phê duyệt bước này.");
    var dsDonViQuanLy = thongTin.donVi.split(",").map(function (s) { return s.trim(); });

    var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === idGiaoDich) {
        if (dsDonViQuanLy.indexOf(data[i][2].toString().trim()) === -1) throw new Error("⛔ Hồ sơ không thuộc đơn vị bạn phụ trách.");
        if (data[i][13] !== "Chờ Quản lý phê duyệt") throw new Error("⛔ Hồ sơ này chưa qua KTTH duyệt hoặc đã xử lý xong (hiện tại: " + data[i][13] + ").");

        var trangThaiMoi = (ketQua === "Duyệt") ? "Đã phê duyệt" : "Quản lý từ chối - Yêu cầu nộp lại";
        var row = i + 1;
        sheet.getRange(row, 14).setValue(trangThaiMoi);
        sheet.getRange(row, 17).setValue(emailQuanLy);
        sheet.getRange(row, 18).setValue(Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm:ss"));
        if (ghiChu) sheet.getRange(row, 19).setValue(ghiChu);

        var emailNguoiNop = data[i][8];
        if (emailNguoiNop) _guiMailKetQuaDuyet(emailNguoiNop, data[i][3], data[i][5], "Quản lý", trangThaiMoi, ghiChu);
        return "✅ Đã " + (ketQua === "Duyệt" ? "PHÊ DUYỆT - CHỐT SỐ LIỆU" : "từ chối, yêu cầu nộp lại") + " hồ sơ " + idGiaoDich;
      }
    }
    return "❌ Không tìm thấy hồ sơ.";
  } finally {
    lock.releaseLock();
  }
}

function _guiMailKetQuaDuyet(email, loaiBaoCao, kyBaoCao, capDuyet, trangThai, ghiChu) {
  try {
    var tieuDe = "[HAK] Kết quả " + capDuyet + " duyệt: " + loaiBaoCao + " - " + kyBaoCao;
    var noiDung = "<p>Hồ sơ <b>" + loaiBaoCao + " - " + kyBaoCao + "</b> bạn đã nộp có kết quả:</p>" +
      "<p style='font-size:16px'><b>" + trangThai + "</b></p>" +
      (ghiChu ? "<p>Ghi chú: " + ghiChu + "</p>" : "") +
      "<p>Trân trọng, Hệ Thống Quản Trị HAK</p>";
    MailApp.sendEmail({ to: email, subject: tieuDe, htmlBody: noiDung });
  } catch (e) { console.log("Lỗi gửi mail kết quả duyệt: " + e.toString()); }
}

function layDanhSachChoDuyet(email) {
  var thongTin = _layThongTinPhanQuyen(email);
  if (!thongTin) return [];

  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
  var data = sheet.getDataRange().getValues();
  var ketQua = [];

  var dsDonViPhuTrach = thongTin.donVi.split(",").map(function (s) { return s.trim(); });
  var trangThaiCanXem = (thongTin.vaiTro === "KTTH") ? "Chờ KTTH duyệt"
    : (thongTin.vaiTro === "Quản lý") ? "Chờ Quản lý phê duyệt" : null;
  if (!trangThaiCanXem) return [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][13] === trangThaiCanXem && data[i][19] === "Có" &&
        dsDonViPhuTrach.indexOf(data[i][2].toString().trim()) !== -1) {
      ketQua.push({
        idGiaoDich: data[i][0], donVi: data[i][2], loaiBaoCao: data[i][3], nam: data[i][4],
        kyBaoCao: data[i][5], tenFile: data[i][6], linkFile: data[i][7], maHoSo: data[i][9],
        trangThaiMau: data[i][10], hanNop: data[i][11], trangThaiHan: data[i][12]
      });
    }
  }
  return ketQua;
}


/* ============================================================================
   8. BÁO CÁO THIẾU HỒ SƠ (dựa trên cột "Là bản mới nhất")
   ============================================================================ */

function layBaoCaoThieuHoSo(donVi, loaiBaoCao, nam, kyBaoCao) {
  var danhMuc = layDanhSachHoSoTheoLoai(loaiBaoCao).filter(function (h) { return h.batBuoc; });

  var sheetDb = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
  var dbData = sheetDb.getDataRange().getValues();

  var maHoSoDaNopHopLe = {};
  for (var i = 1; i < dbData.length; i++) {
    if (dbData[i][2] === donVi && dbData[i][3] === loaiBaoCao &&
        dbData[i][4].toString() === nam.toString() && dbData[i][5] === kyBaoCao &&
        dbData[i][19] === "Có") {
      var ma = dbData[i][9] ? dbData[i][9].toString().trim() : "";
      if (ma && TRANG_THAI_TU_CHOI.indexOf(dbData[i][13]) === -1) maHoSoDaNopHopLe[ma] = true;
    }
  }

  var thieu = danhMuc.filter(function (h) { return !maHoSoDaNopHopLe[h.maHoSo]; });
  return {
    donVi: donVi,
    daDu: thieu.length === 0,
    danhSachThieu: thieu.map(function (h) { return h.tenHoSo; })
  };
}

function layTongHopThieuHoSoHopNhat(loaiBaoCao, nam, kyBaoCao) {
  var dsDonVi = ['HAK_DN', 'CNHAK_QS', 'HAKQN_QC', 'DAIHIEP_DH'];
  return dsDonVi.map(function (dv) { return layBaoCaoThieuHoSo(dv, loaiBaoCao, nam, kyBaoCao); });
}


/* ============================================================================
   9. NHẮC HẠN NỘP TỰ ĐỘNG
   ============================================================================ */

function guiCanhBaoTruocHan(soNgayTruoc) {
  soNgayTruoc = soNgayTruoc || 3;
  var homNay = new Date();

  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('KE_HOACH');
  var data = sheet.getDataRange().getValues();

  var danhSachKy = [];
  for (var t = 1; t <= 12; t++) danhSachKy.push('Tháng ' + t);
  for (var q = 1; q <= 4; q++) danhSachKy.push('Quý ' + q);
  var namXet = [homNay.getFullYear() - 1, homNay.getFullYear()];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var donVi = data[i][0].toString().trim();
    var loaiBaoCao = data[i][1].toString().trim();
    var ngayDenHan = Number(data[i][2]);
    var emails = data[i][3] ? data[i][3].toString().split(',').map(function (e) { return e.trim(); }).filter(Boolean) : [];
    if (!ngayDenHan || emails.length === 0) continue;

    namXet.forEach(function (nam) {
      danhSachKy.forEach(function (ky) {
        var han = tinhHanNop(ky, nam, ngayDenHan);
        if (!han) return;
        var soNgayConLai = Math.floor((han - homNay) / (1000 * 60 * 60 * 24));
        if (soNgayConLai < 0 || soNgayConLai > soNgayTruoc) return;

        var baoCaoThieu = layBaoCaoThieuHoSo(donVi, loaiBaoCao, nam, ky);
        if (baoCaoThieu.daDu) return;

        var tieuDe = "[HAK] ⏰ Nhắc hạn nộp: " + loaiBaoCao + " - " + donVi + " - " + ky + "/" + nam;
        var noiDung = "<p>Đơn vị <b>" + donVi + "</b> còn <b>" + soNgayConLai + "</b> ngày nữa đến hạn nộp <b>" + loaiBaoCao +
          "</b> kỳ <b>" + ky + "/" + nam + "</b> (hạn " + Utilities.formatDate(han, "GMT+7", "dd/MM/yyyy") + ").</p>" +
          "<p>Hồ sơ còn thiếu: <b>" + baoCaoThieu.danhSachThieu.join(", ") + "</b></p>";

        emails.forEach(function (email) {
          try { MailApp.sendEmail({ to: email, subject: tieuDe, htmlBody: noiDung }); } catch (e) { console.log(e); }
        });
      });
    });
  }
}


/* ============================================================================
   10. DASHBOARD TỔNG QUAN (dành cho Admin)
   ============================================================================ */

function layThongKeTongQuan() {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
  var data = sheet.getDataRange().getValues();

  var choKTTH = 0, choQuanLy = 0, treHanDangXuLy = 0, daPheDuyetThangNay = 0;
  var homNay = new Date();
  var thangNay = homNay.getMonth(), namNay = homNay.getFullYear();
  var theoDonVi = {};

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var donVi = data[i][2].toString().trim();
    var trangThaiDuyet = data[i][13];
    var trangThaiHan = data[i][12];
    var laBanMoiNhat = data[i][19];

    if (!theoDonVi[donVi]) theoDonVi[donVi] = { treHan: 0, dangChoXuLy: 0 };

    if (laBanMoiNhat === "Có") {
      if (trangThaiDuyet === "Chờ KTTH duyệt") { choKTTH++; theoDonVi[donVi].dangChoXuLy++; }
      if (trangThaiDuyet === "Chờ Quản lý phê duyệt") { choQuanLy++; theoDonVi[donVi].dangChoXuLy++; }
      if (trangThaiHan === "Trễ hạn" && TRANG_THAI_TU_CHOI.indexOf(trangThaiDuyet) === -1) {
        treHanDangXuLy++;
        theoDonVi[donVi].treHan++;
      }
    }

    if (trangThaiDuyet === "Đã phê duyệt" && data[i][17]) {
      var ngayPD = _parseNgayVN(data[i][17]);
      if (ngayPD && ngayPD.getMonth() === thangNay && ngayPD.getFullYear() === namNay) daPheDuyetThangNay++;
    }
  }

  var dsTheoDonVi = Object.keys(theoDonVi).map(function (dv) {
    return { donVi: dv, treHan: theoDonVi[dv].treHan, dangChoXuLy: theoDonVi[dv].dangChoXuLy };
  });

  return {
    choKTTH: choKTTH,
    choQuanLy: choQuanLy,
    treHanDangXuLy: treHanDangXuLy,
    daPheDuyetThangNay: daPheDuyetThangNay,
    theoDonVi: dsTheoDonVi
  };
}


/* ============================================================================
   11. MIGRATION - CHẠY 1 LẦN CHO DỮ LIỆU CŨ
   ============================================================================
   Nếu bạn đang nâng cấp từ bản trước (chưa có cột "Là bản mới nhất"), hãy:
   1) Tự thêm tiêu đề cột 20 (T1) sheet DATA_FILE = "Là bản mới nhất"
   2) Chạy hàm này 1 lần (thủ công, trong trình soạn thảo Apps Script)
   Hàm sẽ tự xác định dòng nào là bản ghi mới nhất cho từng tổ hợp
   (đơn vị + loại báo cáo + năm + kỳ + mã hồ sơ) và gán "Có"/"Không" tương ứng.
   ============================================================================ */

function capNhatLaBanMoiNhatChoDuLieuCu() {
  var sheet = SpreadsheetApp.openById(SHEET_DB_ID).getSheetByName('DATA_FILE');
  var data = sheet.getDataRange().getValues();

  var chiSoDongCuoiTheoToHop = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var key = [data[i][2], data[i][3], data[i][4], data[i][5], data[i][9] || ""].join("|");
    chiSoDongCuoiTheoToHop[key] = i; // ghi đè mỗi lần -> cuối cùng giữ lại dòng CUỐI CÙNG trùng tổ hợp
  }

  var laDongMoiNhat = {};
  Object.keys(chiSoDongCuoiTheoToHop).forEach(function (key) {
    laDongMoiNhat[chiSoDongCuoiTheoToHop[key]] = true;
  });

  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    sheet.getRange(i + 1, 20).setValue(laDongMoiNhat[i] ? "Có" : "Không");
  }
  return "✅ Đã cập nhật cột 'Là bản mới nhất' cho " + (data.length - 1) + " dòng dữ liệu.";
}
