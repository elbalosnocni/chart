const SHEET_ID = "1caxiuh1jyzZi8rkz1EqDkzwcr_gf5sIk4XzCh-XBjjQ"; // VD: 1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms
const TEN_SHEET = ""; // Để trống nếu muốn lấy sheet đầu tiên, hoặc điền tên tab (VD: "Sheet1")
const FOLDER_ANH_ID = "1Ek1mU_t9M3VO1W2sJw3J_D8rXYesF-pq";

function doGet(e) {
  const cccdNhap = String(e.parameter.cccd || "").trim();
  
  if (!cccdNhap) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Vui lòng nhập số Căn cước công dân!" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 1. Mở file Google Sheets bằng Sheet ID
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = TEN_SHEET ? ss.getSheetByName(TEN_SHEET) : ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();

  // 2. Kiểm tra xem số CCCD nhập vào có tồn tại trong Cột F (index 5) hay không
  let nguoiDangNhap = null;
  for (let i = 1; i < data.length; i++) {
    // Chuẩn hóa số CCCD (xóa khoảng trắng nếu có)
    const cccdTrongSheet = String(data[i][5]).replace(/\s+/g, ''); 
    if (cccdTrongSheet && cccdTrongSheet === cccdNhap.replace(/\s+/g, '')) {
      nguoiDangNhap = String(data[i][0]).trim(); // Lấy họ tên người đăng nhập (Cột A)
      break;
    }
  }

  // Nếu không tìm thấy CCCD trong danh sách nhân viên -> Từ chối truy cập
  if (!nguoiDangNhap) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Số Căn cước không tồn tại trong hệ thống nhân sự!" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 3. Nếu hợp lệ, tiến hành quét ảnh trong Drive theo MaNV
  const photoMap = {};
  try {
    const folder = DriveApp.getFolderById(FOLDER_ANH_ID);
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      const fileNameWithoutExt = file.getName().split('.')[0].trim().toUpperCase();
      photoMap[fileNameWithoutExt] = `https://google.com{file.getId()}&sz=w400`;
    }
  } catch (err) {
    Logger.log("Lỗi đọc thư mục Drive: " + err);
  }

  // 4. Đóng gói danh sách nhân viên (Tuyệt đối không gửi cột SoCancuoc, LCB, NgaySinh về trình duyệt)
  const employees = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const maNV = String(row[1]).trim().toUpperCase(); // Cột B: MaNV
    if (!maNV) continue;

    employees.push({
      hoten: String(row[0]).trim(),      // Cột A: Hoten
      manv: maNV,                        // Cột B: MaNV
      phongban: String(row[6]).trim(),   // Cột G: Phongban
      bophan: String(row[7]).trim(),     // Cột H: Bophan
      chucvu: String(row[8]).trim(),     // Cột I: Chucvu
      anh: photoMap[maNV] || ""
    });
  }

  return ContentService
    .createTextOutput(JSON.stringify({ 
      user: nguoiDangNhap,
      data: employees 
    }))
    .setMimeType(ContentService.MimeType.JSON);
}
