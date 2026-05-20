function capQuyenEmail() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Test quyền", "Cấp quyền thành công!");
}

// ==============================================================================
// 1. CẤU HÌNH & HẰNG SỐ (CONSTANTS)
// (Giúp code dễ đọc, dễ bảo trì thay vì dùng các con số 0, 1, 2 vô hồn)
// ==============================================================================
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Vị trí (Index) các cột trong Sheet 'coupons' (Đếm từ 0 theo mảng Array JS)
const IDX_COUPON = {
  CODE: 0,
  TYPE: 1,
  VALUE: 2,
  MAX_DISCOUNT: 3,
  STOCK: 4,
  EXPIRY: 5,
  APPLICABLE: 6,
  STATUS: 7
};

// ==============================================================================
// 2. ROUTER: XỬ LÝ GET REQUEST (Lấy dữ liệu từ Web)
// ==============================================================================
function doGet(e) {
  const action = e.parameter.action;
  let response = { success: false, message: 'Action not found' };

  try {
    if (action === 'get_coupons') {
      response = getCoupons();
    }
  } catch (error) {
    response.message = error.message;
  }

  // Trả về JSON cho Web hiểu
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================================
// 3. ROUTER: XỬ LÝ POST REQUEST (Nhận dữ liệu từ Web đẩy lên)
// ==============================================================================
function doPost(e) {
  let response = { success: false, message: 'Invalid request' };

  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    if (action === 'create_order') {
      response = createOrder(postData.data);
    }
  } catch (error) {
    response.message = error.message;
  }

  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==============================================================================
// 4. LOGIC: LẤY MÃ GIẢM GIÁ (CÓ SỬ DỤNG CACHE)
// ==============================================================================
function getCoupons() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("COUPONS_DATA");

  // NẾU CÓ CACHE: Trả về luôn, tốc độ 0.05s
  if (cachedData) {
    return { success: true, coupons: JSON.parse(cachedData), source: 'cache' };
  }

  // NẾU KHÔNG CÓ CACHE: Đọc từ Database
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('coupons');
  const data = sheet.getDataRange().getValues();
  let couponsObj = {};

  // i = 1 để bỏ qua dòng tiêu đề
  for (let i = 1; i < data.length; i++) {
    let status = data[i][IDX_COUPON.STATUS];
    let stock = data[i][IDX_COUPON.STOCK];

    // Chỉ lấy mã đang 'active' và còn lượt dùng > 0
    if (status === 'active' && stock > 0) {
      let code = data[i][IDX_COUPON.CODE].toString().toUpperCase();
      couponsObj[code] = {
        type: data[i][IDX_COUPON.TYPE],
        value: data[i][IDX_COUPON.VALUE],
        max_discount: data[i][IDX_COUPON.MAX_DISCOUNT],
        expiry: data[i][IDX_COUPON.EXPIRY],
        applicable: data[i][IDX_COUPON.APPLICABLE]
      };
    }
  }

  // Lưu vào Cache 6 giờ (21600 giây)
  cache.put("COUPONS_DATA", JSON.stringify(couponsObj), 21600);

  return { success: true, coupons: couponsObj, source: 'database' };
}

// ==============================================================================
// 5. TRIGGER: TỰ XÓA CACHE KHI ADMIN SỬA SHEET MÃ GIẢM GIÁ
// (Hàm này tự động chạy ngầm mỗi khi bạn gõ gì đó vào bảng tính)
// ==============================================================================
function onEdit(e) {
  if (!e || !e.source) return;
  const sheetName = e.source.getActiveSheet().getName();

  // Nếu Admin sửa ở Sheet 'coupons', lập tức xóa Cache để lần sau lấy data mới
  if (sheetName === 'coupons') {
    CacheService.getScriptCache().remove("COUPONS_DATA");
  }
}

// ==============================================================================
// 6. LOGIC: TẠO ĐƠN HÀNG & GỬI EMAIL THÚC GIỤC THANH TOÁN
// ==============================================================================
function createOrder(data) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('orders');

  // Lấy giờ Việt Nam
  const timestamp = new Date().toLocaleString("vi-VN", {timeZone: "Asia/Ho_Chi_Minh"});

  // Thêm dòng mới vào Sheet 'orders'. Thứ tự biến phải khớp với cột trong Database V3
  sheet.appendRow([
    data.order_id,             // Cột A
    data.email,                // Cột B
    data.name,                 // Cột C
    data.cart_summary,         // Cột D (Văn xuôi mô tả giỏ hàng)
    JSON.stringify(data.cart), // Cột E (Data JSON để máy tính xử lý sau này)
    data.coupon_code,          // Cột F
    data.discount_calc,        // Cột G (Văn xuôi cách tính tiền)
    data.total_price,          // Cột H
    'pending',                 // Cột I (Trạng thái mặc định: Chờ thanh toán)
    timestamp,                 // Cột J
    data.note,                 // Cột K
    ''                         // Cột L (Admin note - để trống)
  ]);

  SpreadsheetApp.flush(); // Ép Google Sheets ghi data ngay lập tức

  // Format số tiền có dấu phẩy cho đẹp (VD: 150000 -> 150.000)
  const formattedPrice = data.total_price.toLocaleString('vi-VN');

  // Chuẩn bị nội dung Email
  // Tách văn xuôi giỏ hàng thành các dòng bằng thẻ <br>
  const emailCartList = data.cart_summary.split(' | ').join('<br>📦 ');

  // URL mã QR tự động ghép với tổng tiền và mã đơn hàng
  const qrCodeUrl = `https://img.vietqr.io/image/TPBank-0868663315-compact1.jpg?amount=${data.total_price}&addInfo=DIGISTORE%20${data.order_id}&accountName=NGUYEN%20DINH%20HIEU`;

  // HTML Email được làm đẹp, làm nổi bật tiền và thêm ảnh QR
  const emailHtml = `
    <div style="font-family: sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 10px; background: #ffffff;">
      <h2 style="color: #6d28d9; text-align: center;">Cảm ơn bạn đã đặt hàng tại DigiStore (Hiếu iceTea)!</h2>
      <p>Xin chào <b>${data.name}</b>,</p>
      <p>Hệ thống đã ghi nhận đơn hàng <b>${data.order_id}</b> của bạn.</p>

      <div style="background: #f5f3ff; padding: 20px; border-radius: 8px; border-left: 4px solid #7c3aed; margin: 20px 0;">
        <h4 style="margin-top:0; color: #4c1d95; font-size: 15px; text-transform: uppercase;">Danh sách sản phẩm:</h4>
        <div style="color: #333; font-size: 14px; line-height: 1.8; margin-bottom: 15px;">
           📦 ${emailCartList}
        </div>

        <div style="background: #ffffff; padding: 15px; border-radius: 8px; border: 1px dashed #c4b5fd;">
           <p style="color: #6d28d9; font-size: 13px; margin: 0 0 10px 0; line-height: 1.5;"><i>🏷️ Phân tích giá:<br>${data.discount_calc}</i></p>
           <h3 style="color: #ef4444; margin: 0; font-size: 18px; border-top: 1px solid #f1f5f9; padding-top: 10px;">Thành tiền: ${formattedPrice} VNĐ</h3>
        </div>
      </div>

      <p style="text-align: center; font-size: 15px; font-weight: bold; color: #333;">Mời bạn quét mã QR dưới đây để thanh toán:</p>

      <div style="text-align: center; margin: 20px 0;">
         <img src="${qrCodeUrl}" alt="QR Thanh Toán" style="width: 250px; height: 250px; border-radius: 12px; border: 1px solid #eee; padding: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
      </div>

      <div style="background: #f8fafc; padding: 15px; border-radius: 8px; border: 1px solid #e2e8f0; font-size: 14px; line-height: 1.8;">
        <p style="margin: 0 0 10px 0; font-weight: bold; color: #475569;">Hoặc chuyển khoản thủ công:</p>
        <ul style="list-style: none; padding: 0; margin: 0;">
          <li>🏦 Ngân hàng: <b>TP Bank</b></li>
          <li>💳 Số tài khoản: <b>0868663315</b></li>
          <li>👤 Chủ tài khoản: <b>NGUYEN DINH HIEU (Cong ty TNHH DiGiStore Viet Nam)</b></li>
          <li>📝 Nội dung (Bắt buộc): <b style="color: #ef4444; background: #fee2e2; padding: 2px 6px; border-radius: 4px;">DIGISTORE ${data.order_id}</b></li>
        </ul>
      </div>

      <p style="font-size: 12px; color: #666;"><i>Hệ thống sẽ tự động quét giao dịch và gửi tài khoản cho bạn trong vòng 30 giây ngay sau khi nhận được thanh toán thành công!</i></p>
    </div>
  `;

  // Gửi email
  MailApp.sendEmail({
    to: data.email,
    subject: `⏳ Đơn hàng ${data.order_id} đang chờ thanh toán - DigiStore`,
    htmlBody: emailHtml
  });

  return { success: true, order_id: data.order_id };
}