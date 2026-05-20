/**
 * HÀM KHỞI TẠO DATABASE V3 TỰ ĐỘNG (Chạy 1 lần duy nhất)
 * Giúp tạo sẵn 3 bảng: orders, products, coupons với đầy đủ tiêu đề cột chuẩn chỉnh.
 */
function setupDatabaseV3() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ==========================================
  // 1. CẤU HÌNH BẢNG 'orders'
  // ==========================================
  let ordersSheet = ss.getSheetByName('orders');
  if (!ordersSheet) {
    ordersSheet = ss.insertSheet('orders');
  }
  ordersSheet.clear(); // Làm sạch bảng
  const ordersHeaders = [
    'order_id',
    'customer_email',
    'customer_name',
    'cart_summary',
    'cart_json',
    'coupon_code',
    'discount_calculation',
    'total_price',
    'status',
    'created_at',
    'customer_note',
    'admin_note'
  ];
  // Điền tiêu đề, chỉnh chữ đậm, đổi màu nền xanh lá nhạt cho dễ nhìn
  ordersSheet.getRange(1, 1, 1, ordersHeaders.length)
             .setValues([ordersHeaders])
             .setFontWeight("bold")
             .setBackground("#d9ead3");
  ordersSheet.setFrozenRows(1); // Đóng băng dòng tiêu đề

  // ==========================================
  // 2. CẤU HÌNH BẢNG 'products'
  // ==========================================
  let productsSheet = ss.getSheetByName('products');
  if (!productsSheet) {
    productsSheet = ss.insertSheet('products');
  }
  productsSheet.clear();
  const productsHeaders = ['product_id', 'product_name', 'stock_quantity', 'delivery_content', 'status'];
  productsSheet.getRange(1, 1, 1, productsHeaders.length)
               .setValues([productsHeaders])
               .setFontWeight("bold")
               .setBackground("#c9daf8"); // Màu xanh dương nhạt
  productsSheet.setFrozenRows(1);

  // Chèn dữ liệu mẫu cho Kho hàng (Dựa trên file HTML của bạn)
  productsSheet.appendRow(['google-ai-pro', 'Google Gemini Advanced Pro', 10, 'Link kích hoạt Gmail Premium: https://g.co/family/join/test1', 'active']);
  productsSheet.appendRow(['netflix-premium', 'Netflix Premium UltraHD 4K', 5, 'Profile 3 - TK: netflix@digistore.vn | PIN: 2026', 'active']);
  productsSheet.appendRow(['chatgpt-plus', 'Tài khoản ChatGPT Plus (GPT-4o)', 8, 'TK: chatgpt@digistore.vn | Pass: Digistore2026', 'active']);

  // ==========================================
  // 3. CẤU HÌNH BẢNG 'coupons'
  // ==========================================
  let couponsSheet = ss.getSheetByName('coupons');
  if (!couponsSheet) {
    couponsSheet = ss.insertSheet('coupons');
  }
  couponsSheet.clear();
  const couponsHeaders = ['coupon_code', 'discount_type', 'discount_value', 'max_discount_amount', 'stock_quantity', 'expiry_date', 'applicable_product_ids', 'status'];
  couponsSheet.getRange(1, 1, 1, couponsHeaders.length)
               .setValues([couponsHeaders])
               .setFontWeight("bold")
               .setBackground("#fff2cc"); // Màu vàng nhạt
  couponsSheet.setFrozenRows(1);

  // Chèn dữ liệu mẫu cho Mã giảm giá
  // Mã 1: Giảm 10%, tối đa 50k, còn 100 lượt, hạn đến hết năm 2026, áp dụng cho mọi sản phẩm (để trống cột áp dụng lẻ)
  couponsSheet.appendRow(['DIGI2026', 'percent', 10, 50000, 100, '2026-12-31', '', 'active']);
  // Mã 2: Giảm thẳng 20k, không giới hạn tối đa, còn 50 lượt, áp dụng riêng cho mã netflix-premium
  couponsSheet.appendRow(['UUDAI20K', 'fixed', 20000, '', 50, '2026-12-31', 'netflix-premium', 'active']);

  // ==========================================
  // 4. DỌN DẸP SHEET MẶC ĐỊNH
  // ==========================================
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Hiển thị hộp thoại thông báo kết quả
  SpreadsheetApp.getUi().alert("🎉 Chúc mừng! Toàn bộ Database V3 đã được thiết lập tự động hoàn hảo.");
}