const SHEET_NAME = 'Config';

// [GET] API - Lấy toàn bộ dữ liệu cấu hình
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();

  var result = {};
  // Bỏ qua hàng 1 (Header), lặp từ hàng 2
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    if (key !== "") {
      result[key] = {
        value: data[i][1],
        version: data[i][2] || 0
      };
    }
  }

  return ContentService.createTextOutput(JSON.stringify({status: 'success', data: result}))
                       .setMimeType(ContentService.MimeType.JSON);
}

// [POST] API - Xử lý Create, Update, Delete
function doPost(e) {
  var lock = LockService.getScriptLock();

  // Chờ tối đa 3 giây để lấy quyền chạy, chống ghi đè
  if (!lock.tryLock(3000)) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Server đang bận, vui lòng thử lại.'}))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action; // CREATE, UPDATE, DELETE
    var data = sheet.getDataRange().getValues();

    var rowIndex = -1;
    // Tìm vị trí của KEY trong file (nếu có)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.key) {
        rowIndex = i + 1; // Index trong mảng bắt đầu từ 0, Row trong Sheet bắt đầu từ 1
        break;
      }
    }

    // Xử lý nhánh CREATE
    if (action === 'CREATE') {
      if (rowIndex !== -1) {
        return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Key này đã tồn tại.'})).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.appendRow([payload.key, payload.value, 0]); // Version khởi tạo = 0
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'Tạo mới thành công.'})).setMimeType(ContentService.MimeType.JSON);
    }

    // Xử lý nhánh UPDATE (Có Optimistic Locking)
    if (action === 'UPDATE') {
      if (rowIndex === -1) {
        return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Không tìm thấy Key.'})).setMimeType(ContentService.MimeType.JSON);
      }

      var currentVersion = sheet.getRange(rowIndex, 3).getValue();
      if (Number(currentVersion) === Number(payload.version)) {
        sheet.getRange(rowIndex, 2).setValue(payload.value);
        sheet.getRange(rowIndex, 3).setValue(Number(currentVersion) + 1); // Tăng version
        SpreadsheetApp.flush();
        return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'Cập nhật thành công.', new_version: Number(currentVersion) + 1})).setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({status: 'conflict', message: 'Dữ liệu đã bị thay đổi ở nơi khác. Vui lòng tải lại.'})).setMimeType(ContentService.MimeType.JSON);
      }
    }

    // Xử lý nhánh DELETE
    if (action === 'DELETE') {
      if (rowIndex === -1) {
         return ContentService.createTextOutput(JSON.stringify({status: 'error', message: 'Không tìm thấy Key để xóa.'})).setMimeType(ContentService.MimeType.JSON);
      }
      sheet.deleteRow(rowIndex);
      SpreadsheetApp.flush();
      return ContentService.createTextOutput(JSON.stringify({status: 'success', message: 'Xóa thành công.'})).setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: error.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}