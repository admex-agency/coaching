const SPREADSHEET_ID = '1Kygo7rNEKBqEnPV0c8HFOTKcdGheF9E55etj71bVe6s';
const SHEET_NAME = 'Sheet1';

function doPost(e) {
  const fields = e && e.parameter || {};
  const responseOrigin = typeof fields.responseOrigin === 'string' ? fields.responseOrigin : '';
  const responseNonce = typeof fields.responseNonce === 'string' ? fields.responseNonce : '';
  const respond = (state, title, message) =>
    resultPage(state, title, message, responseOrigin, responseNonce);
  const readText = key => typeof fields[key] === 'string' ? fields[key].trim() : '';
  const fullName = readText('fullName');
  const employeeId = readText('employeeId');
  const department = readText('department');
  const jobTitle = readText('jobTitle');
  const question = readText('question');
  const yearsText = readText('yearsAtAdmicro');
  const years = Number(yearsText);
  if (!fullName || fullName.length > 150 || !employeeId || employeeId.length > 50 ||
      !department || department.length > 150 || !jobTitle || jobTitle.length > 150 ||
      question.length > 1200 || !yearsText || !Number.isFinite(years) || years < 0) {
    return respond('warning', 'Thông tin chưa hợp lệ', 'Kiểm tra các trường bắt buộc và gửi lại.');
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) {
    return respond('warning', 'Chưa thể ghi nhận', 'Hệ thống đang bận. Vui lòng thử lại.');
  }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const rowCount = sheet.getLastRow();
    const existing = rowCount > 1 && sheet.getRange(2, 3, rowCount - 1, 1)
      .createTextFinder(employeeId).matchEntireCell(true).useRegularExpression(false).findNext();
    if (existing) {
      return respond('duplicate', 'Mã nhân sự đã được đăng ký', 'Nếu cần cập nhật thông tin, vui lòng phản hồi email mời để liên hệ ban tổ chức.');
    }
    const textCell = value => value ? "'" + value : '';
    sheet.appendRow([
      new Date(), textCell(fullName), textCell(employeeId), textCell(department),
      textCell(jobTitle), years, textCell(question),
    ]);
    SpreadsheetApp.flush();
    return respond('success', 'Đã ghi nhận đăng ký', 'Ban tổ chức sẽ xác nhận suất tham dự riêng. Cảm ơn anh/chị đã đăng ký.');
  } catch {
    return respond('warning', 'Chưa xác nhận được đăng ký', 'Vui lòng thử lại hoặc phản hồi email mời để liên hệ ban tổ chức.');
  } finally {
    lock.releaseLock();
  }
}

function resultPage(state, title, message, origin, nonce) {
  const embedded = /^https:\/\/[a-z0-9.-]+(?::[0-9]{1,5})?$/i.test(origin) &&
    /^[a-f0-9-]{36}$/i.test(nonce);
  const payload = JSON.stringify({ type: 'admex-imc-registration', nonce, state, title, message });
  const callback = embedded
    ? '<script>window.top.postMessage(' + payload + ',' + JSON.stringify(origin) + ');</script>'
    : '';
  const output = HtmlService.createHtmlOutput(
    '<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>' + title +
    '</title><style>body{margin:0;background:#f6fafb;color:#152f41;font:17px/1.6 system-ui,sans-serif}' +
    'main{max-width:580px;margin:12vh auto;padding:28px}h1{line-height:1.15}small{color:#435968}</style></head>' +
    '<body><main><small>admeX · IMC Planning Framework</small><h1>' + title +
    '</h1><p>' + message + '</p></main>' + callback + '</body></html>'
  ).addMetaTag('viewport', 'width=device-width, initial-scale=1');
  return embedded ? output.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL) : output;
}
