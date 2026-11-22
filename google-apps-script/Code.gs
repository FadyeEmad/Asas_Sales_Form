const REGISTRATIONS_SHEET_ = 'التسجيلات';
const HEADERS_ = ['رقم التسجيل', 'وقت التسجيل', 'الاسم بالكامل', 'السن', 'رقم الموبايل', 'خبرة المبيعات', 'نوع العمل', 'حالة المتابعة', 'ملاحظات'];
const EXPERIENCE_ = { '0': 'بدون خبرة', lt1: 'أقل من سنة', '1-3': '1 – 3 سنين', '3-5': '3 – 5 سنين', '5-10': '5 – 10 سنين', '10+': 'أكتر من 10' };
const STATUS_ = { office: 'مكتب', individual: 'فردي' };

// Run once from the Apps Script editor after importing the workbook.
// Data access helpers are private; only registration is exposed to the form.
function setup() {
  const book = SpreadsheetApp.openById('1DIIyrd0CdZWN7l7_9m6HceRv8ndT1hzFzAgYjvVBMJo');
  const sheet = book.getSheetByName(REGISTRATIONS_SHEET_);
  if (!sheet) throw new Error('استورد ملف Asas_Sales_Registrations.xlsx أولًا.');
  checkHeaders_(sheet);
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', book.getId());
  book.setSpreadsheetTimeZone('Africa/Cairo');
  book.setSpreadsheetLocale('ar_EG');
  sheet.setRightToLeft(true);
  sheet.setFrozenRows(1);
  sheet.setHiddenGridlines(true);
  const rows = sheet.getMaxRows() - 1;
  sheet.getRange(1, 1, 1, HEADERS_.length).setBackground('#182321').setFontColor('#F5B800').setFontWeight('bold').setFontSize(12).setWrap(true);
  sheet.setRowHeight(1, 44);
  sheet.setColumnWidths(1, 9, 160);
  sheet.setColumnWidth(3, 240);
  sheet.setColumnWidth(9, 300);
  sheet.getRange(2, 1, rows, 9).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('middle');
  sheet.getRange(2, 2, rows, 1).setNumberFormat('dd/MM/yyyy HH:mm');
  sheet.getRange(2, 5, rows, 1).setNumberFormat('@').setHorizontalAlignment('left');
  sheet.getRange(2, 8, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['جديد', 'تم التواصل', 'مقابلة', 'مقبول', 'غير مناسب'], true).setAllowInvalid(false).build());
  if (!sheet.getFilter()) sheet.getRange(1, 1, sheet.getMaxRows(), 9).createFilter();
  const summary = book.getSheetByName('ملخص');
  if (summary) summary.setRightToLeft(true).setHiddenGridlines(true);
  return book.getUrl();
}

function doGet() {
  return json_({ service: 'asas-sales-registration', ready: !!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') });
}

function doPost(e) {
  let result;
  try {
    const body = e && e.postData && e.postData.contents;
    if (!body || body.length > 4096) return json_({ ok: false, code: 'invalid_data' });
    result = register(JSON.parse(body));
  } catch (_) {
    result = { ok: false, code: 'invalid_data' };
  }
  return json_(result);
}

function register(input) {
  const data = validate_(input);
  if (!data) return { ok: false, code: 'invalid_data' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, code: 'busy' };
  try {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    if (!id) return { ok: false, code: 'not_configured' };
    const sheet = SpreadsheetApp.openById(id).getSheetByName(REGISTRATIONS_SHEET_);
    checkHeaders_(sheet);
    const last = sheet.getLastRow();
    if (last > 1) {
      // Compare only registration fields. Follow-up notes are never returned.
      const existing = sheet.getRange(2, 1, last - 1, 7).getValues();
      const replay = existing.find(row => String(row[0]) === data.requestId);
      if (replay) {
        const same = replay[2] === literal_(data.name) && Number(replay[3]) === data.age && normalizePhone_(replay[4]) === data.phone && replay[5] === EXPERIENCE_[data.experience] && replay[6] === STATUS_[data.status];
        return same ? { ok: true, id: data.requestId } : { ok: false, code: 'invalid_data' };
      }
      if (existing.some(row => normalizePhone_(row[4]) === data.phone)) return { ok: false, code: 'duplicate_phone' };
    }
    const next = last + 1;
    if (next > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), 500);
    sheet.getRange(next, 1, 1, 9).setNumberFormat('@');
    sheet.getRange(next, 2).setNumberFormat('dd/MM/yyyy HH:mm');
    sheet.getRange(next, 4).setNumberFormat('0');
    sheet.getRange(next, 1, 1, 9).setValues([[data.requestId, new Date(), literal_(data.name), data.age, data.phone, EXPERIENCE_[data.experience], STATUS_[data.status], 'جديد', '']]);
    sheet.getRange(next, 1, 1, 9).setBackground(next % 2 ? '#F2F6F4' : '#FFFFFF');
    sheet.getRange(next, 8).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['جديد', 'تم التواصل', 'مقابلة', 'مقبول', 'غير مناسب'], true).setAllowInvalid(false).build());
    const filter = sheet.getFilter();
    if (filter && filter.getRange().getLastRow() < next) {
      const criteria = HEADERS_.map((_, i) => filter.getColumnFilterCriteria(i + 1));
      filter.remove();
      const expanded = sheet.getRange(1, 1, sheet.getMaxRows(), 9).createFilter();
      criteria.forEach((criterion, i) => { if (criterion) expanded.setColumnFilterCriteria(i + 1, criterion); });
    }
    SpreadsheetApp.flush();
    return { ok: true, id: data.requestId };
  } catch (_) {
    return { ok: false, code: 'save_failed' };
  } finally {
    lock.releaseLock();
  }
}

function validate_(input) {
  if (!input || typeof input !== 'object' || Array.isArray(input)) return null;
  const name = typeof input.name === 'string' ? input.name.trim() : '';
  const phone = normalizePhone_(input.phone);
  if (name.length < 3 || name.length > 120 || /[\u0000-\u001f]/.test(name)) return null;
  if (typeof input.age !== 'number' || !Number.isInteger(input.age) || input.age < 18 || input.age > 70) return null;
  if (!/^01[0125]\d{8}$/.test(phone)) return null;
  if (!Object.prototype.hasOwnProperty.call(EXPERIENCE_, input.experience) || !Object.prototype.hasOwnProperty.call(STATUS_, input.status)) return null;
  if (typeof input.requestId !== 'string' || !/^[a-zA-Z0-9-]{12,80}$/.test(input.requestId)) return null;
  return { name, phone, age: input.age, experience: input.experience, status: input.status, requestId: input.requestId };
}

function normalizePhone_(value) {
  return String(value == null ? '' : value).replace(/[\u0660-\u0669]/g, d => d.charCodeAt(0) - 0x660).replace(/[\u06f0-\u06f9]/g, d => d.charCodeAt(0) - 0x6f0).replace(/[\s-]/g, '').replace(/^\+?20/, '0');
}

function literal_(value) { return /^[=+@-]/.test(value) ? '\u200b' + value : value; }

function checkHeaders_(sheet) {
  if (!sheet || JSON.stringify(sheet.getRange(1, 1, 1, HEADERS_.length).getValues()[0]) !== JSON.stringify(HEADERS_)) throw new Error('أعمدة التسجيلات غير مطابقة.');
}

function json_(result) { return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON); }
