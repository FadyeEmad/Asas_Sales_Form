import fs from 'node:fs/promises';
import { Workbook, SpreadsheetFile } from '@oai/artifact-tool';

const out = new URL('../outputs/asas-google-sheets/', import.meta.url);
await fs.mkdir(out, { recursive: true });
const book = Workbook.create();
const summary = book.worksheets.add('ملخص');
const registrations = book.worksheets.add('التسجيلات');
for (const sheet of [summary, registrations]) {
  sheet.showGridLines = false;
  sheet.tabColor = '#F5B800';
  sheet.getRange('A1:I30').format.font.name = 'Arial';
  sheet.getRange('A1:I30').format.font.size = 12;
  sheet.getRange('A1:I30').format.rowHeight = 28;
  sheet.getRange('A1:I30').format.columnWidth = 23;
}
const headers = ['رقم التسجيل','وقت التسجيل','الاسم بالكامل','السن','رقم الموبايل','خبرة المبيعات','نوع العمل','حالة المتابعة','ملاحظات'];
registrations.getRange('A1:I1').values = [headers];
registrations.getRange('A1:I1').format.fill = '#182321';
registrations.getRange('A1:I1').format.font.color = '#F5B800';
registrations.getRange('A1:I1').format.font.bold = true;
registrations.getRange('A1:I1').format.rowHeight = 42;
registrations.getRange('A1:I1').format.wrapText = true;
registrations.getRange('C:C').format.columnWidth = 32;
registrations.getRange('D:D').format.columnWidth = 10;
registrations.getRange('I:I').format.columnWidth = 40;
registrations.getRange('B2:B1000').setNumberFormat('dd/mm/yyyy hh:mm');
registrations.getRange('E2:E1000').setNumberFormat('@');
registrations.getRange('H2:H1000').dataValidation = { rule: { type: 'list', values: ['جديد','تم التواصل','مقابلة','مقبول','غير مناسب'] } };
registrations.freezePanes.freezeRows(1);
for (const [text, color] of [['جديد','#FFF4CE'],['تم التواصل','#DDEBF7'],['مقابلة','#E4DDF5'],['مقبول','#D9EAD3'],['غير مناسب','#F4CCCC']]) {
  registrations.getRange('H2:H1000').conditionalFormats.add('containsText', { text, format: { fill: color } });
}
summary.getRange('A1:D2').merge();
summary.getRange('A1').values = [['أساس | تسجيل سيلز فري لانسر']];
summary.getRange('A1:D2').format.fill = '#182321';
summary.getRange('A1:D2').format.font.color = '#F5B800';
summary.getRange('A1:D2').format.font.bold = true;
summary.getRange('A1:D2').format.font.size = 19;
summary.getRange('A4:B10').values = [['المؤشر','العدد'],['إجمالي التسجيلات',null],['جديد',null],['تم التواصل',null],['مقابلة',null],['مقبول',null],['غير مناسب',null]];
summary.getRange('B5').formulas = [["=COUNTA('التسجيلات'!A:A)-1"]];
for (let row = 6; row <= 10; row++) summary.getRange(`B${row}`).formulas = [[`=COUNTIF('التسجيلات'!H:H,A${row})`]];
summary.getRange('A12:B14').values = [['نوع العمل','العدد'],['مكتب',null],['فردي',null]];
for (let row = 13; row <= 14; row++) summary.getRange(`B${row}`).formulas = [[`=COUNTIF('التسجيلات'!G:G,A${row})`]];
for (const range of ['A4:B4','A12:B12']) {
  summary.getRange(range).format.fill = '#182321';
  summary.getRange(range).format.font.color = '#FFFFFF';
  summary.getRange(range).format.font.bold = true;
}
summary.getRange('A5:B10').format.fill = '#F2F6F4';
summary.getRange('A13:B14').format.fill = '#F2F6F4';
summary.getRange('A:A').format.columnWidth = 36;
summary.getRange('B:B').format.columnWidth = 16;
summary.getRange('B5:B14').setNumberFormat('0');
summary.getRange('A17:D17').merge();
summary.getRange('A17').values = [['الملخص يتحدث تلقائيًا عند وصول التسجيلات.']];
summary.getRange('A19:D20').merge();
summary.getRange('A19').values = [['تابع كل طلب من عمودي حالة المتابعة وملاحظات في تبويب التسجيلات.']];
summary.getRange('A19:D20').format.wrapText = true;
// Verify counts on disposable fixture rows, then remove all test data.
registrations.getRange('A2:I3').values = [['test-1',null,'اختبار',25,'01000000000','بدون خبرة','مكتب','جديد',''],['test-2',null,'اختبار',30,'01100000000','أقل من سنة','فردي','تم التواصل','']];
book.recalculate();
const counts = [summary.getRange('B5').values[0][0],summary.getRange('B6').values[0][0],summary.getRange('B7').values[0][0],summary.getRange('B13').values[0][0],summary.getRange('B14').values[0][0]];
if (JSON.stringify(counts) !== '[2,1,1,1,1]') throw new Error('Summary verification failed: ' + JSON.stringify(counts));
registrations.getRange('A2:I3').clear({ applyTo: 'contents' });
book.recalculate();
console.log((await book.inspect({kind:'table',range:'ملخص!A4:B14',tableMaxRows:11,tableMaxCols:2,maxChars:2200})).ndjson);
console.log((await book.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:1000})).ndjson);
for (const [sheetName, range, filename] of [['ملخص','A1:D20','summary.png'],['التسجيلات','A1:I7','registrations.png']]) {
  const png = await book.render({sheetName,range,scale:1.5,format:'png'});
  await fs.writeFile(new URL(filename,out),new Uint8Array(await png.arrayBuffer()));
}
await (await SpreadsheetFile.exportXlsx(book)).save(new URL('Asas_Sales_Registrations.xlsx',out).pathname.replace(/^\/([A-Za-z]:)/,'$1'));
console.log('Workbook created and verified.');
