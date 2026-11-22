import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const code = fs.readFileSync(new URL('../google-apps-script/Code.gs', import.meta.url),'utf8');
const headers = ['رقم التسجيل','وقت التسجيل','الاسم بالكامل','السن','رقم الموبايل','خبرة المبيعات','نوع العمل','حالة المتابعة','ملاحظات'];
const rows = [headers];
let configured = true, locked = false, failWrites = false, released = 0;
const chain = new Proxy({}, { get: (_, key) => key === 'build' ? () => ({}) : () => chain });
const sheet = {
  getLastRow: () => rows.length,
  getMaxRows: () => 1000,
  getFilter: () => null,
  getRange(row,col,height=1,width=1) {
    return {
      getValues: () => Array.from({length:height},(_,i)=>Array.from({length:width},(_,j)=>rows[row-1+i]?.[col-1+j] ?? '')),
      setValues(values) {
        if (failWrites) throw new Error('storage unavailable');
        values.forEach((values,i) => { rows[row-1+i] ||= []; values.forEach((value,j) => { rows[row-1+i][col-1+j] = value; }); });
        return this;
      },
      setNumberFormat() { return this; },
      setBackground() { return this; },
      setDataValidation() { return this; }
    };
  }
};
const context = vm.createContext({
  SpreadsheetApp: {openById:()=>({getSheetByName:()=>sheet}),newDataValidation:()=>chain,flush:()=>{}},
  PropertiesService: { getScriptProperties:()=>({getProperty:()=>configured?'test-sheet':null}) },
  LockService: {getScriptLock:()=>({tryLock:()=>!locked,releaseLock:()=>released++})},
  ContentService: {createTextOutput:text=>({text,setMimeType(){return this;}}),MimeType:{JSON:'json'}}
});
vm.runInContext(code,context);
const data = {name:'اختبار ربط أساس',age:28,phone:'+20 1000000000',experience:'1-3',status:'individual',requestId:'test-request-0001'};
const register = payload => JSON.parse(JSON.stringify(context.register(payload)));
assert.deepEqual(register(data),{ok:true,id:data.requestId});
assert.equal(rows.length,2);
assert.equal(rows[1][4],'01000000000');
assert.equal(rows[1][5],'1 – 3 سنين');
assert.equal(rows[1][7],'جديد');
assert.deepEqual(register(data),{ok:true,id:data.requestId});
assert.equal(rows.length,2,'retry must not create another row');
assert.equal(register({...data,name:'اسم مختلف'}).code,'invalid_data');
assert.equal(register({...data,requestId:'test-request-0002'}).code,'duplicate_phone');
assert.equal(register({...data,phone:'٠١٠٠٠٠٠٠٠٠٠',requestId:'test-request-0002'}).code,'duplicate_phone');
for (const patch of [{age:17},{age:71},{age:28.5},{age:'28'},{experience:'unknown'},{status:'unknown'},{requestId:'x'},{name:'a'},{phone:'12345'}]) {
  assert.equal(register({...data,...patch}).code,'invalid_data',JSON.stringify(patch));
}
assert.equal(register(null).code,'invalid_data');
assert.equal(register([]).code,'invalid_data');
configured = false;
assert.equal(register(data).code,'not_configured');
configured = true; locked = true;
assert.equal(register(data).code,'busy');
locked = false; failWrites = true;
assert.equal(register({...data,phone:'01100000000',requestId:'test-request-0003'}).code,'save_failed');
failWrites = false;
assert.deepEqual(register({...data,name:'=IMPORTXML("url")',phone:'01100000000',requestId:'test-request-0003'}),{ok:true,id:'test-request-0003'});
assert.equal(rows[2][2].charCodeAt(0),0x200b,'user text must not execute as a formula');
assert.ok(released >= 8,'locks must be released after success and errors');
assert.equal(JSON.parse(context.doPost({postData:{contents:'{bad'}}).text).code,'invalid_data');
assert.equal(JSON.parse(context.doPost({postData:{contents:'x'.repeat(4097)}}).text).code,'invalid_data');

const html = fs.readFileSync(new URL('../Asas_Sales_Form.html',import.meta.url),'utf8');
const script = html.match(/<script>\s*([\s\S]*?)<\/script>/)[1];
new vm.Script(script); // Parse the entire frontend, including embedded template literals.
const sendSource = script.slice(script.indexOf('async function send('),script.indexOf('\nconst en ='));
const makeSender = extras => {
  const context = vm.createContext({GOOGLE_SCRIPT_URL:'',...extras});
  vm.runInContext(sendSource,context);
  return context.send;
};
await assert.rejects(makeSender({})(data), error => error.code === 'not_configured');
const rpc = result => ({google:{script:{run:{withSuccessHandler(fn){this.resolve=fn;return this;},withFailureHandler(fn){this.reject=fn;return this;},register(){this.resolve(result);}}}}});
assert.equal((await makeSender(rpc({ok:true,id:'saved'}))(data)).id,'saved');
await assert.rejects(makeSender(rpc({ok:false,code:'duplicate_phone'}))(data),error=>error.code==='duplicate_phone');
await assert.rejects(makeSender(rpc({ok:true}))(data),error=>error.code==='save_failed');
await assert.rejects(makeSender(rpc(undefined))(data),error=>error.code==='save_failed');
console.log('PASS: backend validation, phone normalization, duplicate prevention, idempotent retries, formula escaping, locking, malformed requests, and confirmed frontend responses.');
