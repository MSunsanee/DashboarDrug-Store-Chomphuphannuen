const SHEETS = { MASTER:'รายการยา', BASE:'ยอดตั้งต้น', INVENTORY:'คลังยา', TX:'รับจ่ายยา', REQ:'หลักฐานใบเบิก', COMMON:'ยาที่ใช้บ่อย' };
let CURRENT_FACILITY_=null;
const BASE_SHEETS_={...SHEETS};
const FACILITIES_={
  '05589':{id:'05589',name:'รพ.สต.บ้านชมภูพานเหนือ',requester:'',director:'',legacy:true},
  '05587':{id:'05587',name:'รพ.สต.บ้านต้อน',requester:'',director:''},
  '05588':{id:'05588',name:'รพ.สต.บ้านนายอ',requester:'',director:''},
  '05590':{id:'05590',name:'รพ.สต.หลุบเลา',requester:'',director:''},
  '05591':{id:'05591',name:'รพ.สต.บ้านฮ่องสิม',requester:'',director:''},
  '05592':{id:'05592',name:'รพ.สต.บ้านนางเติ่ง',requester:'',director:''},
  '05593':{id:'05593',name:'รพ.สต.บ้านบ่อเดือนห้า',requester:'',director:''},
  '05594':{id:'05594',name:'รพ.สต.บ้านกกปลาซิว',requester:'',director:''}
};

function publicFacilities_(){return Object.values(FACILITIES_).map(x=>({id:x.id,name:x.name,requester:x.requester||'',director:x.director||''}))}
function isDistrictAdmin_(user){return !!(user&&user.username==='Admin')}
function facilityForRequest_(p,user){
  const requested=String(p&&p.facilityId||'');
  const accountFacility=user&&(user.facilityId||FACILITIES_[user.username]&&user.username),id=isDistrictAdmin_(user)&&FACILITIES_[requested]?requested:String(accountFacility||'05589');
  return FACILITIES_[id]||FACILITIES_['05589'];
}
function useFacility_(facility){
  Object.keys(BASE_SHEETS_).forEach(k=>SHEETS[k]=BASE_SHEETS_[k]);
  if(!facility.legacy){const prefix=facility.id+'_';['BASE','INVENTORY','TX','REQ','COMMON'].forEach(k=>SHEETS[k]=prefix+BASE_SHEETS_[k])}
  CURRENT_FACILITY_=facility;return facility;
}

function doGet(e){
  const params=(e&&e.parameter)||{};
  if(params.action)return json_(handle_(params.action,parse_(params.payload)));
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบคลังยาและเวชภัณฑ์ อำเภอภูพาน')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function parse_(s){ try{return JSON.parse(s||'{}')}catch(_){return {}} }
function json_(value){ return ContentService.createTextOutput(JSON.stringify({ok:true,data:value})).setMimeType(ContentService.MimeType.JSON); }

function handle_(action,p){
  try{
    // Login/logout do not read spreadsheet data. Bypass sheet setup for a much
    // faster response, especially when the workbook has many rows.
    if(action==='login')return login_(p);
    if(action==='logout')return logout_(p.authToken);
    const user=sessionUser_(p.authToken),facility=useFacility_(facilityForRequest_(p,user));
    if(p.authToken&&!user)throw new Error('เซสชันหมดอายุ กรุณาออกจากระบบแล้วเข้าสู่ระบบใหม่');
    if(user)trackFacilityUsage_(facility,user,'activity');
    setup_();ensureUsageColumn_();formatSheetsOnce_();
    const routes={
      bootstrap:()=>bootstrapFacility_(facility,user), districtOverview:()=>districtOverview_(user), login:()=>login_(p), logout:()=>logout_(p.authToken), createInventory:()=>saveInventory_(p,false), updateInventory:()=>saveInventory_(p,true),
      deleteInventory:()=>deleteInventory_(p.id,p.drugKey,p.lotNo), createTransaction:()=>saveManualTx_(p,false),
      updateTransaction:()=>saveManualTx_(p,true), deleteTransaction:()=>deleteTx_(p.id),
      clearFacilityData:()=>clearFacilityData_(),
      createMaster:()=>createMasterUsage_(p), updateMaster:()=>updateMasterUsageFull_(p), bulkCreateMaster:()=>bulkCreateMasterUsage_(p.rows), deleteMaster:()=>deleteMaster_(p.key), updateMasterStockType:()=>updateMasterStockType_(p), updateMasterUsage:()=>updateMasterUsage_(p), updateMasterUnit:()=>updateMasterUnit_(p), updateMasterLabelInfo:()=>updateMasterLabelInfo_(p),
      createRequisition:()=>createRequisitionV6_(p), deleteRequisition:()=>deleteById_(SHEETS.REQ,p.id),
      createCommon:()=>createCommon_(p), deleteCommon:()=>deleteById_(SHEETS.COMMON,p.id),
      lineAlertStatus:()=>lineAlertStatus_(), testLineAlert:()=>testLineExpiryAlert(), installLineAlert:()=>installDailyExpiryAlertTrigger()
    };
    if(!routes[action]) throw new Error('ไม่รู้จักคำสั่ง '+action);authorizeAction_(action,p.authToken,p);
    return routes[action]();
  }catch(err){ return {__error:String(err.message||err)}; }
}

function doPost(e){
  const out=handle_(e.parameter.action||'bootstrap',parse_(e.parameter.payload));
  if(out&&out.__error)return ContentService.createTextOutput(JSON.stringify({ok:false,error:out.__error})).setMimeType(ContentService.MimeType.JSON);
  return json_(out);
}

function setup_(){
  const ss=SpreadsheetApp.getActive();
  if(ss.getSpreadsheetTimeZone()!=='Asia/Bangkok')ss.setSpreadsheetTimeZone('Asia/Bangkok');
  const master=ss.getSheetByName(SHEETS.MASTER);
  if(master){let changed=false;if(master.getRange(1,7).getDisplayValue()!=='หมวดมูลค่าคลัง'){master.getRange(1,7).setValue('หมวดมูลค่าคลัง').setFontWeight('bold').setBackground('#dff3ee');changed=true}if(master.getRange(1,8).getDisplayValue()!=='ประเภทคงคลัง'){master.getRange(1,8).setValue('ประเภทคงคลัง').setFontWeight('bold').setBackground('#dff3ee');changed=true}if(changed){if(master.getFilter())master.getFilter().remove();master.getRange(1,1,master.getLastRow(),8).createFilter()}}
  ensure_(ss,SHEETS.BASE,['id','drugKey','lotNo','expiry','balance','unit','unitCost','updatedAt']);
  ensure_(ss,SHEETS.INVENTORY,['id','drugKey','lotNo','expiry','balance','unit','unitCost','updatedAt']);
  ensure_(ss,SHEETS.TX,['id','date','type','drugKey','lotNo','expiry','qty','unit','unitCost','documentNo','recorder','note','createdAt']);
  ensure_(ss,SHEETS.REQ,['id','date','month','type','category','reqNo','unit','requester','total','itemsJson','createdAt']);
  ensure_(ss,SHEETS.COMMON,['id','drugKey','createdAt']);
  ensureLabelInfoColumns_();
}

function ensureLabelInfoColumns_(){
  const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);if(!sh)return;
  const headers=['ชื่อยาภาษาไทย','สรรพคุณ','วิธีใช้','คำเตือน'];
  headers.forEach((h,i)=>{const cell=sh.getRange(1,10+i);if(cell.getDisplayValue()!==h)cell.setValue(h).setFontWeight('bold').setBackground('#dff3ee')});
}
function ensure_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0){sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#dff3ee');sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).createFilter()}return sh}
function rows_(name){const sh=SpreadsheetApp.getActive().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const vals=sh.getDataRange().getValues(),heads=vals.shift().map(String),dateFields=new Set(['date','expiry']);return vals.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>Object.fromEntries(heads.map((h,i)=>{const v=r[i];if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v)){const pattern=h==='month'?'yyyy-MM':dateFields.has(h)?'yyyy-MM-dd':"yyyy-MM-dd'T'HH:mm:ss";return[h,Utilities.formatDate(v,'Asia/Bangkok',pattern)]}return[h,v]})))}
function bootstrap_(){return {master:masterUsage_(),inventory:rows_(SHEETS.INVENTORY),transactions:rows_(SHEETS.TX),requisitions:rows_(SHEETS.REQ),common:rows_(SHEETS.COMMON)}}
function bootstrapFacility_(facility,user){const data=bootstrap_();data.facility={id:facility.id,name:facility.name,requester:facility.requester||'',director:facility.director||''};data.facilities=isDistrictAdmin_(user)?publicFacilities_():[data.facility];data.isDistrictAdmin=isDistrictAdmin_(user);return data}
function sheetRowsForFacility_(facility,key){const name=facility.legacy?BASE_SHEETS_[key]:facility.id+'_'+BASE_SHEETS_[key],sh=SpreadsheetApp.getActive().getSheetByName(name);if(!sh||sh.getLastRow()<2)return[];const values=sh.getDataRange().getValues(),heads=values.shift().map(String);return values.filter(r=>r.some(v=>v!==''&&v!==null)).map(r=>Object.fromEntries(heads.map((h,i)=>[h,r[i]])))}
function districtOverview_(user){
  if(!isDistrictAdmin_(user))throw new Error('เมนูภาพรวมระดับอำเภอสำหรับ Admin เท่านั้น');
  const usage=facilityUsage_();
  return publicFacilities_().map(f=>{const cfg=FACILITIES_[f.id],inventory=sheetRowsForFacility_(cfg,'INVENTORY'),tx=sheetRowsForFacility_(cfg,'TX'),req=sheetRowsForFacility_(cfg,'REQ'),positive=inventory.filter(x=>num_(x.balance)>0),used=usage[f.id]||{};return{...f,lots:positive.length,items:new Set(positive.map(x=>String(x.drugKey))).size,balance:positive.reduce((s,x)=>s+num_(x.balance),0),value:positive.reduce((s,x)=>s+num_(x.balance)*num_(x.unitCost),0),transactions:tx.length,requisitions:req.length,expired:positive.filter(x=>{const d=parseExpiryDate_(x.expiry);return d&&d<new Date(new Date().setHours(0,0,0,0))}).length,hasUsed:!!used.firstUsedAt,firstUsedAt:used.firstUsedAt||'',lastLoginAt:used.lastLoginAt||'',lastActiveAt:used.lastActiveAt||'',loginCount:num_(used.loginCount),lastUser:used.lastUser||''}})
}
function facilityUsage_(){try{return JSON.parse(PropertiesService.getScriptProperties().getProperty('FACILITY_USAGE')||'{}')}catch(_){return{}}}
function trackFacilityUsage_(facility,user,event){
  if(!facility||!user||isDistrictAdmin_(user))return;
  const props=PropertiesService.getScriptProperties(),usage=facilityUsage_(),id=String(facility.id),now=new Date().toISOString(),current=usage[id]||{};
  if(event==='activity'&&current.lastActiveAt&&Date.now()-new Date(current.lastActiveAt).getTime()<300000)return;
  usage[id]={firstUsedAt:current.firstUsedAt||now,lastLoginAt:event==='login'?now:(current.lastLoginAt||''),lastActiveAt:now,loginCount:num_(current.loginCount)+(event==='login'?1:0),lastUser:user.name||user.username||''};
  props.setProperty('FACILITY_USAGE',JSON.stringify(usage));
}
function master_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);if(!sh)throw new Error('ไม่พบชีต “รายการยา” กรุณานำไฟล์ Excel เข้า Google Sheets ก่อน');const v=sh.getDataRange().getDisplayValues();v.shift();return v.filter(r=>r[1]).map(r=>({no:r[0],name:r[1],form:r[2],strength:r[3],account:r[4],reserve:r[5],category:r[6]||'ยาทั่วไป',stockType:r[7]||'',thaiName:r[9]||'',indication:r[10]||'',directions:r[11]||'',warning:r[12]||'',key:[r[0],r[1],r[2],r[3]].join('|')}))}
function createMaster_(p){const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],stockTypes=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs'];if(!p.no||!p.name||!p.form)throw new Error('กรุณากรอกรหัส ชื่อ และรูปแบบ');if(!categories.includes(p.category))throw new Error('หมวดมูลค่าคลังไม่ถูกต้อง');if(!stockTypes.includes(p.stockType))throw new Error('ประเภทคงคลังไม่ถูกต้อง');const list=master_();if(list.some(x=>String(x.no)===String(p.no)))throw new Error('รหัสลำดับนี้มีอยู่แล้ว');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);sh.appendRow([p.no,p.name,p.form,p.strength||'',p.account||'',p.reserve||'y',p.category,p.stockType]);return true}
function updateMasterStockType_(p){const stockTypes=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs'];if(!stockTypes.includes(p.stockType))throw new Error('ประเภทคงคลังไม่ถูกต้อง');const list=master_(),item=list.find(x=>x.key===p.key);if(!item)throw new Error('ไม่พบรายการยา');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),row=list.indexOf(item)+2;sh.getRange(row,8).setValue(p.stockType);return true}
function inferStockType_(p){const account=String(p.account||'').toLowerCase(),name=String(p.name||'').toLowerCase();if(account.indexOf('herb')>-1||account.indexOf('กัญชา')>-1)return'ยาสมุนไพร';if(/metformin|glipizide|pioglitazone|insulin|amlodipine|enalapril|losartan|atenolol|metoprolol|doxazosin|simvastatin|atorvastatin|gemfibrozil|aspirin|clopidogrel|digoxin|allopurinol/.test(name))return'ยาโรคNCDs';return'ยาทั่วไป'}
function bulkCreateMaster_(rows){if(!Array.isArray(rows)||!rows.length)throw new Error('ไม่พบรายการสำหรับนำเข้า');if(rows.length>2000)throw new Error('นำเข้าได้ครั้งละไม่เกิน 2,000 รายการ');const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),existing=new Set(master_().map(x=>String(x.no))),categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],stockTypes=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs'],out=[];let skipped=0;rows.forEach(p=>{const no=String(p.no||'').trim();if(!no||!p.name||!p.form||existing.has(no)){skipped++;return}const requested=p.category==='วชย.'?'เวชภัณฑ์ที่มิใช่ยา':p.category,category=categories.includes(requested)?requested:'ยาทั่วไป',stockType=stockTypes.includes(p.stockType)?p.stockType:inferStockType_(p);out.push([no,p.name,p.form,p.strength||'',p.account||'',p.reserve||'y',category,stockType]);existing.add(no)});if(out.length)sh.getRange(sh.getLastRow()+1,1,out.length,8).setValues(out);return{added:out.length,skipped:skipped}}finally{lock.releaseLock()}}
function deleteMaster_(key){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),list=master_(),item=list.find(x=>x.key===key);if(!item)throw new Error('ไม่พบรายการ');const used=rows_(SHEETS.BASE).some(x=>x.drugKey===key)||rows_(SHEETS.INVENTORY).some(x=>x.drugKey===key)||rows_(SHEETS.TX).some(x=>x.drugKey===key)||rows_(SHEETS.COMMON).some(x=>x.drugKey===key);if(used)throw new Error('ลบไม่ได้ เนื่องจากรายการนี้มีข้อมูลคงคลัง ประวัติรับ–จ่าย หรืออยู่ในยาที่ใช้บ่อย');const row=list.indexOf(item)+2;sh.deleteRow(row);return true}
function createRequisition_(p){const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],prefixes={'ยาทั่วไป':'ย','เวชภัณฑ์ที่มิใช่ยา':'วชย','วัสดุการแพทย์':'วพ','วัสดุเภสัช':'วภ'};if(!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date||'')))throw new Error('วันที่ใบเบิกไม่ถูกต้อง');const d=new Date(p.date+'T12:00:00');if(isNaN(d)||d.getDay()!==3||![2,4].includes(Math.floor((d.getDate()-1)/7)+1))throw new Error('ใบเบิกบันทึกได้เฉพาะวันพุธสัปดาห์ที่ 2 หรือ 4 ของเดือน');if(!['hospital','patient'].includes(p.type))throw new Error('กรณีการเบิกไม่ถูกต้อง');if(!categories.includes(p.category))throw new Error('ประเภทรายการเบิกไม่ถูกต้อง');const rows=rows_(SHEETS.REQ),prefix=prefixes[p.category],be=d.getFullYear()+543,re=new RegExp('^'+prefix+'(\\d+)/'+be+'$'),next=Math.max(0,...rows.map(x=>{const m=String(x.reqNo||'').match(re);return m?num_(m[1]):0}))+1,expected=prefix+next+'/'+be;if(p.reqNo!==expected)throw new Error('เลขที่ใบเบิกที่ถูกต้องคือ '+expected);if(rows.some(x=>x.reqNo===p.reqNo))throw new Error('เลขที่ใบเบิกนี้ถูกใช้แล้ว');if(!p.requester||!Array.isArray(p.items)||!p.items.length)throw new Error('ข้อมูลใบเบิกไม่ครบ');const clean=p.items.filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REQ),id=Utilities.getUuid();sh.appendRow([id,p.date,p.date.slice(0,7),p.type,p.category,p.reqNo,p.unit||'',p.requester,num_(p.total),JSON.stringify(clean),new Date()]);sh.getRange(sh.getLastRow(),9).setNumberFormat('#,##0.00');return{id:id}}
function createRequisitionV2_(p){const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],prefixes={'ยาทั่วไป':'ย','เวชภัณฑ์ที่มิใช่ยา':'วชย','วัสดุการแพทย์':'วพ','วัสดุเภสัช':'วภ'};if(!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date||'')))throw new Error('วันที่ใบเบิกไม่ถูกต้อง');const d=new Date(p.date+'T12:00:00');if(isNaN(d)||d.getDay()!==3||![2,4].includes(Math.floor((d.getDate()-1)/7)+1))throw new Error('ใบเบิกบันทึกได้เฉพาะวันพุธสัปดาห์ที่ 2 หรือ 4 ของเดือน');if(!['hospital','patient'].includes(p.type))throw new Error('กรณีการเบิกไม่ถูกต้อง');if(!categories.includes(p.category))throw new Error('ประเภทรายการเบิกไม่ถูกต้อง');const reqRows=rows_(SHEETS.REQ),prefix=prefixes[p.category],be=d.getFullYear()+543,re=new RegExp('^'+prefix+'(\\d+)/'+be+'$'),next=Math.max(0,...reqRows.map(x=>{const m=String(x.reqNo||'').match(re);return m?num_(m[1]):0}))+1,expected=prefix+next+'/'+be;if(p.reqNo!==expected)throw new Error('เลขที่ใบเบิกที่ถูกต้องคือ '+expected);if(reqRows.some(x=>x.reqNo===p.reqNo))throw new Error('เลขที่ใบเบิกนี้ถูกใช้แล้ว');if(!p.requester||!Array.isArray(p.items)||!p.items.length)throw new Error('ข้อมูลใบเบิกไม่ครบ');const clean=p.items.filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const txRows=[];if(p.type==='patient'){const masters=master_(),inventory=rows_(SHEETS.INVENTORY);clean.forEach(item=>{const medicine=masters.find(x=>String(x.no)===item.code);if(!medicine)throw new Error('ไม่พบรหัสรายการ '+item.code+' ในฐานข้อมูล');const lots=inventory.filter(x=>x.drugKey===medicine.key&&num_(x.balance)>0).sort((a,b)=>{const ae=a.expiry?new Date(a.expiry).getTime():Number.MAX_SAFE_INTEGER,be=b.expiry?new Date(b.expiry).getTime():Number.MAX_SAFE_INTEGER;return ae-be});const available=lots.reduce((sum,x)=>sum+num_(x.balance),0);if(available<item.qty)throw new Error('ยอดคงเหลือของ '+(medicine.name||item.name)+' ไม่พอจ่าย (คงเหลือ '+available+' '+(item.unit||'')+')');let remaining=item.qty;lots.forEach(lot=>{if(remaining<=0)return;const take=Math.min(remaining,num_(lot.balance));txRows.push([Utilities.getUuid(),date_(p.date),'จ่ายออก',medicine.key,lot.lotNo,date_(lot.expiry),take,lot.unit||item.unit,num_(lot.unitCost),p.reqNo,p.requester||'นางศันสนีย์','ใบเบิกจากคลังเพื่อจ่ายผู้ป่วย · '+p.category,new Date()]);remaining-=take})})}const ss=SpreadsheetApp.getActive(),reqSh=ss.getSheetByName(SHEETS.REQ),id=Utilities.getUuid();reqSh.appendRow([id,p.date,p.date.slice(0,7),p.type,p.category,p.reqNo,p.unit||'',p.requester,num_(p.total),JSON.stringify(clean),new Date()]);reqSh.getRange(reqSh.getLastRow(),9).setNumberFormat('#,##0.00');if(txRows.length){const txSh=ss.getSheetByName(SHEETS.TX),start=txSh.getLastRow()+1;txSh.getRange(start,1,txRows.length,txRows[0].length).setValues(txRows);txSh.getRange(start,2,txRows.length,1).setNumberFormat('yyyy-mm-dd');txSh.getRange(start,6,txRows.length,1).setNumberFormat('yyyy-mm-dd');txSh.getRange(start,7,txRows.length,1).setNumberFormat('#,##0.00');txSh.getRange(start,9,txRows.length,1).setNumberFormat('#,##0.00');rebuildInventory_()}return{id:id,postedTransactions:txRows.length,pendingReceipt:p.type==='hospital'}}finally{lock.releaseLock()}}
function createRequisitionV3_(p){if(p.type!=='hospital')return createRequisitionV2_(p);const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],prefixes={'ยาทั่วไป':'ย','เวชภัณฑ์ที่มิใช่ยา':'วชย','วัสดุการแพทย์':'วพ','วัสดุเภสัช':'วภ'};if(!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date||'')))throw new Error('วันที่ใบเบิกไม่ถูกต้อง');const d=new Date(p.date+'T12:00:00');if(isNaN(d))throw new Error('วันที่ใบเบิกไม่ถูกต้อง');if(!categories.includes(p.category))throw new Error('ประเภทรายการเบิกไม่ถูกต้อง');const reqRows=rows_(SHEETS.REQ),prefix=prefixes[p.category],be=d.getFullYear()+543,re=new RegExp('^'+prefix+'(\\d+)/'+be+'$'),next=Math.max(0,...reqRows.map(x=>{const m=String(x.reqNo||'').match(re);return m?num_(m[1]):0}))+1,expected=prefix+next+'/'+be;if(p.reqNo!==expected)throw new Error('เลขที่ใบเบิกที่ถูกต้องคือ '+expected);if(reqRows.some(x=>x.reqNo===p.reqNo))throw new Error('เลขที่ใบเบิกนี้ถูกใช้แล้ว');if(!p.requester||!Array.isArray(p.items)||!p.items.length)throw new Error('ข้อมูลใบเบิกไม่ครบ');const clean=p.items.filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REQ),id=Utilities.getUuid();sh.appendRow([id,p.date,p.date.slice(0,7),p.type,p.category,p.reqNo,p.unit||'',p.requester,num_(p.total),JSON.stringify(clean),new Date()]);sh.getRange(sh.getLastRow(),9).setNumberFormat('#,##0.00');return{id:id,postedTransactions:0,pendingReceipt:true}}finally{lock.releaseLock()}}
function createRequisitionV4_(p){if(p.type!=='hospital')return createRequisitionV2_(p);const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'];if(!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date||''))||isNaN(new Date(p.date+'T12:00:00')))throw new Error('วันที่ใบเบิกไม่ถูกต้อง');if(!categories.includes(p.category))throw new Error('ประเภทรายการเบิกไม่ถูกต้อง');const reqRows=rows_(SHEETS.REQ),used=reqRows.filter(x=>x.type==='hospital').map(x=>String(x.reqNo||'').match(/^(\d{1,3})\/(\d+)$/)).filter(Boolean).map(m=>({no:num_(m[1]),book:num_(m[2])})),book=Math.max(1,...used.map(x=>x.book)),last=Math.max(0,...used.filter(x=>x.book===book).map(x=>x.no)),expected=(last>=100?1:last+1)+'/'+(last>=100?book+1:book);if(String(p.reqNo)!==expected)throw new Error('เลขที่/เล่มที่ถูกต้องคือ '+expected);if(reqRows.some(x=>x.type==='hospital'&&x.reqNo===p.reqNo))throw new Error('เลขที่/เล่มนี้ถูกใช้แล้ว');if(!p.requester||!Array.isArray(p.items)||!p.items.length)throw new Error('ข้อมูลใบเบิกไม่ครบ');const clean=p.items.filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REQ),id=Utilities.getUuid();sh.appendRow([id,p.date,p.date.slice(0,7),p.type,p.category,p.reqNo,p.unit||'',p.requester,num_(p.total),JSON.stringify(clean),new Date()]);sh.getRange(sh.getLastRow(),9).setNumberFormat('#,##0.00');return{id:id,postedTransactions:0,pendingReceipt:true}}finally{lock.releaseLock()}}
function createCommon_(p){const key=String(p.drugKey||'');if(!key||!master_().some(x=>x.key===key))throw new Error('ไม่พบรายการยาในฐานข้อมูล');const rows=rows_(SHEETS.COMMON);if(rows.some(x=>x.drugKey===key))throw new Error('รายการนี้อยู่ในยาที่ใช้บ่อยแล้ว');const id=Utilities.getUuid();SpreadsheetApp.getActive().getSheetByName(SHEETS.COMMON).appendRow([id,key,new Date()]);return{id:id}}
function saveManualTx_(p,edit){if(p.type==='จ่ายออก'&&!edit)throw new Error('รายการจ่ายออกต้องบันทึกผ่านฟอร์มใบเบิกยาเท่านั้น');return saveTx_(p,edit)}
function saveInventory_(p,edit){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.BASE),id=edit?p.id:Utilities.getUuid(),row=[id,p.drugKey,p.lotNo,date_(p.expiry),num_(p.balance),p.unit,num_(p.unitCost),new Date()];if(edit)setRow_(sh,id,row);else sh.appendRow(row);rebuildInventory_();return {id:id}}
function deleteInventory_(id,drugKey,lotNo){
  const ss=SpreadsheetApp.getActive(),base=ss.getSheetByName(SHEETS.BASE),inventory=rows_(SHEETS.INVENTORY);
  let key=String(drugKey||''),lot=String(lotNo||'');
  if(!key){const current=inventory.find(x=>String(x.id)===String(id));if(current){key=String(current.drugKey||'');lot=String(current.lotNo||'')}}
  if(!key){const direct=findRow_(base,id);if(direct){base.deleteRow(direct);rebuildInventory_();return {deleted:1,source:'opening'}}throw new Error('ไม่พบข้อมูลรายการที่ต้องการลบ กรุณารีเฟรชแล้วลองใหม่')}
  let deleted=0;
  [SHEETS.BASE,SHEETS.TX].forEach(name=>{const sh=ss.getSheetByName(name);if(!sh||sh.getLastRow()<2)return;const values=sh.getRange(2,1,sh.getLastRow()-1,Math.max(5,sh.getLastColumn())).getDisplayValues(),drugCol=name===SHEETS.TX?3:1,lotCol=name===SHEETS.TX?4:2;for(let i=values.length-1;i>=0;i--){if(String(values[i][drugCol])===key&&String(values[i][lotCol]||'')===lot){sh.deleteRow(i+2);deleted++}}});
  if(!deleted)throw new Error('ไม่พบข้อมูลต้นทางของล็อตนี้ กรุณารีเฟรชแล้วลองใหม่');
  rebuildInventory_();return {deleted:deleted,source:'lot'};
}
function saveTx_(p,edit){validateTx_(p,edit?p.id:null);const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.TX),id=edit?p.id:Utilities.getUuid(),row=[id,date_(p.date),p.type,p.drugKey,p.lotNo,date_(p.expiry),num_(p.qty),p.unit,num_(p.unitCost),p.documentNo||'',p.recorder||'นางศันสนีย์',p.note||'',new Date()];if(edit)setRow_(sh,id,row);else sh.appendRow(row);rebuildInventory_();const result={id:id};if(!edit&&(p.type==='รับเข้า'||p.type==='ยอดยกมา'))result.lineAlert=notifyInboundExpiryToLine_(p);return result}
function deleteTx_(id){deleteById_(SHEETS.TX,id);rebuildInventory_();return true}
function clearFacilityData_(){
  const lock=LockService.getDocumentLock();lock.waitLock(30000);
  try{
    const ss=SpreadsheetApp.getActive(),targets=[SHEETS.BASE,SHEETS.INVENTORY,SHEETS.TX,SHEETS.REQ,SHEETS.COMMON],cleared={};
    targets.forEach(name=>{
      const sh=ss.getSheetByName(name),count=sh?Math.max(0,sh.getLastRow()-1):0;
      if(sh&&count)sh.getRange(2,1,count,Math.max(1,sh.getLastColumn())).clearContent();
      cleared[name]=count;
    });
    return{facilityId:CURRENT_FACILITY_&&CURRENT_FACILITY_.id||'',facilityName:CURRENT_FACILITY_&&CURRENT_FACILITY_.name||'',cleared:cleared,total:Object.values(cleared).reduce((sum,n)=>sum+n,0)};
  }finally{lock.releaseLock()}
}
function deleteById_(name,id){const sh=SpreadsheetApp.getActive().getSheetByName(name),r=findRow_(sh,id);if(!r)throw new Error('ไม่พบรายการ');sh.deleteRow(r);return true}
function setRow_(sh,id,row){const r=findRow_(sh,id);if(!r)throw new Error('ไม่พบรายการ');sh.getRange(r,1,1,row.length).setValues([row])}
function findRow_(sh,id){if(sh.getLastRow()<2)return 0;const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().flat();const i=ids.indexOf(String(id));return i<0?0:i+2}
function validateTx_(p,exclude){if(!['ยอดยกมา','รับเข้า','จ่ายออก'].includes(p.type))throw new Error('ประเภทรายการไม่ถูกต้อง');if(num_(p.qty)<=0)throw new Error('จำนวนต้องมากกว่า 0');const tx=rows_(SHEETS.TX),base=rows_(SHEETS.BASE);if(p.type==='ยอดยกมา'){if(tx.some(x=>x.id!==exclude&&x.drugKey===p.drugKey&&x.lotNo===p.lotNo&&x.type==='ยอดยกมา'))throw new Error('ล็อตนี้มีรายการยอดยกมาแล้ว กรุณาแก้ไขรายการเดิม');if(base.some(x=>x.drugKey===p.drugKey&&x.lotNo===p.lotNo&&num_(x.balance)>0))throw new Error('ล็อตนี้มียอดตั้งต้นเดิมอยู่แล้ว กรุณาแก้ไขข้อมูลคลังเดิมแทน')}if(p.type==='จ่ายออก'){const opening=base.filter(x=>x.drugKey===p.drugKey&&x.lotNo===p.lotNo).reduce((a,x)=>a+num_(x.balance),0),total=opening+tx.filter(x=>x.id!==exclude&&x.drugKey===p.drugKey&&x.lotNo===p.lotNo).reduce((a,x)=>a+(x.type==='จ่ายออก'?-1:1)*num_(x.qty),0);if(total<num_(p.qty))throw new Error('ยอดคงเหลือของล็อตนี้ไม่พอจ่าย (เหลือ '+total+')')}}
function rebuildInventory_(){const ss=SpreadsheetApp.getActive(),sh=ss.getSheetByName(SHEETS.INVENTORY),manual=rows_(SHEETS.BASE),tx=rows_(SHEETS.TX),map={};manual.forEach(x=>map[x.drugKey+'§'+x.lotNo]={...x,balance:num_(x.balance),unitCost:num_(x.unitCost)});tx.forEach(x=>{const k=x.drugKey+'§'+x.lotNo;if(!map[k])map[k]={id:Utilities.getUuid(),drugKey:x.drugKey,lotNo:x.lotNo,expiry:x.expiry,balance:0,unit:x.unit,unitCost:num_(x.unitCost)};map[k].balance+=(x.type==='จ่ายออก'?-1:1)*num_(x.qty);map[k].expiry=x.expiry||map[k].expiry;map[k].unit=x.unit||map[k].unit;if(x.type==='รับเข้า'||x.type==='ยอดยกมา')map[k].unitCost=num_(x.unitCost)});if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();const a=Object.values(map).map(x=>[x.id,x.drugKey,x.lotNo,date_(x.expiry),x.balance,x.unit,x.unitCost,new Date()]);if(a.length)sh.getRange(2,1,a.length,8).setValues(a);formatInventory_(sh)}
function formatInventory_(sh){if(sh.getLastRow()<2)return;const n=sh.getLastRow()-1;sh.getRange(2,4,n,1).setNumberFormat('yyyy-mm-dd');sh.getRange(2,5,n,1).setNumberFormat('#,##0.00');sh.getRange(2,7,n,1).setNumberFormat('#,##0.00');const rg=sh.getRange(2,1,n,8),rules=[SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E2>0,$D2<TODAY())').setBackground('#ffd6d6').setRanges([rg]).build(),SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E2>0,$D2>=TODAY(),$D2<EDATE(TODAY(),3))').setBackground('#ffd6a8').setRanges([rg]).build(),SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($E2>0,$D2>=EDATE(TODAY(),3),$D2<EDATE(TODAY(),6))').setBackground('#fff2a8').setRanges([rg]).build()];sh.setConditionalFormatRules(rules)}
function date_(v){if(!v)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.parseDate(Utilities.formatDate(v,'Asia/Bangkok','yyyy-MM-dd'),'Asia/Bangkok','yyyy-MM-dd');const s=String(v).slice(0,10);if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return v;return Utilities.parseDate(s,'Asia/Bangkok','yyyy-MM-dd')}
function ensureUsageColumn_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);if(!sh)return;const lastRow=sh.getLastRow();if(sh.getRange(1,9).getDisplayValue()!=='สถานะการใช้งาน'){sh.getRange(1,9).setValue('สถานะการใช้งาน').setFontWeight('bold').setBackground('#dff3ee');if(sh.getFilter())sh.getFilter().remove();sh.getRange(1,1,Math.max(1,lastRow),9).createFilter()}if(lastRow>1){const rg=sh.getRange(2,9,lastRow-1,1),current=rg.getDisplayValues(),vals=current.map(r=>[r[0]==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้']);if(current.some((r,i)=>r[0]!==vals[i][0]))rg.setValues(vals)}}
function masterUsage_(){const base=master_(),sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),rows=sh.getDataRange().getDisplayValues().slice(1).filter(r=>r[1]);return base.map((x,i)=>({...x,usage:rows[i]&&rows[i][8]==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้'}))}
function createMasterUsage_(p){createMaster_(p);const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);sh.getRange(sh.getLastRow(),9).setValue(p.usage==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้');return true}
function bulkCreateMasterUsage_(rows){const result=bulkCreateMaster_(rows);ensureUsageColumn_();return result}
function updateMasterUsage_(p){if(!['มีใช้','ไม่มีใช้'].includes(p.usage))throw new Error('สถานะการใช้งานไม่ถูกต้อง');const list=master_(),item=list.find(x=>x.key===p.key);if(!item)throw new Error('ไม่พบรายการยา');SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER).getRange(list.indexOf(item)+2,9).setValue(p.usage);return true}
function updateMasterLabelInfo_(p){const list=master_(),item=list.find(x=>x.key===p.key);if(!item)throw new Error('ไม่พบรายการยา');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),row=list.indexOf(item)+2;sh.getRange(row,10,1,4).setValues([[String(p.thaiName||''),String(p.indication||''),String(p.directions||''),String(p.warning||'')]]);return true}
function createRequisitionV5_(p){if(p.type!=='patient')return createRequisitionV4_(p);const existing=rows_(SHEETS.REQ).find(x=>x.type==='patient'&&String(x.date)===String(p.date)&&x.category===p.category);if(!existing)return createRequisitionV2_(p);const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const reqRows=rows_(SHEETS.REQ),current=reqRows.find(x=>x.id===existing.id);if(!current)throw new Error('ไม่พบใบเบิกรอบเดิม กรุณารีเฟรชแล้วลองใหม่');if(String(p.reqNo)!==String(current.reqNo))throw new Error('รอบนี้ต้องเพิ่มในใบเบิกเลขที่ '+current.reqNo);const d=new Date(p.date+'T12:00:00');if(isNaN(d)||d.getDay()!==3||![2,4].includes(Math.floor((d.getDate()-1)/7)+1))throw new Error('วันที่รอบใบเบิกไม่ถูกต้อง');const clean=(p.items||[]).filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const masters=master_(),inventory=rows_(SHEETS.INVENTORY),txRows=[];clean.forEach(item=>{const medicine=masters.find(x=>String(x.no)===item.code);if(!medicine)throw new Error('ไม่พบรหัสรายการ '+item.code+' ในฐานข้อมูล');const lots=inventory.filter(x=>x.drugKey===medicine.key&&num_(x.balance)>0).sort((a,b)=>{const ae=a.expiry?new Date(a.expiry).getTime():Number.MAX_SAFE_INTEGER,be=b.expiry?new Date(b.expiry).getTime():Number.MAX_SAFE_INTEGER;return ae-be}),available=lots.reduce((sum,x)=>sum+num_(x.balance),0);if(available<item.qty)throw new Error('ยอดคงเหลือของ '+(medicine.name||item.name)+' ไม่พอจ่าย (คงเหลือ '+available+' '+(item.unit||'')+')');let remaining=item.qty;lots.forEach(lot=>{if(remaining<=0)return;const take=Math.min(remaining,num_(lot.balance));txRows.push([Utilities.getUuid(),date_(p.date),'จ่ายออก',medicine.key,lot.lotNo,date_(lot.expiry),take,lot.unit||item.unit,num_(lot.unitCost),current.reqNo,p.requester||'นางศันสนีย์','เพิ่มรายการในใบเบิกรอบเดิม · '+p.category,new Date()]);remaining-=take})});let old=[];try{old=JSON.parse(current.itemsJson||'[]')}catch(_){old=[]}const merged=new Map();old.concat(clean).forEach(item=>{const key=String(item.code),found=merged.get(key);if(found){found.qty+=num_(item.qty);found.value+=num_(item.value);found.unitCost=found.qty?found.value/found.qty:num_(item.unitCost)}else merged.set(key,{code:key,name:String(item.name||''),qty:num_(item.qty),unit:String(item.unit||''),unitCost:num_(item.unitCost),value:num_(item.value)})});const reqSh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REQ),row=findRow_(reqSh,current.id),createdAt=reqSh.getRange(row,11).getValue();reqSh.getRange(row,1,1,11).setValues([[current.id,p.date,p.date.slice(0,7),'patient',p.category,current.reqNo,p.unit||current.unit,p.requester||current.requester,num_(current.total)+num_(p.total),JSON.stringify([...merged.values()]),createdAt]]);reqSh.getRange(row,9).setNumberFormat('#,##0.00');if(txRows.length){const txSh=SpreadsheetApp.getActive().getSheetByName(SHEETS.TX),start=txSh.getLastRow()+1;txSh.getRange(start,1,txRows.length,13).setValues(txRows);rebuildInventory_()}return{id:current.id,reqNo:current.reqNo,postedTransactions:txRows.length,pendingReceipt:false,appended:true}}finally{lock.releaseLock()}}
function hospitalReqPages_(record){try{const items=JSON.parse(record.itemsJson||'[]');return Math.max(1,Math.ceil(items.length/15))}catch(_){return 1}}
function hospitalReqSerial_(rawNo){const m=String(rawNo||'').match(/^(\d{1,3})\/(\d+)$/);return m?((num_(m[2])-1)*100+(num_(m[1])-1)):-1}
function hospitalReqNoFromSerial_(serial){return(serial%100+1)+'/'+(Math.floor(serial/100)+1)}
function createHospitalRequisitionSequential_(p){const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const categories=['ยาทั่วไป','ยาโรคเรื้อรัง','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'];if(!/^\d{4}-\d{2}-\d{2}$/.test(String(p.date||''))||isNaN(new Date(p.date+'T12:00:00')))throw new Error('วันที่ใบเบิกไม่ถูกต้อง');if(!categories.includes(p.category))throw new Error('ประเภทรายการเบิกไม่ถูกต้อง');const reqRows=rows_(SHEETS.REQ),lastSerial=Math.max(-1,...reqRows.filter(x=>x.type==='hospital').map(x=>hospitalReqSerial_(x.reqNo)+hospitalReqPages_(x)-1)),expected=hospitalReqNoFromSerial_(lastSerial+1);if(String(p.reqNo)!==expected)throw new Error('เลขที่/เล่มที่เริ่มต้นที่ถูกต้องคือ '+expected);if(!p.requester||!Array.isArray(p.items)||!p.items.length)throw new Error('ข้อมูลใบเบิกไม่ครบ');const clean=p.items.filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const pages=Math.max(1,Math.ceil(clean.length/15)),endSerial=lastSerial+pages;if(Math.floor((lastSerial+1)/100)!==Math.floor(endSerial/100)&&pages>100)throw new Error('จำนวนแผ่นใบเบิกไม่ถูกต้อง');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REQ),id=Utilities.getUuid();sh.appendRow([id,p.date,p.date.slice(0,7),'hospital',p.category,expected,p.unit||'',p.requester,num_(p.total),JSON.stringify(clean),new Date()]);sh.getRange(sh.getLastRow(),9).setNumberFormat('#,##0.00');return{id:id,reqNo:expected,lastReqNo:hospitalReqNoFromSerial_(endSerial),pages:pages,postedTransactions:0,pendingReceipt:true}}finally{lock.releaseLock()}}
function createRequisitionV6_(p){const user=sessionUser_(p.authToken);if(!user)throw new Error('กรุณาเข้าสู่ระบบก่อนสร้างใบเบิก');p.requester=String(p.requester||'').trim();validateRequisitionCategoryItems_(p);if(p.type==='hospital')return createHospitalRequisitionSequential_(p);if(p.category!=='ยาโรคเรื้อรัง')return createRequisitionV5_(p);return createNcdRequisition_(p)}
function validateRequisitionCategoryItems_(p){const masters=master_(),items=(p.items||[]).filter(x=>x.code&&num_(x.qty)>0),seen=new Set();items.forEach(item=>{const code=String(item.code);if(seen.has(code))throw new Error('ไม่สามารถเบิกรายการรหัส '+code+' ซ้ำในใบเบิกเดียวกันได้');seen.add(code);const d=masters.find(x=>String(x.no)===code);if(!d)throw new Error('ไม่พบรายการยา '+item.code);const valid=p.category==='ยาโรคเรื้อรัง'?d.stockType==='ยาโรคNCDs':p.category==='ยาทั่วไป'?(d.stockType||'ยาทั่วไป')==='ยาทั่วไป':d.category===p.category;if(!valid)throw new Error((d.name||item.name)+' ไม่อยู่ในประเภทใบเบิก '+p.category)})}
function createNcdRequisition_(p){if(p.type==='hospital'){const result=createRequisitionV4_({...p,category:'ยาทั่วไป'}),sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.REQ),row=findRow_(sh,result.id);sh.getRange(row,5).setValue('ยาโรคเรื้อรัง');return result}const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const d=new Date(p.date+'T12:00:00');if(isNaN(d)||d.getDay()!==3||![2,4].includes(Math.floor((d.getDate()-1)/7)+1))throw new Error('ใบเบิกบันทึกได้เฉพาะวันพุธสัปดาห์ที่ 2 หรือ 4 ของเดือน');const reqRows=rows_(SHEETS.REQ),existing=reqRows.find(x=>x.type==='patient'&&String(x.date)===String(p.date)&&x.category==='ยาโรคเรื้อรัง'),be=d.getFullYear()+543,re=new RegExp('^ยร(\\d+)/'+be+'$'),next=Math.max(0,...reqRows.map(x=>{const m=String(x.reqNo||'').match(re);return m?num_(m[1]):0}))+1,expected=existing?existing.reqNo:'ยร'+next+'/'+be;if(String(p.reqNo)!==String(expected))throw new Error('เลขที่ใบเบิกที่ถูกต้องคือ '+expected);const clean=(p.items||[]).filter(x=>x.code&&num_(x.qty)>0).map(x=>({code:String(x.code),name:String(x.name||''),qty:num_(x.qty),unit:String(x.unit||''),unitCost:num_(x.unitCost),value:num_(x.value)}));if(!clean.length)throw new Error('ไม่พบรายการเบิก');const masters=master_(),inventory=rows_(SHEETS.INVENTORY),txRows=[];clean.forEach(item=>{const medicine=masters.find(x=>String(x.no)===item.code),lots=inventory.filter(x=>x.drugKey===medicine.key&&num_(x.balance)>0).sort((a,b)=>new Date(a.expiry||'9999-12-31')-new Date(b.expiry||'9999-12-31')),available=lots.reduce((s,x)=>s+num_(x.balance),0);if(available<item.qty)throw new Error('ยอดคงเหลือของ '+medicine.name+' ไม่พอจ่าย (คงเหลือ '+available+')');let remaining=item.qty;lots.forEach(lot=>{if(remaining<=0)return;const take=Math.min(remaining,num_(lot.balance));txRows.push([Utilities.getUuid(),date_(p.date),'จ่ายออก',medicine.key,lot.lotNo,date_(lot.expiry),take,lot.unit||item.unit,num_(lot.unitCost),expected,p.requester||'นางศันสนีย์',(existing?'เพิ่มรายการในใบเบิกรอบเดิม · ':'ใบเบิกจากคลังเพื่อจ่ายผู้ป่วย · ')+'ยาโรคเรื้อรัง',new Date()]);remaining-=take})});const ss=SpreadsheetApp.getActive(),reqSh=ss.getSheetByName(SHEETS.REQ);let id,appended=false;if(existing){id=existing.id;let old=[];try{old=JSON.parse(existing.itemsJson||'[]')}catch(_){old=[]}const merged=new Map();old.concat(clean).forEach(item=>{const key=String(item.code),found=merged.get(key);if(found){found.qty+=num_(item.qty);found.value+=num_(item.value);found.unitCost=found.qty?found.value/found.qty:num_(item.unitCost)}else merged.set(key,{...item})});const row=findRow_(reqSh,id),created=reqSh.getRange(row,11).getValue();reqSh.getRange(row,1,1,11).setValues([[id,p.date,p.date.slice(0,7),'patient','ยาโรคเรื้อรัง',expected,p.unit||existing.unit,p.requester||existing.requester,num_(existing.total)+num_(p.total),JSON.stringify([...merged.values()]),created]]);appended=true}else{id=Utilities.getUuid();reqSh.appendRow([id,p.date,p.date.slice(0,7),'patient','ยาโรคเรื้อรัง',expected,p.unit||'',p.requester,num_(p.total),JSON.stringify(clean),new Date()])}if(txRows.length){const txSh=ss.getSheetByName(SHEETS.TX),start=txSh.getLastRow()+1;txSh.getRange(start,1,txRows.length,13).setValues(txRows);rebuildInventory_()}return{id:id,reqNo:expected,postedTransactions:txRows.length,pendingReceipt:false,appended:appended}}finally{lock.releaseLock()}}
function systemUsers_(){return {
  'Admin':{password:'1234567',name:'นางศันสนีย์ มีพรม',position:'ผู้ดูแลระบบระดับอำเภอ',role:'admin',facilityId:'05589'},
  'สายฝน':{password:'1234',name:'นางสาวสายฝน แก้วสุวรรณ',position:'นักวิชาการสาธารณสุข',role:'requisition',facilityId:'05589'},
  'ชลธิดา':{password:'1234',name:'นางสาวชลธิดา แพงสา',position:'แพทย์แผนไทย',role:'requisition',facilityId:'05589'},
  'ปริญญา':{password:'1234',name:'นายปริญญา กาญจนารัตน์',position:'พนักงานบริการ',role:'requisition',facilityId:'05589'},
  'ปาลิดาภรณ์':{password:'1234',name:'นางสาวปาลิดาภรณ์ พิมราช',position:'ผู้ช่วยเหลือคนไข้',role:'requisition',facilityId:'05589'},
  'visit':{password:'1234',name:'พี่น้องมาหยาม',position:'ผู้เยี่ยมชม · ดูข้อมูลเท่านั้น',role:'viewer',facilityId:'05589'},
  '05587':{password:'p05587',name:'ผู้ใช้งาน รพ.สต.บ้านต้อน',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05587'},
  '05588':{password:'p05588',name:'ผู้ใช้งาน รพ.สต.บ้านนายอ',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05588'},
  '05590':{password:'p05590',name:'ผู้ใช้งาน รพ.สต.หลุบเลา',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05590'},
  '05591':{password:'p05591',name:'ผู้ใช้งาน รพ.สต.บ้านฮ่องสิม',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05591'},
  '05592':{password:'p05592',name:'ผู้ใช้งาน รพ.สต.บ้านนางเติ่ง',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05592'},
  '05593':{password:'p05593',name:'ผู้ใช้งาน รพ.สต.บ้านบ่อเดือนห้า',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05593'},
  '05594':{password:'p05594',name:'ผู้ใช้งาน รพ.สต.บ้านกกปลาซิว',position:'ผู้ดูแลคลังประจำหน่วยบริการ',role:'admin',facilityId:'05594'}
}}
function login_(p){const username=String(p.username||'').trim(),user=systemUsers_()[username];if(!user||String(p.password||'')!==user.password)throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');const facility=FACILITIES_[user.facilityId]||FACILITIES_['05589'],token=Utilities.getUuid()+Utilities.getUuid(),safe={username:username,name:user.name,position:user.position,role:user.role,facilityId:facility.id,facilityName:facility.name,isDistrictAdmin:username==='Admin',loginAt:new Date().toISOString()},session={user:safe,expires:Date.now()+21600000};PropertiesService.getScriptProperties().setProperty('session:'+token,JSON.stringify(session));trackFacilityUsage_(facility,safe,'login');return{token:token,user:safe,facilities:username==='Admin'?publicFacilities_():[{id:facility.id,name:facility.name,requester:'',director:''}]}}
function logout_(token){if(token)PropertiesService.getScriptProperties().deleteProperty('session:'+token);return true}
function sessionUser_(token){if(!token)return null;const key='session:'+token,props=PropertiesService.getScriptProperties(),raw=props.getProperty(key);if(!raw)return null;try{const session=JSON.parse(raw);if(!session.expires||Date.now()>session.expires){props.deleteProperty(key);return null}return session.user}catch(_){props.deleteProperty(key);return null}}
function authorizeAction_(action,token,p){
  if(['login','logout'].includes(action))return;
  const user=sessionUser_(token);
  if(!user)throw new Error('กรุณาเข้าสู่ระบบก่อนดำเนินการ');
  if(action==='bootstrap')return;
  if(action==='clearFacilityData'&&!isDistrictAdmin_(user))throw new Error('ล้างข้อมูลทั้งหมดได้เฉพาะ Admin ระดับอำเภอเท่านั้น');
  const masterActions=new Set(['createMaster','updateMaster','bulkCreateMaster','deleteMaster','updateMasterStockType','updateMasterUsage','updateMasterUnit','updateMasterLabelInfo']);
  if(masterActions.has(action)&&!isDistrictAdmin_(user))throw new Error('แก้ไขรายชื่อยาและหน่วยนับได้เฉพาะ Admin ระดับอำเภอเท่านั้น');
  if(user.role==='admin')return;
  if(user.role==='viewer')throw new Error('บัญชีเยี่ยมชมมีสิทธิ์ดูข้อมูลเท่านั้น ไม่สามารถเพิ่ม แก้ไข หรือลบข้อมูลได้');
  if(user.role==='requisition'&&action==='createRequisition'&&p&&p.type==='patient')return;
  if(user.role==='requisition'&&action==='createRequisition')throw new Error('บัญชีนี้สร้างได้เฉพาะใบเบิกจากคลังยา รพ.สต.');
  throw new Error('บัญชีนี้มีสิทธิ์ดูข้อมูลและสร้างใบเบิกจากคลังยา รพ.สต. เท่านั้น');
}
function formatSheetsOnce_(){const cache=CacheService.getScriptCache(),version='gs-style-2026-09-v3-'+SHEETS.BASE;if(cache.get(version))return;formatAllSheets_();cache.put(version,'1',21600)}
function formatAllSheets_(){const ss=SpreadsheetApp.getActive(),cfg={
  [SHEETS.MASTER]:{color:'#16877c',widths:[75,260,120,110,75,90,140,130,105],formats:{},validations:{7:['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],8:['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs'],9:['มีใช้','ไม่มีใช้']}},
  [SHEETS.BASE]:{color:'#4f8fda',widths:[210,240,115,105,90,85,105,145],formats:{4:'yyyy-mm-dd',5:'#,##0.00',7:'#,##0.00'}},
  [SHEETS.INVENTORY]:{color:'#2d9b82',widths:[210,240,115,105,90,85,105,145],formats:{4:'yyyy-mm-dd',5:'#,##0.00',7:'#,##0.00'}},
  [SHEETS.TX]:{color:'#d29643',widths:[210,105,90,240,115,105,85,85,105,120,120,220,145],formats:{2:'yyyy-mm-dd',6:'yyyy-mm-dd',7:'#,##0.00',9:'#,##0.00'},validations:{3:['ยอดยกมา','รับเข้า','จ่ายออก']}},
  [SHEETS.REQ]:{color:'#8b6fc0',widths:[210,105,90,90,150,105,160,130,110,320,145],formats:{2:'yyyy-mm-dd',9:'#,##0.00'},validations:{4:['hospital','patient'],5:['ยาทั่วไป','ยาโรคเรื้อรัง','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช']}},
  [SHEETS.COMMON]:{color:'#67a948',widths:[210,260,145],formats:{}}
};Object.keys(cfg).forEach(name=>styleSheet_(ss.getSheetByName(name),cfg[name]))}
function styleSheet_(sh,cfg){if(!sh)return;const rows=Math.max(1,sh.getLastRow()),cols=Math.max(1,sh.getLastColumn()),all=sh.getRange(1,1,rows,cols),head=sh.getRange(1,1,1,cols);sh.setHiddenGridlines(true);sh.setFrozenRows(1);sh.setTabColor(cfg.color);all.setFontFamily('Sarabun').setFontSize(10).setVerticalAlignment('middle');head.setBackground('#17665e').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center').setWrap(true);sh.setRowHeight(1,34);if(rows>1){sh.setRowHeights(2,rows-1,28);const body=sh.getRange(2,1,rows-1,cols);body.setFontColor('#173f3a');sh.getBandings().forEach(b=>b.remove());body.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false).setFirstRowColor('#ffffff').setSecondRowColor('#f2f8f6')}if(sh.getFilter())sh.getFilter().remove();all.createFilter();(cfg.widths||[]).forEach((w,i)=>{if(i<cols)sh.setColumnWidth(i+1,w)});Object.keys(cfg.formats||{}).forEach(col=>{if(rows>1&&+col<=cols)sh.getRange(2,+col,rows-1,1).setNumberFormat(cfg.formats[col])});Object.keys(cfg.validations||{}).forEach(col=>{if(rows>1&&+col<=cols){const rule=SpreadsheetApp.newDataValidation().requireValueInList(cfg.validations[col],true).setAllowInvalid(false).build();sh.getRange(2,+col,rows-1,1).setDataValidation(rule)}});if(nameOf_(sh)===SHEETS.MASTER&&rows>1){const usage=sh.getRange(2,9,rows-1,1),rules=sh.getConditionalFormatRules().filter(r=>!r.getRanges().some(x=>x.getSheet().getName()===SHEETS.MASTER&&x.getColumn()===9));rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('ไม่มีใช้').setBackground('#f3f4f6').setFontColor('#8a9693').setRanges([usage]).build());sh.setConditionalFormatRules(rules)}}
function nameOf_(sh){return sh.getName()}
function num_(v){return Number(String(v||0).replace(/,/g,''))||0}

/* หน่วยนับในฐานข้อมูลหลัก: เก็บที่คอลัมน์ N เพื่อไม่กระทบข้อมูลฉลากยาเดิม */
function ensureMasterUnitColumn_(){const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);if(!sh)return;if(sh.getRange(1,14).getDisplayValue()!=='หน่วยนับ')sh.getRange(1,14).setValue('หน่วยนับ').setFontWeight('bold').setBackground('#dff3ee');sh.setColumnWidth(14,100);if(sh.getLastRow()>1){const types=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],rule=SpreadsheetApp.newDataValidation().requireValueInList(types,true).setAllowInvalid(false).build();sh.getRange(2,8,sh.getLastRow()-1,1).setDataValidation(rule)}}
function updateMasterStockType_(p){const types=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'];if(!types.includes(p.stockType))throw new Error('ประเภทคงคลังไม่ถูกต้อง');ensureMasterUnitColumn_();const list=master_(),item=list.find(x=>x.key===p.key);if(!item)throw new Error('ไม่พบรายการ');SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER).getRange(list.indexOf(item)+2,8).setValue(p.stockType);return true}
function updateMasterUnit_(p){ensureMasterUnitColumn_();const unit=String(p.unit||'').trim();if(!unit)throw new Error('กรุณาระบุหน่วยนับ');if(unit.length>50)throw new Error('หน่วยนับยาวเกินไป');const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const list=master_(),item=list.find(x=>x.key===String(p.key||''));if(!item)throw new Error('ไม่พบรายการในฐานข้อมูล');SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER).getRange(list.indexOf(item)+2,14).setValue(unit);return{key:item.key,unit:unit}}finally{lock.releaseLock()}}
function master_(){ensureMasterUnitColumn_();const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);if(!sh)throw new Error('ไม่พบชีต “รายการยา” กรุณานำไฟล์ Excel เข้า Google Sheets ก่อน');const v=sh.getDataRange().getDisplayValues();v.shift();return v.filter(r=>r[1]).map(r=>({no:r[0],name:r[1],form:r[2],strength:r[3],account:r[4],reserve:r[5],category:r[6]||'ยาทั่วไป',stockType:r[7]||'',thaiName:r[9]||'',indication:r[10]||'',directions:r[11]||'',warning:r[12]||'',unit:r[13]||'',key:[r[0],r[1],r[2],r[3]].join('|')}))}
function createMasterUsage_(p){const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],types=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'];if(!p.no||!p.name||!p.form)throw new Error('กรุณากรอกรหัส ชื่อ และรูปแบบ');if(!categories.includes(p.category))throw new Error('หมวดมูลค่าคลังไม่ถูกต้อง');if(!types.includes(p.stockType))throw new Error('ประเภทคงคลังไม่ถูกต้อง');ensureMasterUnitColumn_();if(master_().some(x=>String(x.no)===String(p.no)))throw new Error('รหัสลำดับนี้มีอยู่แล้ว');const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER);sh.appendRow([p.no,p.name,p.form,p.strength||'',p.account||'',p.reserve||'y',p.category,p.stockType,p.usage==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้','','','','',p.unit||'']);return true}
function updateMasterUsageFull_(p){
  const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],types=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'];
  if(!p.originalKey||!p.no||!p.name||!p.form)throw new Error('ข้อมูลสำหรับแก้ไขไม่ครบ');if(!categories.includes(p.category))throw new Error('หมวดมูลค่าคลังไม่ถูกต้อง');if(!types.includes(p.stockType))throw new Error('ประเภทคงคลังไม่ถูกต้อง');
  ensureMasterUnitColumn_();const list=master_(),item=list.find(x=>x.key===String(p.originalKey));if(!item)throw new Error('ไม่พบรายการเดิม กรุณารีเฟรชแล้วลองใหม่');if(list.some(x=>x!==item&&String(x.no)===String(p.no)))throw new Error('รหัสลำดับนี้ถูกใช้โดยรายการอื่นแล้ว');
  const ss=SpreadsheetApp.getActive(),sh=ss.getSheetByName(SHEETS.MASTER),row=list.indexOf(item)+2,oldKey=item.key,newKey=[p.no,p.name,p.form,p.strength||''].join('|');
  sh.getRange(row,1,1,9).setValues([[String(p.no),String(p.name),String(p.form),String(p.strength||''),String(p.account||''),String(p.reserve||'y'),p.category,p.stockType,p.usage==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้']]);sh.getRange(row,14).setValue(String(p.unit||''));
  if(newKey!==oldKey){[[SHEETS.BASE,2],[SHEETS.INVENTORY,2],[SHEETS.TX,4],[SHEETS.COMMON,2]].forEach(cfg=>{const target=ss.getSheetByName(cfg[0]);if(!target||target.getLastRow()<2)return;const range=target.getRange(2,cfg[1],target.getLastRow()-1,1),values=range.getValues();let changed=false;values.forEach(r=>{if(String(r[0])===oldKey){r[0]=newKey;changed=true}});if(changed)range.setValues(values)})}
  return{oldKey:oldKey,newKey:newKey};
}
function bulkCreateMasterUsage_(rows){if(!Array.isArray(rows)||!rows.length)throw new Error('ไม่พบรายการสำหรับนำเข้า');if(rows.length>2000)throw new Error('นำเข้าได้ครั้งละไม่เกิน 2,000 รายการ');ensureMasterUnitColumn_();const lock=LockService.getDocumentLock();lock.waitLock(30000);try{const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),list=master_(),byNo=new Map(list.map((x,i)=>[String(x.no),i+2])),categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],drugTypes=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs'],out=[];let updated=0,skipped=0;rows.forEach(p=>{const no=String(p.no||'').trim(),unit=String(p.unit||'').trim();if(!no||!p.name||!p.form){skipped++;return}if(byNo.has(no)){if(unit){sh.getRange(byNo.get(no),14).setValue(unit);updated++}else skipped++;return}const requested=p.category==='วชย.'?'เวชภัณฑ์ที่มิใช่ยา':p.category,category=categories.includes(requested)?requested:'ยาทั่วไป',stockType=drugTypes.includes(p.stockType)?p.stockType:inferStockType_(p),usage=p.usage==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้';out.push([no,p.name,p.form,p.strength||'',p.account||'',p.reserve||'y',category,stockType,usage,'','','','',unit]);byNo.set(no,sh.getLastRow()+out.length)});if(out.length)sh.getRange(sh.getLastRow()+1,1,out.length,14).setValues(out);return{added:out.length+updated,created:out.length,updated:updated,skipped:skipped}}finally{lock.releaseLock()}}
function bulkCreateMasterUsage_(rows){
  if(!Array.isArray(rows)||!rows.length)throw new Error('ไม่พบรายการสำหรับนำเข้า');
  if(rows.length>2000)throw new Error('นำเข้าได้ครั้งละไม่เกิน 2,000 รายการ');
  ensureMasterUnitColumn_();
  const lock=LockService.getDocumentLock();lock.waitLock(30000);
  try{
    const sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MASTER),list=master_(),byNo=new Map(list.map((x,i)=>[String(x.no),i+2]));
    const categories=['ยาทั่วไป','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],drugTypes=['ยาทั่วไป','ยาสมุนไพร','ยาโรคNCDs','เวชภัณฑ์ที่มิใช่ยา','วัสดุการแพทย์','วัสดุเภสัช'],out=[];
    let updated=0,skipped=0;
    rows.forEach(p=>{
      const no=String(p.no||'').trim(),unit=String(p.unit||'').trim();
      if(!no||!p.name||!p.form){skipped++;return}
      if(byNo.has(no)){
        const row=byNo.get(no);
        sh.getRange(row,3,1,2).setValues([[p.form||'',p.strength||'']]);
        if(unit)sh.getRange(row,14).setValue(unit);
        const existingType=drugTypes.includes(p.stockType)?p.stockType:(categories.includes(p.category)&&p.category!=='ยาทั่วไป'?p.category:inferStockType_(p));
        sh.getRange(row,8).setValue(existingType);
        updated++;return;
      }
      const requested=p.category==='วชย.'?'เวชภัณฑ์ที่มิใช่ยา':p.category,category=categories.includes(requested)?requested:'ยาทั่วไป';
      const stockType=drugTypes.includes(p.stockType)?p.stockType:inferStockType_(p),usage=p.usage==='ไม่มีใช้'?'ไม่มีใช้':'มีใช้';
      out.push([no,p.name,p.form,p.strength||'',p.account||'',p.reserve||'y',category,stockType,usage,'','','','',unit]);
      byNo.set(no,sh.getLastRow()+out.length);
    });
    if(out.length)sh.getRange(sh.getLastRow()+1,1,out.length,14).setValues(out);
    return{added:out.length+updated,created:out.length,updated,skipped};
  }finally{lock.releaseLock()}
}

/* LINE OA: แจ้งเตือนล็อตคงคลังที่หมดอายุหรือเหลืออายุไม่เกิน 3 เดือน
 * ตั้งค่าใน Apps Script > Project Settings > Script Properties:
 * LINE_CHANNEL_ACCESS_TOKEN และ LINE_ALERT_USER_ID
 */
const LINE_EXPIRY_HANDLER_='sendExpiryAlertsToLine';
const LINE_EXPIRY_TOKEN_KEY_='LINE_CHANNEL_ACCESS_TOKEN';
const LINE_EXPIRY_USER_KEY_='LINE_ALERT_USER_ID';

function installDailyExpiryAlertTrigger(){
  try{
    lineExpiryConfig_();
    ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===LINE_EXPIRY_HANDLER_).forEach(t=>ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger(LINE_EXPIRY_HANDLER_).timeBased().everyDays(1).atHour(8).create();
    return 'ติดตั้งการตรวจวันหมดอายุทุกวันช่วงเวลา 08:00–09:00 น. แล้ว';
  }catch(err){
    const message=String(err&&err.message||err);
    if(/ScriptApp|script\.scriptapp|อนุญาต|permission|authorization/i.test(message))throw new Error('ต้องอนุญาตสิทธิ์สร้าง Trigger ก่อน: เปิดหน้า Apps Script เลือกฟังก์ชัน installDailyExpiryAlertTrigger แล้วกด Run 1 ครั้ง จากนั้นอนุญาตสิทธิ์ทั้งหมด');
    throw err;
  }
}

function removeDailyExpiryAlertTrigger(){
  let removed=0;
  ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===LINE_EXPIRY_HANDLER_).forEach(t=>{ScriptApp.deleteTrigger(t);removed++});
  return 'ลบทริกเกอร์แจ้งเตือน '+removed+' รายการแล้ว';
}

function sendExpiryAlertsToLine(){return sendExpiryAlertsToLine_(false)}
function testLineExpiryAlert(){return sendExpiryAlertsToLine_(true)}

function lineAlertStatus_(){
  const props=PropertiesService.getScriptProperties(),hasToken=!!String(props.getProperty(LINE_EXPIRY_TOKEN_KEY_)||'').trim(),hasUser=!!String(props.getProperty(LINE_EXPIRY_USER_KEY_)||'').trim();
  const triggers=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()===LINE_EXPIRY_HANDLER_).length;
  return {configured:hasToken&&hasUser,hasToken:hasToken,hasUser:hasUser,triggers:triggers,lastSentAt:props.getProperty('LINE_EXPIRY_LAST_SENT_AT')||'',lastDate:props.getProperty('LINE_EXPIRY_LAST_SENT_DATE')||''};
}

function sendExpiryAlertsToLine_(force){
  const config=lineExpiryConfig_(),now=new Date(),tz='Asia/Bangkok';
  const today=new Date(now.getFullYear(),now.getMonth(),now.getDate());
  const threeMonths=new Date(today);threeMonths.setMonth(threeMonths.getMonth()+3);
  const masters=master_(),masterByKey=new Map(masters.map(x=>[String(x.key),x]));
  const alerts=rows_(SHEETS.INVENTORY).map(lot=>{
    const expiry=parseExpiryDate_(lot.expiry),balance=num_(lot.balance);
    if(!expiry||balance<=0||expiry>threeMonths)return null;
    const item=masterByKey.get(String(lot.drugKey))||{},days=Math.floor((expiry-today)/86400000);
    return {item:item,lot:lot,expiry:expiry,balance:balance,days:days,expired:expiry<today};
  }).filter(Boolean).sort((a,b)=>a.expiry-b.expiry||String(a.item.name||'').localeCompare(String(b.item.name||''),'th'));
  const dateKey=Utilities.formatDate(today,tz,'yyyy-MM-dd'),props=PropertiesService.getScriptProperties();
  const digest=Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(alerts.map(x=>[x.lot.drugKey,x.lot.lotNo,Utilities.formatDate(x.expiry,tz,'yyyy-MM-dd'),x.balance]))));
  if(!force&&props.getProperty('LINE_EXPIRY_LAST_SENT_DATE')===dateKey&&props.getProperty('LINE_EXPIRY_LAST_DIGEST')===digest)return {sent:false,reason:'ส่งรายการชุดนี้แล้วในวันนี้',count:alerts.length};
  const messages=buildExpiryLineMessages_(alerts,today,force);
  linePushMessages_(config,messages);
  props.setProperties({LINE_EXPIRY_LAST_SENT_DATE:dateKey,LINE_EXPIRY_LAST_DIGEST:digest,LINE_EXPIRY_LAST_SENT_AT:new Date().toISOString()});
  return {sent:true,count:alerts.length,expired:alerts.filter(x=>x.expired).length,messages:messages.length};
}

function lineExpiryConfig_(){
  const props=PropertiesService.getScriptProperties(),token=String(props.getProperty(LINE_EXPIRY_TOKEN_KEY_)||'').trim(),userId=String(props.getProperty(LINE_EXPIRY_USER_KEY_)||'').trim();
  if(!token||!userId)throw new Error('กรุณาตั้ง Script Properties: LINE_CHANNEL_ACCESS_TOKEN และ LINE_ALERT_USER_ID ก่อน');
  return {token:token,userId:userId};
}

function parseExpiryDate_(value){
  if(value instanceof Date&&!isNaN(value))return new Date(value.getFullYear(),value.getMonth(),value.getDate());
  const s=String(value||'').trim();if(!s)return null;
  let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/),y,mo,d;
  if(m){y=+m[1];mo=+m[2];d=+m[3]}else{m=s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);if(!m)return null;d=+m[1];mo=+m[2];y=+m[3];if(y>2400)y-=543}
  const result=new Date(y,mo-1,d);return result.getFullYear()===y&&result.getMonth()===mo-1&&result.getDate()===d?result:null;
}

function buildExpiryLineMessages_(alerts,today,force){
  const tz='Asia/Bangkok',dateText=Utilities.formatDate(today,tz,'dd/MM/')+(today.getFullYear()+543),expired=alerts.filter(x=>x.expired).length;
  const header='แจ้งเตือนวันหมดอายุคลังยาและเวชภัณฑ์\n'+(CURRENT_FACILITY_&&CURRENT_FACILITY_.name||'หน่วยบริการ')+'\nวันที่ '+dateText+'\nหมดอายุแล้ว '+expired+' ล็อต | เหลือไม่เกิน 3 เดือน '+(alerts.length-expired)+' ล็อต';
  if(!alerts.length)return [header+'\n\n✅ '+(force?'ทดสอบสำเร็จ — ':'')+'ไม่พบรายการที่ต้องแจ้งเตือน'];
  const lines=alerts.map((x,i)=>{
    const item=x.item||{},detail=[item.name||String(x.lot.drugKey||'ไม่ทราบชื่อ'),item.form,item.strength].filter(Boolean).join(' · ');
    const type=item.stockType||item.category||'ไม่ระบุประเภท',status=x.expired?'🔴🔴 หมดอายุแล้ว '+Math.abs(x.days)+' วัน':'🔴 เหลือ '+x.days+' วัน (ไม่เกิน 3 เดือน)';
    return (i+1)+'. '+status+'\n'+detail+'\nประเภท: '+type+' | Lot: '+(x.lot.lotNo||'-')+'\nหมดอายุ: '+Utilities.formatDate(x.expiry,tz,'dd/MM/yyyy')+' | คงเหลือ: '+x.balance+' '+(x.lot.unit||item.unit||'');
  });
  const chunks=[],limit=3900;let current=header;
  lines.forEach(line=>{const next=current+'\n\n'+line;if(next.length>limit){chunks.push(current);current=line}else current=next});if(current)chunks.push(current);
  return chunks;
}

function linePushMessages_(config,texts){
  for(let i=0;i<texts.length;i+=5){
    const response=UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+config.token},payload:JSON.stringify({to:config.userId,messages:texts.slice(i,i+5).map(text=>({type:'text',text:text}))}),muteHttpExceptions:true});
    const code=response.getResponseCode();if(code<200||code>=300)throw new Error('LINE Messaging API ส่งไม่สำเร็จ (HTTP '+code+'): '+response.getContentText().slice(0,500));
  }
}

/* แจ้งเตือนทันทีหลังบันทึกรับเข้า โดยไม่ทำให้การบันทึกหลักล้มเหลวเมื่อ LINE ขัดข้อง */
function notifyInboundExpiryToLine_(p){
  try{
    const expiry=parseExpiryDate_(p.expiry);if(!expiry)return {sent:false,reason:'ไม่มีวันหมดอายุ'};
    const now=new Date(),today=new Date(now.getFullYear(),now.getMonth(),now.getDate()),limit=new Date(today);limit.setMonth(limit.getMonth()+3);
    if(expiry>limit)return {sent:false,reason:'อายุคงเหลือเกิน 3 เดือน'};
    const item=master_().find(x=>String(x.key)===String(p.drugKey))||{};
    const lot=rows_(SHEETS.INVENTORY).find(x=>String(x.drugKey)===String(p.drugKey)&&String(x.lotNo||'')===String(p.lotNo||''))||{};
    const balance=num_(lot.balance)||num_(p.qty),days=Math.floor((expiry-today)/86400000),expired=expiry<today,tz='Asia/Bangkok';
    const detail=[item.name||String(p.drugKey||'ไม่ทราบชื่อ'),item.form,item.strength].filter(Boolean).join(' · ');
    const status=expired?'🔴🔴 หมดอายุแล้ว '+Math.abs(days)+' วัน':'🔴 เหลือ '+days+' วัน (ไม่เกิน 3 เดือน)';
    const text='แจ้งเตือนรับเข้ารายการใกล้หมดอายุ\n'+(CURRENT_FACILITY_&&CURRENT_FACILITY_.name||'หน่วยบริการ')+'\n\n'+status+'\n'+detail+'\nประเภท: '+(item.stockType||item.category||'ไม่ระบุ')+' | Lot: '+(p.lotNo||'-')+'\nหมดอายุ: '+Utilities.formatDate(expiry,tz,'dd/MM/yyyy')+' | คงเหลือ: '+balance+' '+(p.unit||lot.unit||'')+'\nผู้บันทึก: '+(p.recorder||'-');
    linePushMessages_(lineExpiryConfig_(),[text]);
    return {sent:true,expired:expired,days:days};
  }catch(err){
    console.error('Immediate LINE expiry alert failed: '+String(err&&err.stack||err));
    return {sent:false,error:String(err&&err.message||err)};
  }
}
