// TSShourtracker backend: Google Apps Script bound to a Google Sheet.
// Deploy > New deployment > Web app, Execute as: Me, Access: Anyone. Paste the /exec URL into index.html (API const).
const TZ='America/New_York';
// Per-employee settings. to/cc = notification recipients; weekReport = include week pay + take-home estimate; shiftPay = include this shift's pay.
const EMP={
 '7340':{name:'William Griffin',email:'william@techsourcedsolutions.com',rate:30,to:'william@techsourcedsolutions.com',cc:'taylor@techsourcedsolutions.com,austin@techsourcedsolutions.com',weekReport:true,shiftPay:false},
 '1234':{name:'Paul Sorenson',email:'paulsor3@gmail.com',rate:30,to:'paulsor3@gmail.com',cc:'austin@techsourcedsolutions.com',weekReport:true,shiftPay:true}};
const JOBS=['WT Moore'];
function sh(){const ss=SpreadsheetApp.getActive();let s=ss.getSheetByName('Punches');
 if(!s){s=ss.insertSheet('Punches');s.appendRow(['code','name','job','in_iso','out_iso','hours','photos','photo_folder']);}return s;}
function rows(code){return sh().getDataRange().getValues().slice(1).map((r,i)=>({row:i+2,code:String(r[0]),name:r[1],job:r[2],inT:r[3],outT:r[4],hours:r[5]})).filter(r=>r.code===code);}
function fmt(d){return Utilities.formatDate(new Date(d),TZ,'EEE MMM d, h:mm a')+' ET';}
function weekKey(d){const day=Number(Utilities.formatDate(new Date(d),TZ,'u'));const t=new Date(new Date(d).getTime()-(day-1)*864e5);return Utilities.formatDate(t,TZ,'yyyy-MM-dd');}
function summary(code){const e=EMP[code],rs=rows(code),weeks={};
 rs.filter(r=>r.outT).forEach(r=>{const k=weekKey(r.inT);(weeks[k]=weeks[k]||{week:k,hours:0,shifts:[]});weeks[k].hours+=Number(r.hours);weeks[k].shifts.push({job:r.job,in:fmt(r.inT),out:fmt(r.outT),hours:+Number(r.hours).toFixed(2)});});
 const list=Object.values(weeks).sort((a,b)=>a.week<b.week?1:-1).map(w=>{const reg=Math.min(40,w.hours),ot=Math.max(0,w.hours-40);
  return {...w,hours:+w.hours.toFixed(2),regular:+reg.toFixed(2),overtime:+ot.toFixed(2),pay:+(reg*e.rate+ot*e.rate*1.5).toFixed(2)};});
 const open=rs.find(r=>!r.outT);return {name:e.name,rate:e.rate,jobs:JOBS,open:open?{job:open.job,in:fmt(open.inT),inIso:String(open.inT)}:null,weeks:list};}

// ===== Take-home ESTIMATE settings (easy to change) =====
const TAX={
  payPeriodsPerYear:52,            // weekly
  // Federal: IRS Pub 15-T (2026) Percentage Method, 2020+ W-4, Single, Step 2 box NOT checked, no Step 3/4 entries
  fedW4Line1g:8600,                // Worksheet 1A line 1g (single, box unchecked)
  fedSingleStd:[[0,0,0],[7500,0,0.10],[19900,1240,0.12],[57900,5800,0.22],[113200,17966,0.24],[209275,41024,0.32],[263725,58448,0.35],[648100,192979.25,0.37]], // [over, base, rate]
  ssRate:0.062, medicareRate:0.0145,
  // Georgia: 2026 Employer's Tax Guide percentage method, single, 0 allowances
  gaStdDeduction:15000, gaRate:0.0499
};
function estimateTakeHome(gross){const n=TAX.payPeriodsPerYear,a=Math.max(0,gross*n-TAX.fedW4Line1g);let row=TAX.fedSingleStd[0];
 TAX.fedSingleStd.forEach(r=>{if(a>=r[0])row=r;});const fed=(row[1]+(a-row[0])*row[2])/n;
 const ss=gross*TAX.ssRate,med=gross*TAX.medicareRate,ga=Math.max(0,gross*n-TAX.gaStdDeduction)*TAX.gaRate/n;
 const r=x=>+x.toFixed(2);return {fed:r(fed),ss:r(ss),med:r(med),ga:r(ga),net:r(gross-fed-ss-med-ga)};}
function weekReport(code,d){const w=summary(code).weeks.find(x=>x.week===weekKey(d));if(!w)return'';const t=estimateTakeHome(w.pay),e=EMP[code];
 return '\n\n--- This week (Mon '+w.week+') so far ---\nRegular: '+w.regular+' h x $'+e.rate+' = $'+(w.regular*e.rate).toFixed(2)+
 '\nOvertime: '+w.overtime+' h x $'+(e.rate*1.5)+' = $'+(w.overtime*e.rate*1.5).toFixed(2)+'\nGross pay: $'+w.pay.toFixed(2)+
 '\n\nESTIMATED take-home (estimate only, not a pay stub):\nFederal income tax: -$'+t.fed+'\nSocial Security 6.2%: -$'+t.ss+'\nMedicare 1.45%: -$'+t.med+'\nGeorgia income tax: -$'+t.ga+'\nEstimated net: $'+t.net+
 '\nAssumptions: weekly pay, single, standard deduction, no extra W-4/G-4 adjustments; 2026 IRS Pub 15-T percentage method; GA 4.99% after $15,000 std deduction.';}

// ===== Photos (Drive). Needs Drive scope: re-authorize when deploying this version. =====
function sub(parent,name){const it=parent.getFoldersByName(name);return it.hasNext()?it.next():parent.createFolder(name);}
function sheetParent(){const f=DriveApp.getFileById(SpreadsheetApp.getActive().getId()).getParents();return f.hasNext()?f.next():DriveApp.getRootFolder();}
function mastersSheet(){const ss=SpreadsheetApp.getActive();let m=ss.getSheetByName('Masters');if(!m){m=ss.insertSheet('Masters');m.appendRow(['job','folder_id','folder_name']);}return m;}
// One master folder per job, next to the Sheet. Job -> folder ID kept in Masters tab so renames are reused.
function masterFolder(job){const m=mastersSheet(),v=m.getDataRange().getValues();
 for(let i=1;i<v.length;i++)if(v[i][0]===job){try{const f=DriveApp.getFolderById(v[i][1]);if(!f.isTrashed())return f;}catch(x){}
  const f=sub(sheetParent(),v[i][2]||job);m.getRange(i+1,2,1,2).setValues([[f.getId(),f.getName()]]);return f;}
 const f=sub(sheetParent(),job);m.appendRow([job,f.getId(),f.getName()]);return f;}
function cleanName(n){return String(n||'').replace(/[\\/:*?"<>|]/g,'').trim().slice(0,80);}
function renameMaster(job,name){if(JOBS.indexOf(job)<0)return {error:'Unknown job'};name=cleanName(name);if(!name)return {error:'Enter a name'};
 const f=masterFolder(job);f.setName(name);const m=mastersSheet(),v=m.getDataRange().getValues();
 for(let i=1;i<v.length;i++)if(v[i][0]===job)m.getRange(i+1,3).setValue(name);return {ok:true,master:name,url:f.getUrl()};}
function sessionRow(code,inIso){return rows(code).find(r=>String(r.inT)===String(inIso));}
function uploadPhoto(p,e){const r=sessionRow(p.code,p.inIso);if(!r)return {error:'Unknown session'};
 const room=String(p.room).replace(/[^\w .-]/g,'').slice(0,40)||'0';const sf=masterFolder(r.job),rf=sub(sf,'Room '+room);
 const name='Room '+room+' '+Utilities.formatDate(new Date(p.takenAt),TZ,'yyyy-MM-dd HH-mm-ss')+(p.seq?' '+p.seq:'')+'.jpg';
 if(rf.getFilesByName(name).hasNext())return {ok:true,dup:true};// idempotent retry
 rf.createFile(Utilities.newBlob(Utilities.base64Decode(p.data),'image/jpeg',name));
 const c=sh().getRange(r.row,7,1,2);const v=c.getValues()[0];c.setValues([[Number(v[0]||0)+1,sf.getUrl()]]);return {ok:true};}
function shiftPayReport(code,inT,h){const e=EMP[code],w=summary(code).weeks.find(x=>x.week===weekKey(inT));const prior=w?Math.max(0,w.hours-h):0;
 const reg=Math.max(0,Math.min(h,40-prior)),ot=h-reg,r=e.rate;
 return '\n\n--- Pay earned this shift ---\nRegular: '+reg.toFixed(2)+' h x $'+r+' = $'+(reg*r).toFixed(2)+(ot>0?'\nOvertime: '+ot.toFixed(2)+' h x $'+(r*1.5)+' = $'+(ot*r*1.5).toFixed(2):'')+'\nShift pay: $'+(reg*r+ot*r*1.5).toFixed(2);}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doPost(ev){const lock=LockService.getScriptLock();lock.waitLock(10000);try{
 const p=JSON.parse(ev.postData.contents),e=EMP[p.code];if(!e)return out({error:'Invalid code'});
 if(p.action==='master')return out(p.name?renameMaster(p.job,p.name):(()=>{const f=masterFolder(p.job);return {master:f.getName(),url:f.getUrl()};})());
 if(p.action==='photo')return out(uploadPhoto(p,e));
 if(p.action==='punch'){const open=rows(p.code).find(r=>!r.outT),now=new Date();
  if(!open){if(JOBS.indexOf(p.job)<0)return out({error:'Pick a job'});
   sh().appendRow([p.code,e.name,p.job,now.toISOString(),'','']);
   MailApp.sendEmail(e.to,'TSShourtracker: Clocked in – '+p.job,e.name+' clocked in\nJob: '+p.job+'\nTime: '+fmt(now),{cc:e.cc});}
  else{const h=(now-new Date(open.inT))/36e5,np=Number(p.photos||0);let pl='';if(np>0){try{pl='\nPhotos: '+np+' uploading to '+masterFolder(open.job).getUrl();}catch(x){pl='\nPhotos: '+np+' uploading';}}sh().getRange(open.row,5,1,2).setValues([[now.toISOString(),h.toFixed(4)]]);
   MailApp.sendEmail(e.to,'TSShourtracker: Clocked out – '+open.job,e.name+' clocked out\nJob: '+open.job+'\nIn: '+fmt(open.inT)+'\nOut: '+fmt(now)+'\nShift hours: '+h.toFixed(2)+pl+(e.shiftPay?shiftPayReport(p.code,open.inT,h):'')+(e.weekReport?weekReport(p.code,open.inT):''),{cc:e.cc});}}
 return out(summary(p.code));}finally{lock.releaseLock();}}
