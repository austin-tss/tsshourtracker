// TSShourtracker backend: Google Apps Script bound to a Google Sheet.
// Deploy > New deployment > Web app, Execute as: Me, Access: Anyone. Paste the /exec URL into index.html (API const).
const TZ='America/New_York';
const EMP={'7340':{name:'William Griffin',email:'william@techsourcedsolutions.com',rate:30}};
const JOBS=['WT Moore'];
function sh(){const ss=SpreadsheetApp.getActive();let s=ss.getSheetByName('Punches');
 if(!s){s=ss.insertSheet('Punches');s.appendRow(['code','name','job','in_iso','out_iso','hours']);}return s;}
function rows(code){return sh().getDataRange().getValues().slice(1).map((r,i)=>({row:i+2,code:String(r[0]),name:r[1],job:r[2],inT:r[3],outT:r[4],hours:r[5]})).filter(r=>r.code===code);}
function fmt(d){return Utilities.formatDate(new Date(d),TZ,'EEE MMM d, h:mm a')+' ET';}
function weekKey(d){const day=Number(Utilities.formatDate(new Date(d),TZ,'u'));const t=new Date(new Date(d).getTime()-(day-1)*864e5);return Utilities.formatDate(t,TZ,'yyyy-MM-dd');}
function summary(code){const e=EMP[code],rs=rows(code),weeks={};
 rs.filter(r=>r.outT).forEach(r=>{const k=weekKey(r.inT);(weeks[k]=weeks[k]||{week:k,hours:0,shifts:[]});weeks[k].hours+=Number(r.hours);weeks[k].shifts.push({job:r.job,in:fmt(r.inT),out:fmt(r.outT),hours:+Number(r.hours).toFixed(2)});});
 const list=Object.values(weeks).sort((a,b)=>a.week<b.week?1:-1).map(w=>{const reg=Math.min(40,w.hours),ot=Math.max(0,w.hours-40);
  return {...w,hours:+w.hours.toFixed(2),regular:+reg.toFixed(2),overtime:+ot.toFixed(2),pay:+(reg*e.rate+ot*e.rate*1.5).toFixed(2)};});
 const open=rs.find(r=>!r.outT);return {name:e.name,rate:e.rate,jobs:JOBS,open:open?{job:open.job,in:fmt(open.inT)}:null,weeks:list};}
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function doPost(ev){const lock=LockService.getScriptLock();lock.waitLock(10000);try{
 const p=JSON.parse(ev.postData.contents),e=EMP[p.code];if(!e)return out({error:'Invalid code'});
 if(p.action==='punch'){const open=rows(p.code).find(r=>!r.outT),now=new Date();
  if(!open){if(JOBS.indexOf(p.job)<0)return out({error:'Pick a job'});
   sh().appendRow([p.code,e.name,p.job,now.toISOString(),'','']);
   MailApp.sendEmail(e.email,'TSShourtracker: Clocked in – '+p.job,e.name+' clocked in\nJob: '+p.job+'\nTime: '+fmt(now));}
  else{const h=(now-new Date(open.inT))/36e5;sh().getRange(open.row,5,1,2).setValues([[now.toISOString(),h.toFixed(4)]]);
   MailApp.sendEmail(e.email,'TSShourtracker: Clocked out – '+open.job,e.name+' clocked out\nJob: '+open.job+'\nIn: '+fmt(open.inT)+'\nOut: '+fmt(now)+'\nShift hours: '+h.toFixed(2));}}
 return out(summary(p.code));}finally{lock.releaseLock();}}
