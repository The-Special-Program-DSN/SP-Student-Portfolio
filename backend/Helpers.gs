function adminAdvanceYear(token, code){
  const admin=requireSession_(token,'admin'); if(String(code)!==secret_('YEAR_CHANGE_CODE')) throw new Error('รหัสยืนยันไม่ถูกต้อง');
  const ss=SpreadsheetApp.openById(SPREADSHEET_ID); const year=getAcademicYear_(); const batch=Utilities.getUuid(); const now=new Date().toISOString();
  const lock=LockService.getScriptLock(); lock.waitLock(30000);
  try{
    const rows=getRowsWithIndex_('Students'); const snap=ss.getSheetByName('YearSnapshots');
    rows.forEach(item=>{ const s=item.data; snap.appendRow([batch,year,year+1,s.id,s.grade,s.graduated,s.graduation_stage,s.status,s.expiry_academic_year,now]);
      if(!toBool_(s.graduated)){
        const g=Number(s.grade); if(g===1)s.grade=2; else if(g===2)s.grade=3; else if(g===3){s.graduated=true;s.graduation_stage='m3';} else if(g===4)s.grade=5; else if(g===5)s.grade=6; else if(g===6){s.graduated=true;s.graduation_stage='m6';}
        s.updated_at=now; writeObjectAtRow_('Students',item.row,s);
      }
    });
    setConfig_('academicYear',String(year+1)); log_('admin',admin.id,'advance-year',JSON.stringify({from:year,to:year+1,batch}));
  } finally { lock.releaseLock(); }
  return {ok:true,academicYear:year+1};
}

function adminUndoYear(token, code){
  const admin=requireSession_(token,'admin'); if(String(code)!==secret_('YEAR_CHANGE_CODE')) throw new Error('รหัสยืนยันไม่ถูกต้อง');
  const snaps=getRowsWithIndex_('YearSnapshots'); if(!snaps.length) throw new Error('ไม่มีข้อมูลให้ย้อนกลับ');
  const latest=snaps[snaps.length-1].data; const batch=String(latest.batch_id); const batchRows=snaps.filter(x=>String(x.data.batch_id)===batch);
  const studentRows=getRowsWithIndex_('Students'); const map=new Map(studentRows.map(x=>[String(x.data.id),x]));
  batchRows.forEach(x=>{ const b=x.data, item=map.get(String(b.student_id)); if(!item)return; const s=item.data; s.grade=Number(b.grade); s.graduated=toBool_(b.graduated); s.graduation_stage=clean_(b.graduation_stage); s.status=clean_(b.status)||'active'; s.expiry_academic_year=Number(b.expiry_academic_year)||''; s.updated_at=new Date().toISOString(); writeObjectAtRow_('Students',item.row,s); });
  deleteRowsByKey_('YearSnapshots','batch_id',[batch]); setConfig_('academicYear',String(Number(latest.from_year))); log_('admin',admin.id,'undo-year',JSON.stringify({batch,to:latest.from_year}));
  return {ok:true,academicYear:Number(latest.from_year)};
}

function adminYearHistory(token){ requireSession_(token,'admin'); const rows=getRows_('YearSnapshots'); const grouped={}; rows.forEach(r=>{ const k=r.batch_id; if(!grouped[k]) grouped[k]={batchId:k,fromYear:Number(r.from_year),toYear:Number(r.to_year),at:r.created_at,count:0}; grouped[k].count++; }); return Object.values(grouped).sort((a,b)=>String(b.at).localeCompare(String(a.at))); }

function getAcademicYear_(){ ensureDatabaseBare_(); const map=getConfigMap_(); return Number(map.academicYear)||2569; }
function ensureDatabaseBare_(){ const ss=SpreadsheetApp.openById(SPREADSHEET_ID); if(!ss.getSheetByName('Config')){ const sh=ss.insertSheet('Config'); sh.appendRow(SHEETS.Config); sh.appendRow(['academicYear','2569']); } }
function getConfigMap_(){ const ss=SpreadsheetApp.openById(SPREADSHEET_ID), sh=ss.getSheetByName('Config'); if(!sh||sh.getLastRow()<2)return{}; const vals=sh.getRange(2,1,sh.getLastRow()-1,2).getValues(); const out={}; vals.forEach(r=>out[String(r[0])]=String(r[1])); return out; }
function setConfig_(key,value){ const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Config'); const vals=sh.getRange(1,1,Math.max(sh.getLastRow(),1),2).getValues(); for(let i=1;i<vals.length;i++){ if(String(vals[i][0])===key){sh.getRange(i+1,2).setValue(value);return;} } sh.appendRow([key,value]); }
function initialExpiry_(year,grade){ return Number(year)+(6-Number(grade))+(Number(grade)>=4?2:0); }
function isExpired_(s,year){ return Number(year) > (Number(s.expiry_academic_year)||9999); }
function studentView_(s){ return {id:s.id,citizen:String(s.citizen),dob:String(s.dob),title:s.title,first:s.first_name,last:s.last_name,grade:Number(s.grade),room:s.room,no:s.student_no,program:s.program,registeredAt:s.registered_at,registeredAcademicYear:Number(s.registered_academic_year),expiryAcademicYear:Number(s.expiry_academic_year),status:s.status,graduated:toBool_(s.graduated),graduationStage:s.graduation_stage,continuedM4At:s.continued_m4_at,continuedM4AcademicYear:Number(s.continued_m4_academic_year)||null}; }
function getStudentCounts_(id){ return {activities:getRows_('Activities').filter(r=>String(r.student_id)===String(id)).length,courses:getRows_('CertsCourses').filter(r=>String(r.student_id)===String(id)).length,prize:getRows_('Prize').filter(r=>String(r.student_id)===String(id)).length,projects:getRows_('Projects').filter(r=>String(r.student_id)===String(id)).length}; }
function getCategoryRows_(category){ if(category==='activities')return getRows_('Activities').map(r=>({...r,category})); if(category==='courses')return getRows_('CertsCourses').map(r=>({...r,category})); if(category==='prize')return getRows_('Prize').map(r=>({...r,category})); if(category==='projects')return getRows_('Projects').map(r=>({...r,category})); throw new Error('ประเภทข้อมูลไม่ถูกต้อง'); }
function recordView_(r){ const c=r.category; let title='',extra=''; if(c==='activities'){title=r.program_title;extra=r.exp_name;} if(c==='courses'){title=r.course_name;extra=r.course_level;} if(c==='prize'){title=r.program_title;extra=r.prize_name;} if(c==='projects'){title=r.project_title;extra=r.project_type;} return {id:r.record_id,studentId:r.student_id,citizen:r.citizen_id,studentName:`${r.title||''}${r.first_name||''} ${r.last_name||''}`.trim(),category:c,title,extra,description:r.description||'',date:r.date||r.issue_date||'',endDate:r.end_date||r.expired_date||'',score:r.score||'',year:r.year,courseCategory:r.category||'',level:r.level||'',hours:r.hours||'',fee:r.fee||'',reflection:r.reflection||'',evidenceUrl:r.evidence_url||'',createdAt:r.created_at||''}; }
function findStudentById_(id){ return getRows_('Students').find(s=>String(s.id)===String(id)); }
function syncIdentityToRecords_(s){ const mapping=['Activities','CertsCourses','Prize','Projects']; mapping.forEach(name=>{ const rows=getRowsWithIndex_(name); rows.filter(x=>String(x.data.student_id)===String(s.id)).forEach(x=>{ x.data.citizen_id=s.citizen; x.data.title=s.title; x.data.first_name=s.first_name; x.data.last_name=s.last_name; x.data.updated_at=new Date().toISOString(); writeObjectAtRow_(name,x.row,x.data); }); }); }
function getRows_(sheetName){ return getRowsWithIndex_(sheetName).map(x=>x.data); }
function getRowsWithIndex_(sheetName){ const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName); if(!sh||sh.getLastRow()<2)return[]; const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String); const vals=sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues(); return vals.map((row,i)=>{const o={};headers.forEach((h,j)=>o[h]=row[j]);return{row:i+2,data:o};}); }
function appendObject_(sheetName,obj){ const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName); const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String); sh.appendRow(headers.map(h=>obj[h]!==undefined?obj[h]:'')); }
function writeObjectAtRow_(sheetName,row,obj){ const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName); const headers=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(String); sh.getRange(row,1,1,headers.length).setValues([headers.map(h=>obj[h]!==undefined?obj[h]:'')]); }
function deleteRowsByKey_(sheetName,key,ids){ const sh=SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName); const rows=getRowsWithIndex_(sheetName).filter(x=>ids.includes(String(x.data[key]))).map(x=>x.row).sort((a,b)=>b-a); rows.forEach(r=>sh.deleteRow(r)); }
function createSession_(type,id){ const token=Utilities.getUuid()+Utilities.getUuid(); CacheService.getScriptCache().put('sess:'+token,JSON.stringify({type,id,at:Date.now()}),SESSION_TTL); return token; }
function requireSession_(token,type){ const raw=CacheService.getScriptCache().get('sess:'+String(token||'')); if(!raw) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่'); const s=JSON.parse(raw); if(type&&s.type!==type) throw new Error('ไม่มีสิทธิ์ใช้งาน'); return s; }
function clean_(v){ return String(v===undefined||v===null?'':v).trim(); }
function toBool_(v){ return v===true || String(v).toLowerCase()==='true' || String(v)==='1'; }
function sha256_(s){ const bytes=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8); return bytes.map(b=>('0'+((b<0?b+256:b).toString(16))).slice(-2)).join(''); }
function log_(actorType,actorId,action,detail){ appendObject_('AuditLog',{at:new Date().toISOString(),actor_type:actorType,actor_id:actorId,action,detail}); }