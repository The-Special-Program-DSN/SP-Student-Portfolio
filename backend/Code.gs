const SPREADSHEET_ID = '1RjNMXWVXnrDEalB9Nqx10mHm2GZq0dq7rc_K_UGuUAY';
const SESSION_TTL = 21600;

const SHEETS = {
  Config: ['key','value'],
  Students: ['id','citizen','dob','title','first_name','last_name','grade','room','student_no','program','registered_at','registered_academic_year','expiry_academic_year','status','graduated','graduation_stage','continued_m4_at','continued_m4_academic_year','updated_at'],
  Activities: ['record_id','student_id','citizen_id','title','first_name','last_name','program_title','exp_name','description','date','end_date','year','level','hours','fee','evidence_url','created_at','updated_at'],
  CertsCourses: ['record_id','student_id','citizen_id','title','first_name','last_name','course_name','course_level','description','issue_date','expired_date','score','year','category','level','hours','fee','reflection','evidence_url','created_at','updated_at'],
  Prize: ['record_id','student_id','citizen_id','title','first_name','last_name','program_title','prize_name','description','date','end_date','year','level','hours','fee','evidence_url','created_at','updated_at'],
  Projects: ['record_id','student_id','citizen_id','title','first_name','last_name','project_title','project_type','description','date','end_date','year','level','hours','fee','evidence_url','created_at','updated_at'],
  YearSnapshots: ['batch_id','from_year','to_year','student_id','grade','graduated','graduation_stage','status','expiry_academic_year','created_at'],
  AuditLog: ['at','actor_type','actor_id','action','detail']
};

function doGet() {
  return jsonResponse_({ok:true, service:'SP Student Portfolio API', academicYear:getAcademicYear_()});
}

function doPost(e) {
  try {
    ensureDatabase_();
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(body.action || '');
    const args = Array.isArray(body.args) ? body.args : [];
    return jsonResponse_({ok:true, result:dispatchApi_(action,args)});
  } catch (err) {
    return jsonResponse_({ok:false, error:(err && err.message) ? err.message : String(err)});
  }
}

function jsonResponse_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function dispatchApi_(action,args){
  const allowed = {
    getPublicConfig, registerStudent, studentLogin, adminLogin, getStudentHome, getStudentRecords,
    saveStudentRecord, continueM4, updateStudentProfile, getAdminDashboard, adminListStudents,
    adminUpdateStudent, adminDeleteStudents, adminListRecords, adminAdvanceYear, adminUndoYear,
    adminYearHistory
  };
  if(!allowed[action]) throw new Error('API action ไม่ได้รับอนุญาต');
  return allowed[action].apply(null,args);
}

function secret_(key){
  const value = PropertiesService.getScriptProperties().getProperty(key);
  if(!value) throw new Error('ยังไม่ได้ตั้งค่า Script Property: ' + key);
  return value;
}

function ensureDatabase_(){
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.entries(SHEETS).forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(headers);
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach((h,i)=>{ if(current[i]!==h) sh.getRange(1,i+1).setValue(h); });
    sh.setFrozenRows(1);
  });
  const cfg = ss.getSheetByName('Config');
  const map = getConfigMap_();
  if (!map.academicYear) cfg.appendRow(['academicYear','2569']);
  return true;
}

function getPublicConfig(){ ensureDatabase_(); return { academicYear:getAcademicYear_(), appName:'SP Student Portfolio' }; }

function registerStudent(payload){
  ensureDatabase_();
  payload = payload || {};
  const citizen = clean_(payload.citizen);
  const dob = clean_(payload.dob);
  if (!/^\d{13}$/.test(citizen)) throw new Error('กรุณากรอกเลขบัตรประชาชน 13 หลัก');
  if (!/^\d{8}$/.test(dob)) throw new Error('กรุณากรอกวันเดือนปีเกิด 8 หลัก');
  if (!clean_(payload.first) || !clean_(payload.last)) throw new Error('กรุณากรอกชื่อและนามสกุล');
  const students = getRows_('Students');
  if (students.some(s => String(s.citizen) === citizen)) throw new Error('เลขบัตรประชาชนนี้มีบัญชีอยู่แล้ว');
  const year = getAcademicYear_();
  const grade = Number(payload.grade);
  if (![1,2,3,4,5,6].includes(grade)) throw new Error('ระดับชั้นไม่ถูกต้อง');
  const now = new Date().toISOString();
  const id = Utilities.getUuid();
  const row = {
    id, citizen, dob,
    title:clean_(payload.title), first_name:clean_(payload.first), last_name:clean_(payload.last),
    grade, room:clean_(payload.room), student_no:clean_(payload.no), program:clean_(payload.program),
    registered_at:now, registered_academic_year:year,
    expiry_academic_year:initialExpiry_(year, grade), status:'active', graduated:false, graduation_stage:'',
    continued_m4_at:'', continued_m4_academic_year:'', updated_at:now
  };
  appendObject_('Students',row);
  log_('student',id,'register',JSON.stringify({citizen,grade,year}));
  const token = createSession_('student',id);
  return {ok:true, token, student:studentView_(row), academicYear:year};
}

function studentLogin(citizen,dob){
  ensureDatabase_();
  citizen = clean_(citizen); dob = clean_(dob);
  const row = getRows_('Students').find(s => String(s.citizen)===citizen && String(s.dob)===dob);
  if (!row) throw new Error('ข้อมูลเข้าสู่ระบบไม่ถูกต้อง');
  const year = getAcademicYear_();
  if (isExpired_(row,year)) throw new Error('บัญชีหมดอายุ กรุณาติดต่อฝ่ายแนะแนว');
  const token = createSession_('student',row.id);
  return {ok:true, token, student:studentView_(row), academicYear:year};
}

function adminLogin(user,pass){
  ensureDatabase_();
  const adminUser=secret_('ADMIN_USER');
  const adminPassword=secret_('ADMIN_PASSWORD');
  if (String(user)!==adminUser || String(pass)!==adminPassword) throw new Error('Username หรือ Password ไม่ถูกต้อง');
  const token = createSession_('admin',adminUser);
  return {ok:true, token, academicYear:getAcademicYear_()};
}

function getStudentHome(token){
  const sess = requireSession_(token,'student');
  const s = findStudentById_(sess.id);
  if (!s) throw new Error('ไม่พบบัญชีนักเรียน');
  return {student:studentView_(s), academicYear:getAcademicYear_(), counts:getStudentCounts_(s.id)};
}

function getStudentRecords(token, category){
  const sess = requireSession_(token,'student');
  return getCategoryRows_(category).filter(r => String(r.student_id)===String(sess.id)).map(recordView_).reverse();
}

function saveStudentRecord(token, category, payload){
  const sess = requireSession_(token,'student');
  const s = findStudentById_(sess.id);
  if (!s) throw new Error('ไม่พบบัญชีนักเรียน');
  payload = payload || {};
  const year = getAcademicYear_();
  const now = new Date().toISOString();
  const base = {
    record_id:Utilities.getUuid(), student_id:s.id, citizen_id:s.citizen, title:s.title,
    first_name:s.first_name, last_name:s.last_name, year:Number(payload.year)||year,
    level:clean_(payload.level), hours:clean_(payload.hours), fee:clean_(payload.fee),
    evidence_url:clean_(payload.evidenceUrl), created_at:now, updated_at:now
  };
  let sheet,row;
  if(category==='activities'){
    sheet='Activities'; row={...base, program_title:clean_(payload.programTitle), exp_name:clean_(payload.expName), description:clean_(payload.description), date:clean_(payload.date), end_date:clean_(payload.endDate)};
    if(!row.program_title) throw new Error('กรุณากรอกชื่อกิจกรรม / โครงการ');
  } else if(category==='courses'){
    sheet='CertsCourses'; row={...base, course_name:clean_(payload.courseName), course_level:clean_(payload.courseLevel), description:clean_(payload.description), issue_date:clean_(payload.issueDate), expired_date:clean_(payload.expiredDate), score:clean_(payload.score), category:clean_(payload.courseCategory), reflection:clean_(payload.reflection)};
    if(!row.course_name) throw new Error('กรุณากรอกชื่อรายวิชา / หลักสูตร');
  } else if(category==='prize'){
    sheet='Prize'; row={...base, program_title:clean_(payload.programTitle), prize_name:clean_(payload.prizeName), description:clean_(payload.description), date:clean_(payload.date), end_date:clean_(payload.endDate)};
    if(!row.program_title || !row.prize_name) throw new Error('กรุณากรอกชื่อกิจกรรมและรางวัล');
  } else if(category==='projects'){
    sheet='Projects'; row={...base, project_title:clean_(payload.projectTitle), project_type:clean_(payload.projectType), description:clean_(payload.description), date:clean_(payload.date), end_date:clean_(payload.endDate)};
    if(!row.project_title) throw new Error('กรุณากรอกชื่อโครงงาน');
  } else throw new Error('ประเภทข้อมูลไม่ถูกต้อง');
  appendObject_(sheet,row);
  log_('student',s.id,'add-record',JSON.stringify({category,recordId:row.record_id}));
  return {ok:true, record:recordView_(row)};
}

function continueM4(token, phrase){
  const sess = requireSession_(token,'student');
  if(clean_(phrase)!=='ศึกษาต่อ ม.4 โครงการห้องเรียนพิเศษ') throw new Error('ข้อความยืนยันไม่ถูกต้อง');
  const rows = getRowsWithIndex_('Students');
  const item = rows.find(x=>String(x.data.id)===String(sess.id));
  if(!item) throw new Error('ไม่พบบัญชีนักเรียน');
  const s=item.data;
  if(!(toBool_(s.graduated) && String(s.graduation_stage)==='m3')) throw new Error('บัญชีนี้ไม่อยู่ในสถานะจบ ม.3');
  const year=getAcademicYear_(), now=new Date().toISOString();
  s.grade=4; s.graduated=false; s.graduation_stage=''; s.status='active';
  s.expiry_academic_year=Math.max(Number(s.expiry_academic_year)||0, year+4);
  s.continued_m4_at=now; s.continued_m4_academic_year=year; s.updated_at=now;
  writeObjectAtRow_('Students',item.row,s);
  log_('student',s.id,'continue-m4',JSON.stringify({year,expiry:s.expiry_academic_year}));
  return {ok:true, student:studentView_(s), academicYear:year};
}

function updateStudentProfile(token,payload){
  const sess=requireSession_(token,'student');
  const rows=getRowsWithIndex_('Students'); const item=rows.find(x=>String(x.data.id)===String(sess.id));
  if(!item) throw new Error('ไม่พบบัญชีนักเรียน');
  const s=item.data;
  s.room=clean_(payload.room!==undefined?payload.room:s.room); s.student_no=clean_(payload.no!==undefined?payload.no:s.student_no); s.program=clean_(payload.program!==undefined?payload.program:s.program);
  s.updated_at=new Date().toISOString(); writeObjectAtRow_('Students',item.row,s);
  return {ok:true,student:studentView_(s)};
}

function getAdminDashboard(token){ requireSession_(token,'admin'); const year=getAcademicYear_(); const students=getRows_('Students');
  return {academicYear:year,totalStudents:students.length,active:students.filter(s=>!isExpired_(s,year)).length,expired:students.filter(s=>isExpired_(s,year)).length,counts:{activities:getRows_('Activities').length,courses:getRows_('CertsCourses').length,prize:getRows_('Prize').length,projects:getRows_('Projects').length}};
}

function adminListStudents(token, query, grade, expiredOnly){
  requireSession_(token,'admin'); const year=getAcademicYear_(); query=clean_(query).toLowerCase(); grade=Number(grade)||0;
  return getRows_('Students').filter(s=>{
    const expired=isExpired_(s,year); if(expiredOnly && !expired) return false;
    if(grade && Number(s.grade)!==grade) return false;
    if(query){ const hay=[s.citizen,s.first_name,s.last_name,`${s.first_name} ${s.last_name}`].join(' ').toLowerCase(); if(!hay.includes(query)) return false; }
    return true;
  }).sort((a,b)=>String(a.first_name).localeCompare(String(b.first_name),'th')).map(s=>({...studentView_(s),expired:isExpired_(s,year)}));
}

function adminUpdateStudent(token, id, payload){
  const admin=requireSession_(token,'admin'); const rows=getRowsWithIndex_('Students'); const item=rows.find(x=>String(x.data.id)===String(id)); if(!item) throw new Error('ไม่พบบัญชี');
  const s=item.data; const citizen=clean_(payload.citizen), dob=clean_(payload.dob);
  if(!/^\d{13}$/.test(citizen)) throw new Error('เลขบัตรประชาชนต้องมี 13 หลัก'); if(!/^\d{8}$/.test(dob)) throw new Error('วันเดือนปีเกิดต้องมี 8 หลัก');
  if(getRows_('Students').some(x=>String(x.id)!==String(id)&&String(x.citizen)===citizen)) throw new Error('เลขบัตรประชาชนนี้ซ้ำกับบัญชีอื่น');
  Object.assign(s,{citizen,dob,title:clean_(payload.title),first_name:clean_(payload.first),last_name:clean_(payload.last),grade:Number(payload.grade)||s.grade,room:clean_(payload.room),student_no:clean_(payload.no),program:clean_(payload.program),updated_at:new Date().toISOString()});
  writeObjectAtRow_('Students',item.row,s); syncIdentityToRecords_(s);
  log_('admin',admin.id,'edit-student',JSON.stringify({studentId:id})); return {ok:true,student:studentView_(s)};
}

function adminDeleteStudents(token, ids){
  const admin=requireSession_(token,'admin'); ids=(ids||[]).map(String); if(!ids.length) throw new Error('ยังไม่ได้เลือกบัญชี');
  deleteRowsByKey_('Students','id',ids); ['Activities','CertsCourses','Prize','Projects'].forEach(n=>deleteRowsByKey_(n,'student_id',ids));
  log_('admin',admin.id,'delete-students',JSON.stringify({count:ids.length})); return {ok:true,count:ids.length};
}

function adminListRecords(token, category){ requireSession_(token,'admin'); if(category) return getCategoryRows_(category).map(recordView_).reverse();
  return ['activities','courses','prize','projects'].flatMap(c=>getCategoryRows_(c).map(r=>({...recordView_(r),category:c}))).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
}