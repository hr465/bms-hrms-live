const express=require("express");
const path=require("path");
const crypto=require("crypto");
const db=require("./db");
const ZKLib=require("node-zklib");
const {sendMail,layout}=require("./mail");
const ExcelJS=require("exceljs");

const app=express();
const PORT=process.env.PORT||3000;
app.set("trust proxy",1);
app.use(express.json({limit:"15mb"}));
app.use(express.urlencoded({extended:true}));
app.use(express.static(path.join(__dirname,"..","public")));

async function initSchema(){
await db.exec(`
CREATE TABLE IF NOT EXISTS companies(
 id SERIAL PRIMARY KEY, name TEXT NOT NULL, code TEXT UNIQUE NOT NULL, industry TEXT,
 address TEXT, contact_email TEXT, contact_phone TEXT, status TEXT DEFAULT 'Active', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
 smtp_user TEXT, smtp_pass TEXT, policy_agreement_text TEXT, increment_policy TEXT
);
CREATE TABLE IF NOT EXISTS users(
 id SERIAL PRIMARY KEY, company_id INTEGER, username TEXT UNIQUE, password_hash TEXT,
 role TEXT, employee_id INTEGER, active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,user_id INTEGER,expires_at BIGINT,active_company_id INTEGER);
CREATE TABLE IF NOT EXISTS employees(
 id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL, employee_code TEXT, name TEXT, email TEXT, phone TEXT,
 department TEXT, designation TEXT, manager TEXT, reporting_manager_id INTEGER, branch TEXT, joining_date TEXT, status TEXT DEFAULT 'Active',
 biometric_id TEXT, salary REAL DEFAULT 0, bank_name TEXT, bank_account TEXT, ifsc TEXT,
 pf_number TEXT, esic_number TEXT, uan_number TEXT, pan_number TEXT,
 date_of_birth TEXT, basic_salary REAL DEFAULT 0, hra REAL DEFAULT 0, other_allowances REAL DEFAULT 0,
 pf_applicable INTEGER DEFAULT 0, esic_applicable INTEGER DEFAULT 0,
 created_at TEXT DEFAULT CURRENT_TIMESTAMP, UNIQUE(company_id,employee_code)
);
CREATE TABLE IF NOT EXISTS departments(id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT,UNIQUE(company_id,name));
CREATE TABLE IF NOT EXISTS leave_types(id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT,annual_balance REAL DEFAULT 0,UNIQUE(company_id,name));
CREATE TABLE IF NOT EXISTS leave_requests(
 id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL, employee_id INTEGER, leave_type TEXT, from_date TEXT, to_date TEXT,
 days REAL, reason TEXT, status TEXT DEFAULT 'Pending', approved_by TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP,
 category TEXT DEFAULT 'Leave'
);
CREATE TABLE IF NOT EXISTS attendance(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,work_date TEXT,first_in TEXT,last_out TEXT,
 status TEXT DEFAULT 'Present',late_minutes INTEGER DEFAULT 0,overtime_minutes INTEGER DEFAULT 0,source TEXT DEFAULT 'Manual',
 UNIQUE(employee_id,work_date)
);
CREATE TABLE IF NOT EXISTS punches(
 id SERIAL PRIMARY KEY,company_id INTEGER,biometric_id TEXT,punch_time TEXT,punch_type TEXT,device_id INTEGER,raw_payload TEXT,
 UNIQUE(company_id,biometric_id,punch_time)
);
CREATE TABLE IF NOT EXISTS biometric_devices(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT,model TEXT,serial_no TEXT,branch TEXT,ip TEXT,port INTEGER DEFAULT 4370,
 protocol TEXT DEFAULT 'ZKTeco/eSSL (LAN)',status TEXT DEFAULT 'Not Tested',last_sync TEXT,last_error TEXT,api_key TEXT
);
CREATE TABLE IF NOT EXISTS payroll(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,month TEXT,gross REAL DEFAULT 0,deductions REAL DEFAULT 0,
 lop REAL DEFAULT 0,ot REAL DEFAULT 0,net REAL DEFAULT 0,status TEXT DEFAULT 'Draft',payslip_no TEXT,
 pf_employee REAL DEFAULT 0,esic_employee REAL DEFAULT 0,tds REAL DEFAULT 0,lop_days REAL DEFAULT 0,
 UNIQUE(employee_id,month)
);
CREATE TABLE IF NOT EXISTS candidates(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT,email TEXT,phone TEXT,position TEXT,status TEXT DEFAULT 'Applied',
 interview_date TEXT,notes TEXT
);
CREATE TABLE IF NOT EXISTS onboarding(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,offer INTEGER DEFAULT 0,documents INTEGER DEFAULT 0,
 verification INTEGER DEFAULT 0,assets INTEGER DEFAULT 0,policy INTEGER DEFAULT 0,completed INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS performance(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,cycle TEXT,goals TEXT,rating REAL,manager_comments TEXT,status TEXT DEFAULT 'Open',
 reviewer TEXT,finalized_at TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS increments(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER NOT NULL,cycle TEXT,rating REAL,percent REAL,
 old_ctc REAL,new_ctc REAL,effective_date TEXT,applied_by TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS awards(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER NOT NULL,award_type TEXT DEFAULT 'Employee of the Month',
 period TEXT,score REAL,note TEXT,declared_by TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP,
 UNIQUE(company_id,award_type,period)
);
CREATE TABLE IF NOT EXISTS assets(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,asset_code TEXT,name TEXT,category TEXT,serial_no TEXT,status TEXT DEFAULT 'Available',
 employee_id INTEGER,issued_date TEXT,return_date TEXT,UNIQUE(company_id,asset_code)
);
CREATE TABLE IF NOT EXISTS expenses(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,category TEXT,amount REAL,expense_date TEXT,description TEXT,status TEXT DEFAULT 'Pending'
);
CREATE TABLE IF NOT EXISTS documents(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,name TEXT,doc_type TEXT,expiry_date TEXT,status TEXT DEFAULT 'Active',
 file_name TEXT,file_mime TEXT,file_data BYTEA,uploaded_by INTEGER,created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS agreements(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER NOT NULL,title TEXT,content TEXT,
 status TEXT DEFAULT 'Pending Employee',
 employee_signature TEXT,employee_signed_name TEXT,employee_signed_at TEXT,
 hr_signature TEXT,hr_signed_name TEXT,hr_signed_at TEXT,hr_signed_by INTEGER,
 director_signature TEXT,director_signed_name TEXT,director_signed_at TEXT,director_signed_by INTEGER,
 created_at TEXT DEFAULT CURRENT_TIMESTAMP,completed_at TEXT
);
CREATE TABLE IF NOT EXISTS announcements(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,title TEXT,body TEXT,audience TEXT DEFAULT 'All',created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS tickets(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,subject TEXT,description TEXT,priority TEXT DEFAULT 'Medium',
 status TEXT DEFAULT 'Open',assigned_to TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS exit_requests(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,resignation_date TEXT,last_working_date TEXT,
 reason TEXT,status TEXT DEFAULT 'Pending',clearance TEXT DEFAULT 'Pending',fnf_status TEXT DEFAULT 'Pending'
);
CREATE TABLE IF NOT EXISTS audit_logs(
 id SERIAL PRIMARY KEY,company_id INTEGER,user_id INTEGER,action TEXT,module TEXT,details TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS password_resets(
 token TEXT PRIMARY KEY,user_id INTEGER NOT NULL,expires_at BIGINT NOT NULL,used INTEGER DEFAULT 0
);
`);
for(const col of ["serial_no TEXT","last_error TEXT","api_key TEXT"]){
  try{await db.exec(`ALTER TABLE biometric_devices ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["email TEXT"]){
  try{await db.exec(`ALTER TABLE users ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["smtp_user TEXT","smtp_pass TEXT","policy_agreement_text TEXT","increment_policy TEXT"]){
  try{await db.exec(`ALTER TABLE companies ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["reviewer TEXT","finalized_at TEXT","created_at TEXT DEFAULT CURRENT_TIMESTAMP"]){
  try{await db.exec(`ALTER TABLE performance ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["file_name TEXT","file_mime TEXT","file_data BYTEA","uploaded_by INTEGER","created_at TEXT DEFAULT CURRENT_TIMESTAMP"]){
  try{await db.exec(`ALTER TABLE documents ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["reporting_manager_id INTEGER","pf_number TEXT","esic_number TEXT","uan_number TEXT","pan_number TEXT","date_of_birth TEXT","basic_salary REAL DEFAULT 0","hra REAL DEFAULT 0","other_allowances REAL DEFAULT 0","pf_applicable INTEGER DEFAULT 0","esic_applicable INTEGER DEFAULT 0"]){
  try{await db.exec(`ALTER TABLE employees ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["category TEXT DEFAULT 'Leave'"]){
  try{await db.exec(`ALTER TABLE leave_requests ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["pf_employee REAL DEFAULT 0","esic_employee REAL DEFAULT 0","tds REAL DEFAULT 0","lop_days REAL DEFAULT 0"]){
  try{await db.exec(`ALTER TABLE payroll ADD COLUMN ${col}`)}catch(e){}
}
for(const [oldN,newN] of [["Casual Leave","CL - Casual Leave"],["Sick Leave","SL - Sick Leave"],["Earned Leave","PL - Privilege Leave"]]){
  try{await db.exec(`UPDATE leave_types SET name='${newN}' WHERE name='${oldN}'`)}catch(e){}
}
for(const [idx,sql] of Object.entries({
  ux_emp_pf:"CREATE UNIQUE INDEX IF NOT EXISTS ux_emp_pf ON employees(company_id,pf_number) WHERE pf_number IS NOT NULL AND pf_number<>''",
  ux_emp_esic:"CREATE UNIQUE INDEX IF NOT EXISTS ux_emp_esic ON employees(company_id,esic_number) WHERE esic_number IS NOT NULL AND esic_number<>''",
  ux_emp_uan:"CREATE UNIQUE INDEX IF NOT EXISTS ux_emp_uan ON employees(company_id,uan_number) WHERE uan_number IS NOT NULL AND uan_number<>''",
  ux_emp_pan:"CREATE UNIQUE INDEX IF NOT EXISTS ux_emp_pan ON employees(company_id,pan_number) WHERE pan_number IS NOT NULL AND pan_number<>''",
  ux_emp_bio:"CREATE UNIQUE INDEX IF NOT EXISTS ux_emp_bio ON employees(company_id,biometric_id) WHERE biometric_id IS NOT NULL AND biometric_id<>''"
})){
  try{await db.exec(sql)}catch(e){}
}
}

const hash=(p,s)=>crypto.scryptSync(p,s,64).toString("hex")+":"+s;
const verify=(p,stored)=>{
  try{
    const [h,s]=stored.split(":"); const c=crypto.scryptSync(p,s,64).toString("hex");
    return crypto.timingSafeEqual(Buffer.from(h,"hex"),Buffer.from(c,"hex"));
  }catch{return false}
};

async function seedCompanyDefaults(companyId){
  for(const n of ["Operations","HR","Finance","IT","Sales","Admin"]){
    await db.prepare("INSERT OR IGNORE INTO departments(company_id,name) VALUES(?,?)").run(companyId,n);
  }
  for(const [n,b] of [["CL - Casual Leave",12],["SL - Sick Leave",12],["PL - Privilege Leave",18],["Unpaid Leave",0]]){
    await db.prepare("INSERT OR IGNORE INTO leave_types(company_id,name,annual_balance) VALUES(?,?,?)").run(companyId,n,b);
  }
}

async function seed(){
  if(!await db.prepare("SELECT id FROM users WHERE username=?").get("admin")){
    const s=crypto.randomBytes(16).toString("hex");
    await db.prepare("INSERT INTO users(company_id,username,password_hash,role) VALUES(NULL,?,?,?)").run("admin",hash("Admin@12345",s),"Super Admin");
  }
  if(!await db.prepare("SELECT id FROM companies LIMIT 1").get()){
    const c=await db.prepare("INSERT INTO companies(name,code,industry,address,contact_email,contact_phone,status) VALUES(?,?,?,?,?,?,?)")
      .run("BMS Demo Company","BMSDEMO","Business Services","Head Office, India","hr@bmsdemo.local","9999999999","Active");
    const companyId=c.lastInsertRowid;
    await seedCompanyDefaults(companyId);
    const defs=[["hr","HR@12345","HR Admin"],["manager","Manager@12345","Manager"],["finance","Finance@12345","Finance"]];
    for(const [u,p,r] of defs){
      const s=crypto.randomBytes(16).toString("hex");
      await db.prepare("INSERT INTO users(company_id,username,password_hash,role) VALUES(?,?,?,?)").run(companyId,u,hash(p,s),r);
    }
    const e=await db.prepare(`INSERT INTO employees(company_id,employee_code,name,email,phone,department,designation,branch,joining_date,salary,biometric_id)
      VALUES(?,?,?,?,?,?,?,?,?,?,?)`).run(companyId,"BMS001","Demo Employee","employee@bmsdemo.local","9999999999","Operations","Executive","Head Office",new Date().toISOString().slice(0,10),30000,"1001");
    await db.prepare("INSERT INTO users(company_id,username,password_hash,role,employee_id) VALUES(?,?,?,?,?)")
      .run(companyId,"employee",hash("Employee@12345",crypto.randomBytes(16).toString("hex")),"Employee",e.lastInsertRowid);
    await db.prepare("INSERT INTO onboarding(company_id,employee_id) VALUES(?,?)").run(companyId,e.lastInsertRowid);
  }
}

function token(req){
  const c=req.headers.cookie||""; const m=c.match(/(?:^|;\s*)bms_session=([^;]+)/);
  return m?decodeURIComponent(m[1]):null;
}
async function me(req){
  const t=token(req); if(!t)return null;
  const row=await db.prepare(`SELECT u.id,u.username,u.role,u.employee_id,u.company_id,s.active_company_id FROM sessions s JOIN users u ON u.id=s.user_id
    WHERE s.token=? AND s.expires_at>? AND u.active=1`).get(t,Date.now());
  if(!row)return null;
  const effectiveCompany=row.role==="Super Admin"?(row.active_company_id||null):row.company_id;
  return {id:row.id,username:row.username,role:row.role,employee_id:row.employee_id,company_id:effectiveCompany};
}
async function auth(req,res,next){
  try{
    const u=await me(req); if(!u)return res.status(401).json({error:"Login required"});
    req.user=u; next();
  }catch(e){next(e)}
}
function roles(...allowed){return (req,res,next)=>allowed.includes(req.user.role)?next():res.status(403).json({error:"Permission denied"});}
function requireCompany(req,res,next){
  if(!req.user.company_id)return res.status(400).json({error:"Select a company first"});
  next();
}
async function audit(req,action,module,details=""){await db.prepare("INSERT INTO audit_logs(company_id,user_id,action,module,details) VALUES(?,?,?,?,?)").run(req.user.company_id||null,req.user.id,action,module,details)}
async function companySender(companyId){
  if(!companyId)return null;
  const c=await db.prepare("SELECT name,smtp_user,smtp_pass FROM companies WHERE id=?").get(companyId);
  return c?{name:c.name,smtp_user:c.smtp_user,smtp_pass:c.smtp_pass}:null;
}
function wrap(fn){return (req,res)=>fn(req,res).catch(e=>{console.error(e);res.status(500).json({error:e.message||"Server error"})})}

const LOGIN_FAILS=new Map();
const LOGIN_MAX=8,LOGIN_WINDOW=15*60*1000;
app.post("/api/login",wrap(async(req,res)=>{
  const key=(req.ip||"")+"|"+String(req.body.username||"").toLowerCase();
  const now=Date.now();
  const rec=LOGIN_FAILS.get(key);
  if(rec && now-rec.first>LOGIN_WINDOW)LOGIN_FAILS.delete(key);
  const cur=LOGIN_FAILS.get(key);
  if(cur && cur.count>=LOGIN_MAX)return res.status(429).json({error:"Too many failed sign-in attempts. Please try again in 15 minutes."});
  const u=await db.prepare("SELECT * FROM users WHERE username=? AND active=1").get(req.body.username||"");
  if(!u||!verify(req.body.password||"",u.password_hash)){
    const f=LOGIN_FAILS.get(key)||{count:0,first:now};f.count++;LOGIN_FAILS.set(key,f);
    return res.status(401).json({error:"Invalid username or password"});
  }
  LOGIN_FAILS.delete(key);
  const t=crypto.randomBytes(32).toString("hex");
  await db.prepare("INSERT INTO sessions(token,user_id,expires_at,active_company_id) VALUES(?,?,?,?)").run(t,u.id,Date.now()+8*60*60*1000,u.company_id||null);
  res.setHeader("Set-Cookie",`bms_session=${t}; HttpOnly; SameSite=Lax; Path=/; Max-Age=28800${process.env.NODE_ENV==="production"?"; Secure":""}`);
  await db.prepare("INSERT INTO audit_logs(company_id,user_id,action,module,details) VALUES(?,?,?,?,?)").run(u.company_id||null,u.id,"LOGIN","AUTH","");
  res.json({ok:true,user:{username:u.username,role:u.role,employee_id:u.employee_id}});
}));
app.post("/api/logout",auth,wrap(async(req,res)=>{
  const t=token(req);await db.prepare("DELETE FROM sessions WHERE token=?").run(t);
  res.setHeader("Set-Cookie","bms_session=; HttpOnly; SameSite=Lax; Path=/; Max-Age=0");
  res.json({ok:true});
}));
app.get("/api/me",auth,wrap(async(req,res)=>{
  let company=null;
  if(req.user.company_id)company=await db.prepare("SELECT id,name,code,industry,status FROM companies WHERE id=?").get(req.user.company_id);
  res.json({user:req.user,company});
}));

/* ---------------- Companies (Super Admin / platform) ---------------- */
app.get("/api/companies",auth,roles("Super Admin"),wrap(async(req,res)=>{
  res.json(await db.prepare(`SELECT c.*,(SELECT COUNT(*) FROM employees e WHERE e.company_id=c.id AND e.status='Active') employee_count
    FROM companies c ORDER BY c.id DESC`).all());
}));
app.post("/api/companies",auth,roles("Super Admin"),wrap(async(req,res)=>{
  const x=req.body;
  if(!x.name||!x.code)return res.status(400).json({error:"Company name and code are required"});
  if(!x.admin_username||!x.admin_password)return res.status(400).json({error:"First HR Admin username and password are required"});
  if(x.admin_password.length<8)return res.status(400).json({error:"Admin password must be at least 8 characters"});
  if(await db.prepare("SELECT id FROM users WHERE username=?").get(x.admin_username))return res.status(400).json({error:"Username already taken"});
  try{
    const c=await db.prepare("INSERT INTO companies(name,code,industry,address,contact_email,contact_phone,status,smtp_user,smtp_pass) VALUES(?,?,?,?,?,?,?,?,?)")
      .run(x.name,x.code.toUpperCase(),x.industry||"",x.address||"",x.contact_email||"",x.contact_phone||"","Active",x.smtp_user||null,x.smtp_pass||null);
    const companyId=c.lastInsertRowid;
    await seedCompanyDefaults(companyId);
    const s=crypto.randomBytes(16).toString("hex");
    await db.prepare("INSERT INTO users(company_id,username,password_hash,role,email) VALUES(?,?,?,?,?)").run(companyId,x.admin_username,hash(x.admin_password,s),"HR Admin",x.contact_email||null);
    await audit(req,"ONBOARD","COMPANY",x.name);
    if(x.contact_email){
      sendMail(x.contact_email,`Welcome to BMS HRMS — ${x.name}`,layout("Your company workspace is ready",
        `<p>Hi,</p><p>Your company <b>${x.name}</b> has been onboarded on BMS Enterprise HRMS. You've been set up as the HR Admin.</p>
         <p><b>Login URL:</b> ${req.protocol}://${req.get("host")}<br><b>Username:</b> ${x.admin_username}<br><b>Password:</b> (the one you set during onboarding)</p>
         <p>Please log in and change your password from the header menu.</p>`),
        {smtp_user:x.smtp_user,smtp_pass:x.smtp_pass,name:x.name}).catch(()=>{});
    }
    res.json({ok:true,id:companyId});
  }catch(e){res.status(400).json({error:/duplicate key|unique/i.test(e.message)?"Company code already exists":e.message})}
}));
app.put("/api/companies/:id",auth,roles("Super Admin"),wrap(async(req,res)=>{
  const x=req.body;
  await db.prepare("UPDATE companies SET name=?,industry=?,address=?,contact_email=?,contact_phone=? WHERE id=?")
    .run(x.name,x.industry||"",x.address||"",x.contact_email||"",x.contact_phone||"",req.params.id);
  await audit(req,"UPDATE","COMPANY",req.params.id);res.json({ok:true});
}));
app.get("/api/email-settings",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const c=await db.prepare("SELECT smtp_user FROM companies WHERE id=?").get(req.user.company_id);
  res.json({smtp_user:c?.smtp_user||"",configured:!!c?.smtp_user});
}));
app.post("/api/email-settings",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const {smtp_user,smtp_pass}=req.body||{};
  if(smtp_pass) await db.prepare("UPDATE companies SET smtp_user=?,smtp_pass=? WHERE id=?").run(smtp_user||null,smtp_pass,req.user.company_id);
  else await db.prepare("UPDATE companies SET smtp_user=? WHERE id=?").run(smtp_user||null,req.user.company_id);
  await audit(req,"UPDATE","EMAIL_SETTINGS","");res.json({ok:true});
}));
app.post("/api/companies/:id/status",auth,roles("Super Admin"),wrap(async(req,res)=>{
  await db.prepare("UPDATE companies SET status=? WHERE id=?").run(req.body.status,req.params.id);
  await audit(req,"STATUS","COMPANY",req.params.id+":"+req.body.status);res.json({ok:true});
}));
app.post("/api/switch-company",auth,roles("Super Admin"),wrap(async(req,res)=>{
  const t=token(req);
  const cid=req.body.company_id?Number(req.body.company_id):null;
  if(cid && !await db.prepare("SELECT id FROM companies WHERE id=?").get(cid))return res.status(404).json({error:"Company not found"});
  await db.prepare("UPDATE sessions SET active_company_id=? WHERE token=?").run(cid,t);
  res.json({ok:true});
}));

app.get("/api/calendar",auth,requireCompany,wrap(async(req,res)=>{
  const rows=await db.prepare("SELECT id,name,employee_code,date_of_birth,joining_date,department FROM employees WHERE company_id=? AND status='Active'").all(req.user.company_id);
  const today=new Date();
  const upcoming=(dateStr,type)=>{
    if(!dateStr)return null;
    const d=new Date(dateStr);
    if(isNaN(d))return null;
    let next=new Date(today.getFullYear(),d.getMonth(),d.getDate());
    if(next<new Date(today.getFullYear(),today.getMonth(),today.getDate()))next.setFullYear(next.getFullYear()+1);
    const daysAway=Math.round((next-new Date(today.getFullYear(),today.getMonth(),today.getDate()))/86400000);
    return {type,date:next.toISOString().slice(0,10),daysAway,years:type==="anniversary"?(next.getFullYear()-d.getFullYear()):null};
  };
  const events=[];
  for(const e of rows){
    const b=upcoming(e.date_of_birth,"birthday");
    if(b)events.push({employee_id:e.id,name:e.name,employee_code:e.employee_code,department:e.department,...b});
    const a=upcoming(e.joining_date,"anniversary");
    if(a)events.push({employee_id:e.id,name:e.name,employee_code:e.employee_code,department:e.department,...a});
  }
  events.sort((x,y)=>x.daysAway-y.daysAway);
  res.json(events);
}));

app.get("/api/dashboard",auth,wrap(async(req,res)=>{
  const today=new Date().toISOString().slice(0,10);
  if(req.user.role==="Super Admin" && !req.user.company_id){
    res.json({
      platform:true,
      companies:(await db.prepare("SELECT COUNT(*) c FROM companies").get()).c,
      activeCompanies:(await db.prepare("SELECT COUNT(*) c FROM companies WHERE status='Active'").get()).c,
      employees:(await db.prepare("SELECT COUNT(*) c FROM employees WHERE status='Active'").get()).c,
      tickets:(await db.prepare("SELECT COUNT(*) c FROM tickets WHERE status!='Closed'").get()).c
    });
    return;
  }
  const cid=req.user.company_id;
  res.json({
    platform:false,
    employees:(await db.prepare("SELECT COUNT(*) c FROM employees WHERE status='Active' AND company_id=?").get(cid)).c,
    present:(await db.prepare("SELECT COUNT(*) c FROM attendance WHERE work_date=? AND status='Present' AND company_id=?").get(today,cid)).c,
    onLeave:(await db.prepare("SELECT COUNT(*) c FROM leave_requests WHERE status='Approved' AND from_date<=? AND to_date>=? AND company_id=?").get(today,today,cid)).c,
    payroll:(await db.prepare("SELECT COALESCE(SUM(net),0) s FROM payroll WHERE month=? AND company_id=?").get(today.slice(0,7),cid)).s,
    candidates:(await db.prepare("SELECT COUNT(*) c FROM candidates WHERE status NOT IN ('Rejected','Joined') AND company_id=?").get(cid)).c,
    devices:(await db.prepare("SELECT COUNT(*) c FROM biometric_devices WHERE company_id=?").get(cid)).c,
    tickets:(await db.prepare("SELECT COUNT(*) c FROM tickets WHERE status!='Closed' AND company_id=?").get(cid)).c,
    attendanceTrend:await db.prepare(`SELECT work_date,COUNT(*) c FROM attendance WHERE company_id=? AND status='Present' AND work_date::date>=(?::date - interval '6 days') GROUP BY work_date ORDER BY work_date`).all(cid,today)
  });
}));

app.get("/api/employees",auth,requireCompany,wrap(async(req,res)=>{
  let rows=await db.prepare("SELECT * FROM employees WHERE company_id=? ORDER BY id DESC").all(req.user.company_id);
  if(req.user.role==="Employee") rows=rows.filter(x=>x.id===req.user.employee_id);
  else if(req.user.role==="Manager" && req.user.employee_id) rows=rows.filter(x=>x.reporting_manager_id===req.user.employee_id || x.id===req.user.employee_id);
  res.json(rows);
}));
function friendlyDupError(e){
  const m=e.message||"";
  const map={ux_emp_pf:"PF Number",ux_emp_esic:"ESIC Number",ux_emp_uan:"UAN Number",ux_emp_pan:"PAN Number",ux_emp_bio:"Biometric ID"};
  for(const [idx,label] of Object.entries(map)) if(m.includes(idx))return `This ${label} is already used by another employee in this company`;
  if(m.includes("employee_code")||/duplicate key|unique/i.test(m))return "Employee code already exists";
  return m;
}
app.post("/api/employees",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await db.prepare(`INSERT INTO employees(company_id,employee_code,name,email,phone,department,designation,manager,reporting_manager_id,branch,joining_date,status,biometric_id,salary,bank_name,bank_account,ifsc,pf_number,esic_number,uan_number,pan_number,date_of_birth,basic_salary,hra,other_allowances,pf_applicable,esic_applicable)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(req.user.company_id,x.employee_code,x.name,x.email,x.phone,x.department,x.designation,x.manager,x.reporting_manager_id||null,x.branch,x.joining_date,x.status||"Active",x.biometric_id||null,x.salary||0,x.bank_name,x.bank_account,x.ifsc,x.pf_number||null,x.esic_number||null,x.uan_number||null,x.pan_number||null,x.date_of_birth||null,x.basic_salary||0,x.hra||0,x.other_allowances||0,+!!x.pf_applicable,+!!x.esic_applicable);
    await db.prepare("INSERT OR IGNORE INTO onboarding(company_id,employee_id) VALUES(?,?)").run(req.user.company_id,r.lastInsertRowid);
    await audit(req,"CREATE","EMPLOYEE",x.employee_code);res.json({id:r.lastInsertRowid});
  }catch(e){res.status(400).json({error:friendlyDupError(e)})}
}));
app.put("/api/employees/:id",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await db.prepare(`UPDATE employees SET employee_code=?,name=?,email=?,phone=?,department=?,designation=?,manager=?,reporting_manager_id=?,branch=?,joining_date=?,status=?,biometric_id=?,salary=?,bank_name=?,bank_account=?,ifsc=?,pf_number=?,esic_number=?,uan_number=?,pan_number=?,date_of_birth=?,basic_salary=?,hra=?,other_allowances=?,pf_applicable=?,esic_applicable=? WHERE id=? AND company_id=?`)
      .run(x.employee_code,x.name,x.email,x.phone,x.department,x.designation,x.manager,x.reporting_manager_id||null,x.branch,x.joining_date,x.status||"Active",x.biometric_id||null,x.salary||0,x.bank_name,x.bank_account,x.ifsc,x.pf_number||null,x.esic_number||null,x.uan_number||null,x.pan_number||null,x.date_of_birth||null,x.basic_salary||0,x.hra||0,x.other_allowances||0,+!!x.pf_applicable,+!!x.esic_applicable,req.params.id,req.user.company_id);
    if(r.changes===0)return res.status(404).json({error:"Employee not found"});
    await audit(req,"UPDATE","EMPLOYEE",x.employee_code);res.json({ok:true});
  }catch(e){res.status(400).json({error:friendlyDupError(e)})}
}));
app.post("/api/employees/:id/status",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE employees SET status=? WHERE id=? AND company_id=?").run(req.body.status,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Employee not found"});
  await audit(req,"STATUS","EMPLOYEE",req.params.id+":"+req.body.status);res.json({ok:true});
}));
app.post("/api/employees/:id/create-login",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found"});
  if(await db.prepare("SELECT id FROM users WHERE employee_id=? AND company_id=?").get(emp.id,req.user.company_id))
    return res.status(400).json({error:"This employee already has a login"});
  const username=(req.body.username||emp.employee_code||"").toLowerCase().trim();
  if(!username)return res.status(400).json({error:"Username is required"});
  if(await db.prepare("SELECT id FROM users WHERE username=?").get(username))return res.status(400).json({error:"Username already taken"});
  const tempPassword=crypto.randomBytes(6).toString("base64").replace(/[^a-zA-Z0-9]/g,"").slice(0,10)+"@1";
  const s=crypto.randomBytes(16).toString("hex");
  await db.prepare("INSERT INTO users(company_id,username,password_hash,role,employee_id,email) VALUES(?,?,?,?,?,?)")
    .run(req.user.company_id,username,hash(tempPassword,s),"Employee",emp.id,emp.email||null);
  await audit(req,"CREATE_LOGIN","EMPLOYEE",username);
  if(emp.email){
    sendMail(emp.email,"Your BMS HRMS login",layout("Welcome aboard!",
      `<p>Hi ${emp.name},</p><p>Your employee self-service login has been created.</p>
       <p><b>Login URL:</b> ${req.protocol}://${req.get("host")}<br><b>Username:</b> ${username}<br><b>Temporary Password:</b> ${tempPassword}</p>
       <p>Please log in and change your password from the header menu.</p>`),
      await companySender(req.user.company_id)).catch(()=>{});
  }
  res.json({ok:true,username,temp_password:tempPassword,emailed:!!emp.email});
}));
app.post("/api/change-password",auth,wrap(async(req,res)=>{
  const u=await db.prepare("SELECT * FROM users WHERE id=?").get(req.user.id);
  if(!verify(req.body.current_password||"",u.password_hash))return res.status(400).json({error:"Current password incorrect"});
  if(!req.body.new_password||req.body.new_password.length<8)return res.status(400).json({error:"New password must be at least 8 characters"});
  const s=crypto.randomBytes(16).toString("hex");
  await db.prepare("UPDATE users SET password_hash=? WHERE id=?").run(hash(req.body.new_password,s),req.user.id);
  await audit(req,"CHANGE_PASSWORD","AUTH","");res.json({ok:true});
}));

app.post("/api/forgot-password",wrap(async(req,res)=>{
  const u=await db.prepare("SELECT * FROM users WHERE username=? AND active=1").get(req.body.username||"");
  // Always return ok (don't reveal whether a username exists), but only actually email if we have an address.
  if(u){
    let email=u.email;
    if(!email && u.employee_id){
      const emp=await db.prepare("SELECT email FROM employees WHERE id=?").get(u.employee_id);
      email=emp?.email;
    }
    if(email){
      const t=crypto.randomBytes(32).toString("hex");
      await db.prepare("INSERT INTO password_resets(token,user_id,expires_at) VALUES(?,?,?)").run(t,u.id,Date.now()+30*60*1000);
      const link=`${req.protocol}://${req.get("host")}/?reset=${t}`;
      sendMail(email,"Reset your BMS HRMS password",layout("Password reset requested",
        `<p>Hi ${u.username},</p><p>Click the link below to reset your password. This link expires in 30 minutes.</p>
         <p><a href="${link}" style="background:#4f46e5;color:#fff;padding:10px 16px;border-radius:8px;text-decoration:none;display:inline-block">Reset Password</a></p>
         <p style="font-size:12px;color:#64748b">If you didn't request this, you can safely ignore this email.</p>`)).catch(()=>{});
    }
  }
  res.json({ok:true});
}));
app.post("/api/reset-password",wrap(async(req,res)=>{
  const {token,new_password}=req.body||{};
  if(!token||!new_password||new_password.length<8)return res.status(400).json({error:"Invalid request. Password must be at least 8 characters."});
  const r=await db.prepare("SELECT * FROM password_resets WHERE token=?").get(token);
  if(!r||r.used||Number(r.expires_at)<Date.now())return res.status(400).json({error:"This reset link is invalid or has expired."});
  const s=crypto.randomBytes(16).toString("hex");
  await db.prepare("UPDATE users SET password_hash=? WHERE id=?").run(hash(new_password,s),r.user_id);
  await db.prepare("UPDATE password_resets SET used=1 WHERE token=?").run(token);
  await db.prepare("DELETE FROM sessions WHERE user_id=?").run(r.user_id);
  res.json({ok:true});
}));

/* ---------------- Team (internal, non-employee logins: Director, extra HR/Manager/Finance) ---------------- */
const TEAM_ROLES=["Director","HR Admin","Manager","Finance"];
app.get("/api/team",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  res.json(await db.prepare(`SELECT u.id,u.username,u.role,u.email,u.active,e.name employee_name,e.employee_code FROM users u LEFT JOIN employees e ON e.id=u.employee_id WHERE u.company_id=? AND u.role<>'Employee' ORDER BY u.id DESC`).all(req.user.company_id));
}));
app.post("/api/team",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  if(!TEAM_ROLES.includes(x.role))return res.status(400).json({error:"Invalid role"});
  if(!x.username||!x.password||x.password.length<8)return res.status(400).json({error:"Username and a password (min 8 chars) are required"});
  if(await db.prepare("SELECT id FROM users WHERE username=?").get(x.username))return res.status(400).json({error:"Username already taken"});
  let linkedEmployeeId=null;
  if(x.employee_id){
    const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(x.employee_id,req.user.company_id);
    if(!emp)return res.status(400).json({error:"Selected employee not found in this company"});
    linkedEmployeeId=emp.id;
  }
  const s=crypto.randomBytes(16).toString("hex");
  await db.prepare("INSERT INTO users(company_id,username,password_hash,role,email,employee_id) VALUES(?,?,?,?,?,?)").run(req.user.company_id,x.username,hash(x.password,s),x.role,x.email||null,linkedEmployeeId);
  await audit(req,"CREATE","TEAM",`${x.username} (${x.role})`);
  if(x.email){
    sendMail(x.email,"Your BMS HRMS login",layout("Welcome to the team",
      `<p>Hi,</p><p>You've been added as <b>${x.role}</b> on BMS Enterprise HRMS.</p>
       <p><b>Login URL:</b> ${req.protocol}://${req.get("host")}<br><b>Username:</b> ${x.username}<br><b>Password:</b> (the one shared with you)</p>`),
      await companySender(req.user.company_id)).catch(()=>{});
  }
  res.json({ok:true});
}));
app.post("/api/team/:id/status",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE users SET active=? WHERE id=? AND company_id=? AND role<>'Employee'").run(req.body.active?1:0,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"User not found"});
  res.json({ok:true});
}));

/* ---------------- Onboarding Agreements (Employee -> HR -> Director e-signature) ---------------- */
app.get("/api/policy-agreement",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const c=await db.prepare("SELECT policy_agreement_text FROM companies WHERE id=?").get(req.user.company_id);
  res.json({text:c?.policy_agreement_text||""});
}));
app.post("/api/policy-agreement",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  await db.prepare("UPDATE companies SET policy_agreement_text=? WHERE id=?").run(req.body.text||"",req.user.company_id);
  res.json({ok:true});
}));

app.get("/api/agreements",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT a.*,e.name employee_name,e.employee_code FROM agreements a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND a.employee_id=?";p.push(req.user.employee_id)}
  q+=" ORDER BY a.id DESC";
  res.json(await db.prepare(q).all(...p));
}));
app.post("/api/employees/:id/agreements",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found"});
  const company=await db.prepare("SELECT policy_agreement_text FROM companies WHERE id=?").get(req.user.company_id);
  if(!company?.policy_agreement_text?.trim())return res.status(400).json({error:"Set up the company's Policy Agreement text first (HR Policies page)"});
  const existing=await db.prepare("SELECT id FROM agreements WHERE employee_id=? AND company_id=? AND status<>'Completed'").get(emp.id,req.user.company_id);
  if(existing)return res.status(400).json({error:"This employee already has an agreement in progress"});
  const r=await db.prepare("INSERT INTO agreements(company_id,employee_id,title,content,status) VALUES(?,?,?,?,?)")
    .run(req.user.company_id,emp.id,`Company Policy Agreement — ${emp.name}`,company.policy_agreement_text,"Pending Employee");
  await audit(req,"CREATE","AGREEMENT",emp.employee_code);
  res.json({id:r.lastInsertRowid});
}));
app.post("/api/agreements/:id/sign",auth,requireCompany,wrap(async(req,res)=>{
  const ag=await db.prepare("SELECT * FROM agreements WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!ag)return res.status(404).json({error:"Agreement not found"});
  const {signature,signed_name}=req.body||{};
  if(!signature||!signed_name)return res.status(400).json({error:"Signature and name are required"});
  const now=new Date().toISOString();
  if(ag.status==="Pending Employee"){
    if(req.user.role!=="Employee"||req.user.employee_id!==ag.employee_id)return res.status(403).json({error:"Only the employee can sign this step"});
    await db.prepare("UPDATE agreements SET employee_signature=?,employee_signed_name=?,employee_signed_at=?,status='Pending HR' WHERE id=?").run(signature,signed_name,now,ag.id);
  }else if(ag.status==="Pending HR"){
    if(!["Super Admin","HR Admin"].includes(req.user.role))return res.status(403).json({error:"Only HR can sign this step"});
    await db.prepare("UPDATE agreements SET hr_signature=?,hr_signed_name=?,hr_signed_at=?,hr_signed_by=?,status='Pending Director' WHERE id=?").run(signature,signed_name,now,req.user.id,ag.id);
  }else if(ag.status==="Pending Director"){
    if(!["Super Admin","Director"].includes(req.user.role))return res.status(403).json({error:"Only the Director can sign this step"});
    await db.prepare("UPDATE agreements SET director_signature=?,director_signed_name=?,director_signed_at=?,director_signed_by=?,status='Completed',completed_at=? WHERE id=?").run(signature,signed_name,now,req.user.id,now,ag.id);
  }else{
    return res.status(400).json({error:"This agreement is already completed"});
  }
  await audit(req,"SIGN","AGREEMENT",String(ag.id));
  const updated=await db.prepare("SELECT * FROM agreements WHERE id=?").get(ag.id);
  if(updated.status==="Completed"){
    const emp=await db.prepare("SELECT * FROM employees WHERE id=?").get(updated.employee_id);
    const company=await db.prepare("SELECT name,contact_email,smtp_user,smtp_pass FROM companies WHERE id=?").get(req.user.company_id);
    const sigBlock=(label,name,sig,at)=>`<div style="margin:14px 0"><b>${label}:</b> ${esc(name)} <span style="color:#64748b;font-size:12px">(${at?new Date(at).toLocaleString("en-IN"):""})</span><br>${sig?`<img src="${sig}" style="height:70px;border-bottom:1px solid #94a3b8;margin-top:4px">`:""}</div>`;
    const html=layout(updated.title,
      `<div style="white-space:pre-wrap;border:1px solid #e5e7eb;padding:14px;border-radius:8px;background:#f8fafc">${esc(updated.content)}</div>
       ${sigBlock("Employee",updated.employee_signed_name,updated.employee_signature,updated.employee_signed_at)}
       ${sigBlock("HR",updated.hr_signed_name,updated.hr_signature,updated.hr_signed_at)}
       ${sigBlock("Director",updated.director_signed_name,updated.director_signature,updated.director_signed_at)}
       <p style="color:#166534;font-weight:700">Fully executed on ${new Date(updated.completed_at).toLocaleString("en-IN")}</p>`);
    const sender={smtp_user:company?.smtp_user,smtp_pass:company?.smtp_pass,name:company?.name};
    if(emp?.email)sendMail(emp.email,`Signed: ${updated.title}`,html,sender).catch(()=>{});
    if(company?.contact_email)sendMail(company.contact_email,`Signed: ${updated.title}`,html,sender).catch(()=>{});
  }
  res.json({ok:true,status:updated.status});
}));

function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]))}

app.get("/api/departments",auth,requireCompany,wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM departments WHERE company_id=? ORDER BY name").all(req.user.company_id))));
app.post("/api/departments",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  try{await db.prepare("INSERT INTO departments(company_id,name) VALUES(?,?)").run(req.user.company_id,req.body.name);res.json({ok:true})}
  catch(e){res.status(400).json({error:"Department already exists"})}
}));

app.get("/api/attendance",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT a.*,e.employee_code,e.name,e.department FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=?`;
  let params=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND e.id=?";params.push(req.user.employee_id)}
  else if(req.user.role==="Manager" && req.user.employee_id){q+=" AND (e.reporting_manager_id=? OR e.id=?)";params.push(req.user.employee_id,req.user.employee_id)}
  q+=" ORDER BY a.work_date DESC,a.first_in DESC";
  res.json(await db.prepare(q).all(...params));
}));
app.post("/api/attendance/manual",auth,requireCompany,wrap(async(req,res)=>{
  const employeeId=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  if(req.user.role==="Employee" && req.body.employee_id && Number(req.body.employee_id)!==req.user.employee_id)return res.status(403).json({error:"Permission denied"});
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(employeeId,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const d=req.body.work_date||new Date().toISOString().slice(0,10);
  await db.prepare(`INSERT INTO attendance(company_id,employee_id,work_date,first_in,last_out,status,late_minutes,overtime_minutes,source)
    VALUES(?,?,?,?,?,?,?,?,?) ON CONFLICT(employee_id,work_date) DO UPDATE SET first_in=COALESCE(excluded.first_in,attendance.first_in),last_out=COALESCE(excluded.last_out,attendance.last_out),status=excluded.status`)
    .run(req.user.company_id,employeeId,d,req.body.first_in||null,req.body.last_out||null,req.body.status||"Present",req.body.late_minutes||0,req.body.overtime_minutes||0,"Manual");
  await audit(req,"UPSERT","ATTENDANCE",String(employeeId));res.json({ok:true});
}));

app.get("/api/leaves",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT l.*,e.employee_code,e.name FROM leave_requests l JOIN employees e ON e.id=l.employee_id WHERE l.company_id=?`;
  let p=[req.user.company_id];if(req.user.role==="Employee"){q+=" AND l.employee_id=?";p.push(req.user.employee_id)}
  else if(req.user.role==="Manager" && req.user.employee_id){q+=" AND (e.reporting_manager_id=? OR e.id=?)";p.push(req.user.employee_id,req.user.employee_id)}
  if(req.query.category){q+=" AND l.category=?";p.push(req.query.category)}
  q+=" ORDER BY l.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/leaves",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const category=["WFH","Permission"].includes(req.body.category)?req.body.category:"Leave";
  const r=await db.prepare(`INSERT INTO leave_requests(company_id,employee_id,leave_type,from_date,to_date,days,reason,category) VALUES(?,?,?,?,?,?,?,?)`)
    .run(req.user.company_id,eid,req.body.leave_type,req.body.from_date,req.body.to_date,Number(req.body.days)||1,req.body.reason||"",category);
  await audit(req,"CREATE","LEAVE",String(r.lastInsertRowid));res.json({id:r.lastInsertRowid});
}));
app.post("/api/leaves/:id/status",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE leave_requests SET status=?,approved_by=? WHERE id=? AND company_id=?").run(req.body.status,req.user.username,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Leave request not found"});
  await audit(req,req.body.status,"LEAVE",req.params.id);res.json({ok:true});
}));
app.get("/api/leaves/balance",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.query.employee_id||req.user.employee_id);
  if(!eid)return res.json([]);
  const year=new Date().toISOString().slice(0,4);
  const types=await db.prepare("SELECT * FROM leave_types WHERE company_id=?").all(req.user.company_id);
  const used=await db.prepare(`SELECT leave_type,COALESCE(SUM(days),0) d FROM leave_requests WHERE employee_id=? AND company_id=? AND status='Approved' AND category='Leave' AND from_date LIKE ? GROUP BY leave_type`).all(eid,req.user.company_id,year+"%");
  const usedMap=Object.fromEntries(used.map(u=>[u.leave_type,u.d]));
  res.json(types.map(t=>({leave_type:t.name,annual_balance:t.annual_balance,used:usedMap[t.name]||0,remaining:t.annual_balance-(usedMap[t.name]||0)})));
}));

async function computePayroll(companyId,emp,month){
  const [y,m]=month.split("-").map(Number);
  const daysInMonth=new Date(y,m,0).getDate();
  const atts=await db.prepare("SELECT status FROM attendance WHERE employee_id=? AND work_date LIKE ?").all(emp.id,month+"%");
  let absentDays=atts.filter(a=>a.status==="Absent").length + atts.filter(a=>a.status==="Half Day").length*0.5;
  const unpaidLeaves=await db.prepare("SELECT COALESCE(SUM(days),0) d FROM leave_requests WHERE employee_id=? AND status='Approved' AND category='Leave' AND leave_type ILIKE '%Unpaid%' AND from_date LIKE ?").get(emp.id,month+"%");
  const lopDays=Math.min(daysInMonth,absentDays+Number(unpaidLeaves.d||0));
  const basic=Number(emp.basic_salary)||0,hra=Number(emp.hra)||0,other=Number(emp.other_allowances)||0;
  const gross=basic+hra+other;
  const perDay=gross/daysInMonth;
  const lop=+(perDay*lopDays).toFixed(2);
  const payableRatio=(daysInMonth-lopDays)/daysInMonth;
  const pfEmployee=emp.pf_applicable?+(basic*payableRatio*0.12).toFixed(2):0;
  const esicEmployee=(emp.esic_applicable && gross<=21000)?+(gross*payableRatio*0.0075).toFixed(2):0;
  return {daysInMonth,lopDays,gross:+gross.toFixed(2),lop,pfEmployee,esicEmployee,basic,hra,other};
}
async function processOnePayroll(req,employeeId,month,overrides={}){
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(employeeId,req.user.company_id);
  if(!emp)throw Object.assign(new Error("Employee not found in this company"),{status:404});
  const calc=await computePayroll(req.user.company_id,emp,month);
  const gross=overrides.gross!=null?Number(overrides.gross):calc.gross;
  const deductions=Number(overrides.deductions)||0;
  const lop=overrides.lop!=null?Number(overrides.lop):calc.lop;
  const ot=Number(overrides.ot)||0;
  const pfEmployee=overrides.pf_employee!=null?Number(overrides.pf_employee):calc.pfEmployee;
  const esicEmployee=overrides.esic_employee!=null?Number(overrides.esic_employee):calc.esicEmployee;
  const tds=Number(overrides.tds)||0;
  const net=+(gross-deductions-lop-pfEmployee-esicEmployee-tds+ot).toFixed(2);
  const no="BMS-"+Date.now()+"-"+employeeId;
  await db.prepare(`INSERT INTO payroll(company_id,employee_id,month,gross,deductions,lop,ot,net,status,payslip_no,pf_employee,esic_employee,tds,lop_days) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(employee_id,month) DO UPDATE SET gross=excluded.gross,deductions=excluded.deductions,lop=excluded.lop,ot=excluded.ot,net=excluded.net,status=excluded.status,payslip_no=excluded.payslip_no,pf_employee=excluded.pf_employee,esic_employee=excluded.esic_employee,tds=excluded.tds,lop_days=excluded.lop_days`)
    .run(req.user.company_id,employeeId,month,gross,deductions,lop,ot,net,overrides.status||"Processed",no,pfEmployee,esicEmployee,tds,calc.lopDays);
  await audit(req,"UPSERT","PAYROLL",month+":"+emp.employee_code);
  if(emp.email){
    const company=await db.prepare("SELECT name,smtp_user,smtp_pass FROM companies WHERE id=?").get(req.user.company_id);
    const inr=n=>"₹"+Number(n).toLocaleString("en-IN");
    sendMail(emp.email,`Payslip for ${month} — ${company?.name||"BMS HRMS"}`,layout(`Payslip — ${month}`,
      `<p>Hi ${emp.name},</p><p>Your payslip for <b>${month}</b> has been processed.</p>
       <table style="width:100%;border-collapse:collapse;margin-top:10px">
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Payslip No</td><td style="padding:6px;border:1px solid #e5e7eb">${no}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Gross</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(gross)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">LOP (${calc.lopDays} day(s))</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(lop)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">PF (Employee)</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(pfEmployee)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">ESIC (Employee)</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(esicEmployee)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">TDS</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(tds)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Other Deductions</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(deductions)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Overtime</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(ot)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb"><b>Net Pay</b></td><td style="padding:6px;border:1px solid #e5e7eb"><b>${inr(net)}</b></td></tr>
       </table>
       <p style="font-size:12px;color:#64748b">Log in to the HRMS to view or print your full payslip.</p>`),
      {smtp_user:company?.smtp_user,smtp_pass:company?.smtp_pass,name:company?.name}).catch(()=>{});
  }
  return {net,payslip_no:no,gross,lop,pfEmployee,esicEmployee};
}

app.get("/api/payroll",auth,requireCompany,roles("Super Admin","HR Admin","Finance","Manager"),wrap(async(req,res)=>{
  res.json(await db.prepare(`SELECT p.*,e.employee_code,e.name FROM payroll p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=? ORDER BY p.id DESC`).all(req.user.company_id));
}));
app.get("/api/payroll/calculate",auth,requireCompany,roles("Super Admin","HR Admin","Finance"),wrap(async(req,res)=>{
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(req.query.employee_id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  if(!req.query.month)return res.status(400).json({error:"month is required (YYYY-MM)"});
  res.json(await computePayroll(req.user.company_id,emp,req.query.month));
}));
app.post("/api/payroll/run-month",auth,requireCompany,roles("Super Admin","HR Admin","Finance"),wrap(async(req,res)=>{
  const month=req.body.month;
  if(!month)return res.status(400).json({error:"month is required (YYYY-MM)"});
  const employees=await db.prepare("SELECT id FROM employees WHERE company_id=? AND status='Active'").all(req.user.company_id);
  let processed=0,skipped=0;
  for(const e of employees){
    const already=await db.prepare("SELECT id FROM payroll WHERE employee_id=? AND month=?").get(e.id,month);
    if(already && !req.body.overwrite){skipped++;continue}
    try{await processOnePayroll(req,e.id,month,{});processed++}catch(err){skipped++}
  }
  res.json({ok:true,processed,skipped,total:employees.length});
}));
app.post("/api/payroll",auth,requireCompany,roles("Super Admin","HR Admin","Finance"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await processOnePayroll(req,x.employee_id,x.month,x);
    res.json({ok:true,net:r.net,payslip_no:r.payslip_no});
  }catch(e){res.status(e.status||400).json({error:e.message})}
}));

app.get("/api/candidates",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM candidates WHERE company_id=? ORDER BY id DESC").all(req.user.company_id))));
app.post("/api/candidates",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;const r=await db.prepare("INSERT INTO candidates(company_id,name,email,phone,position,status,interview_date,notes) VALUES(?,?,?,?,?,?,?,?)").run(req.user.company_id,x.name,x.email,x.phone,x.position,x.status||"Applied",x.interview_date,x.notes);
  await audit(req,"CREATE","RECRUITMENT",x.name);res.json({id:r.lastInsertRowid});
}));
app.post("/api/candidates/:id/status",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE candidates SET status=? WHERE id=? AND company_id=?").run(req.body.status,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Candidate not found"});
  await audit(req,"STATUS","RECRUITMENT",req.params.id+":"+req.body.status);res.json({ok:true});
}));

app.get("/api/onboarding",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT o.*,e.employee_code,e.name FROM onboarding o JOIN employees e ON e.id=o.employee_id WHERE o.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND e.id=?";p.push(req.user.employee_id)}q+=" ORDER BY o.id DESC";
  res.json(await db.prepare(q).all(...p));
}));
app.post("/api/onboarding/:id",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;const r=await db.prepare("UPDATE onboarding SET offer=?,documents=?,verification=?,assets=?,policy=?,completed=? WHERE id=? AND company_id=?")
    .run(+!!x.offer,+!!x.documents,+!!x.verification,+!!x.assets,+!!x.policy,+!!x.completed,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Record not found"});res.json({ok:true});
}));

async function isInTeam(req,employeeId){
  if(req.user.role==="Manager" && req.user.employee_id){
    const e=await db.prepare("SELECT reporting_manager_id FROM employees WHERE id=? AND company_id=?").get(employeeId,req.user.company_id);
    return !!e && e.reporting_manager_id===req.user.employee_id;
  }
  return true;
}
app.get("/api/performance",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT p.*,e.employee_code,e.name FROM performance p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND e.id=? AND p.status='Finalized'";p.push(req.user.employee_id)}
  else if(req.user.role==="Manager" && req.user.employee_id){q+=" AND e.reporting_manager_id=?";p.push(req.user.employee_id)}
  q+=" ORDER BY p.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/performance",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>{
  const x=req.body;
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(x.employee_id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  if(!await isInTeam(req,emp.id))return res.status(403).json({error:"You can only review employees who report to you"});
  const rating=Number(x.rating);
  if(!x.cycle||!String(x.cycle).trim())return res.status(400).json({error:"Review cycle is required (for example FY 2026-27)"});
  if(!(rating>=1&&rating<=5))return res.status(400).json({error:"Rating must be between 1 and 5"});
  const cycle=String(x.cycle).trim();
  const existing=await db.prepare("SELECT id,status FROM performance WHERE employee_id=? AND cycle=? AND company_id=?").get(emp.id,cycle,req.user.company_id);
  if(existing){
    if(existing.status==="Finalized")return res.status(400).json({error:"This review is already finalized and cannot be changed"});
    await db.prepare("UPDATE performance SET goals=?,rating=?,manager_comments=?,reviewer=? WHERE id=?").run(x.goals||"",rating,x.manager_comments||"",req.user.username,existing.id);
    await audit(req,"UPDATE","PERFORMANCE",String(existing.id));
    return res.json({id:existing.id,updated:true});
  }
  const r=await db.prepare("INSERT INTO performance(company_id,employee_id,cycle,goals,rating,manager_comments,status,reviewer) VALUES(?,?,?,?,?,?,?,?)").run(req.user.company_id,emp.id,cycle,x.goals||"",rating,x.manager_comments||"","Open",req.user.username);
  await audit(req,"CREATE","PERFORMANCE",String(r.lastInsertRowid));
  res.json({id:r.lastInsertRowid});
}));
app.post("/api/performance/:id/finalize",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE performance SET status='Finalized',finalized_at=? WHERE id=? AND company_id=? AND status<>'Finalized'").run(new Date().toISOString(),req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Review not found or already finalized"});
  await audit(req,"FINALIZE","PERFORMANCE",req.params.id);res.json({ok:true});
}));

/* ---------------- Yearly increment engine (rating -> % -> new CTC) ---------------- */
const DEFAULT_SLABS=[{min:4.5,pct:15},{min:4,pct:12},{min:3.5,pct:9},{min:3,pct:6},{min:2,pct:3},{min:0,pct:0}];
async function getSlabs(companyId){
  const c=await db.prepare("SELECT increment_policy FROM companies WHERE id=?").get(companyId);
  try{const s=JSON.parse(c?.increment_policy||"");if(Array.isArray(s)&&s.length)return s}catch{}
  return DEFAULT_SLABS;
}
function pctFor(slabs,rating){
  for(const s of [...slabs].sort((a,b)=>b.min-a.min)) if(rating>=s.min) return s.pct;
  return 0;
}
app.get("/api/increment-policy",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  res.json({slabs:await getSlabs(req.user.company_id)});
}));
app.post("/api/increment-policy",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const slabs=req.body?.slabs;
  if(!Array.isArray(slabs)||!slabs.length||slabs.length>10)return res.status(400).json({error:"Provide between 1 and 10 rating slabs"});
  const clean=[];
  for(const s of slabs){
    const min=Number(s.min),pct=Number(s.pct);
    if(!(min>=0&&min<=5)||!(pct>=0&&pct<=100))return res.status(400).json({error:"Each slab needs a minimum rating between 0 and 5 and a percentage between 0 and 100"});
    clean.push({min,pct});
  }
  if(!clean.some(s=>s.min===0))clean.push({min:0,pct:0});
  await db.prepare("UPDATE companies SET increment_policy=? WHERE id=?").run(JSON.stringify(clean),req.user.company_id);
  await audit(req,"UPDATE","INCREMENT_POLICY","");res.json({ok:true});
}));
app.get("/api/increments/preview",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const cycle=String(req.query.cycle||"").trim();
  if(!cycle)return res.status(400).json({error:"Review cycle is required"});
  const slabs=await getSlabs(req.user.company_id);
  const emps=await db.prepare("SELECT id,employee_code,name,department,designation,basic_salary,hra,other_allowances FROM employees WHERE company_id=? AND status='Active' ORDER BY name").all(req.user.company_id);
  const ratings=await db.prepare("SELECT employee_id,AVG(rating) r,COUNT(*) n FROM performance WHERE company_id=? AND cycle=? AND status='Finalized' GROUP BY employee_id").all(req.user.company_id,cycle);
  const rmap=Object.fromEntries(ratings.map(r=>[r.employee_id,Number(r.r)]));
  const done=await db.prepare("SELECT employee_id FROM increments WHERE company_id=? AND cycle=?").all(req.user.company_id,cycle);
  const doneSet=new Set(done.map(d=>d.employee_id));
  res.json({cycle,slabs,rows:emps.map(e=>{
    const ctc=(Number(e.basic_salary)||0)+(Number(e.hra)||0)+(Number(e.other_allowances)||0);
    const rating=rmap[e.id]??null;
    const pct=rating==null?0:pctFor(slabs,rating);
    let note="";
    if(doneSet.has(e.id))note="Already applied for this cycle";
    else if(rating==null)note="No finalized review for this cycle";
    else if(!ctc)note="Salary structure (CTC) not set";
    return {employee_id:e.id,employee_code:e.employee_code,name:e.name,department:e.department,designation:e.designation,
      rating:rating==null?null:+rating.toFixed(2),percent:pct,current_ctc:ctc,new_ctc:Math.round(ctc*(1+pct/100)),
      applicable:!note,note};
  })});
}));
app.post("/api/increments/apply",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const {cycle,effective_date,items}=req.body||{};
  if(!cycle||!effective_date||!Array.isArray(items)||!items.length)return res.status(400).json({error:"Cycle, effective date and at least one employee are required"});
  const company=await db.prepare("SELECT name,smtp_user,smtp_pass FROM companies WHERE id=?").get(req.user.company_id);
  const inr=n=>"₹"+Number(n).toLocaleString("en-IN");
  let applied=0,skipped=0;
  for(const it of items){
    const pct=Number(it.percent);
    if(!(pct>0&&pct<=100)){skipped++;continue}
    const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=? AND status='Active'").get(it.employee_id,req.user.company_id);
    if(!emp){skipped++;continue}
    if(await db.prepare("SELECT id FROM increments WHERE employee_id=? AND cycle=? AND company_id=?").get(emp.id,cycle,req.user.company_id)){skipped++;continue}
    const basic=Number(emp.basic_salary)||0,hra=Number(emp.hra)||0,other=Number(emp.other_allowances)||0;
    const oldCtc=basic+hra+other;
    if(!oldCtc){skipped++;continue}
    const f=1+pct/100;
    const nb=Math.round(basic*f),nh=Math.round(hra*f),no=Math.round(other*f);
    const rev=await db.prepare("SELECT AVG(rating) r FROM performance WHERE employee_id=? AND cycle=? AND status='Finalized'").get(emp.id,cycle);
    await db.prepare("UPDATE employees SET basic_salary=?,hra=?,other_allowances=? WHERE id=?").run(nb,nh,no,emp.id);
    await db.prepare("INSERT INTO increments(company_id,employee_id,cycle,rating,percent,old_ctc,new_ctc,effective_date,applied_by) VALUES(?,?,?,?,?,?,?,?,?)")
      .run(req.user.company_id,emp.id,cycle,rev?.r!=null?Number(rev.r):null,pct,oldCtc,nb+nh+no,effective_date,req.user.username);
    applied++;
    if(emp.email){
      sendMail(emp.email,`Salary revision — ${cycle}`,layout("Your salary revision",
        `<p>Hi ${esc(emp.name)},</p><p>Based on your performance review for <b>${esc(cycle)}</b>, your salary has been revised by <b>${pct}%</b>, effective <b>${esc(effective_date)}</b>.</p>
         <table style="width:100%;border-collapse:collapse;margin-top:10px">
         <tr><td style="padding:6px;border:1px solid #e5e7eb">Previous monthly CTC</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(oldCtc)}</td></tr>
         <tr><td style="padding:6px;border:1px solid #e5e7eb"><b>Revised monthly CTC</b></td><td style="padding:6px;border:1px solid #e5e7eb"><b>${inr(nb+nh+no)}</b></td></tr></table>
         <p style="font-size:12px;color:#64748b">Congratulations and thank you for your contribution.</p>`),
        {smtp_user:company?.smtp_user,smtp_pass:company?.smtp_pass,name:company?.name}).catch(()=>{});
    }
  }
  await audit(req,"APPLY","INCREMENT",`${cycle}: ${applied} applied, ${skipped} skipped`);
  res.json({ok:true,applied,skipped});
}));
app.get("/api/increments",auth,requireCompany,roles("Super Admin","HR Admin","Director","Finance"),wrap(async(req,res)=>{
  res.json(await db.prepare(`SELECT i.*,e.employee_code,e.name FROM increments i JOIN employees e ON e.id=i.employee_id WHERE i.company_id=? ORDER BY i.id DESC LIMIT 500`).all(req.user.company_id));
}));

app.get("/api/assets",auth,requireCompany,wrap(async(req,res)=>res.json(await db.prepare(`SELECT a.*,e.name employee_name,e.employee_code FROM assets a LEFT JOIN employees e ON e.id=a.employee_id WHERE a.company_id=? ORDER BY a.id DESC`).all(req.user.company_id))));
app.post("/api/assets",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await db.prepare("INSERT INTO assets(company_id,asset_code,name,category,serial_no,status,employee_id,issued_date) VALUES(?,?,?,?,?,?,?,?)").run(req.user.company_id,x.asset_code,x.name,x.category,x.serial_no,x.status||"Available",x.employee_id||null,x.issued_date||null);
    res.json({id:r.lastInsertRowid});
  }catch(e){res.status(400).json({error:/duplicate key|unique/i.test(e.message)?"Asset code already exists":e.message})}
}));

app.get("/api/expenses",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT x.*,e.name employee_name,e.employee_code FROM expenses x JOIN employees e ON e.id=x.employee_id WHERE x.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND x.employee_id=?";p.push(req.user.employee_id)}
  else if(req.user.role==="Manager" && req.user.employee_id){q+=" AND (e.reporting_manager_id=? OR e.id=?)";p.push(req.user.employee_id,req.user.employee_id)}
  q+=" ORDER BY x.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/expenses",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const x=req.body;const r=await db.prepare("INSERT INTO expenses(company_id,employee_id,category,amount,expense_date,description) VALUES(?,?,?,?,?,?)").run(req.user.company_id,eid,x.category,Number(x.amount)||0,x.expense_date,x.description||"");
  res.json({id:r.lastInsertRowid});
}));
app.post("/api/expenses/:id/status",auth,requireCompany,roles("Super Admin","HR Admin","Finance","Manager"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE expenses SET status=? WHERE id=? AND company_id=?").run(req.body.status,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Expense not found"});res.json({ok:true});
}));

app.get("/api/documents",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT d.id,d.company_id,d.employee_id,d.name,d.doc_type,d.expiry_date,d.status,d.file_name,d.file_mime,d.created_at,e.name employee_name,e.employee_code FROM documents d JOIN employees e ON e.id=d.employee_id WHERE d.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND d.employee_id=?";p.push(req.user.employee_id)}q+=" ORDER BY d.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.get("/api/documents/:id/file",auth,requireCompany,wrap(async(req,res)=>{
  const d=await db.prepare("SELECT * FROM documents WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!d||!d.file_data)return res.status(404).json({error:"File not found"});
  if(req.user.role==="Employee" && d.employee_id!==req.user.employee_id)return res.status(403).json({error:"Permission denied"});
  res.setHeader("Content-Type",d.file_mime||"application/octet-stream");
  res.setHeader("Content-Disposition",`inline; filename="${(d.file_name||"document").replace(/[^\w.\-]/g,"_")}"`);
  res.send(d.file_data);
}));
app.post("/api/documents",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const x=req.body;
  let fileBuf=null;
  if(x.file_base64){
    if(x.file_base64.length>13000000)return res.status(400).json({error:"File too large (max ~10MB)"});
    fileBuf=Buffer.from(x.file_base64,"base64");
  }
  const r=await db.prepare("INSERT INTO documents(company_id,employee_id,name,doc_type,expiry_date,status,file_name,file_mime,file_data,uploaded_by) VALUES(?,?,?,?,?,?,?,?,?,?)")
    .run(req.user.company_id,eid,x.name,x.doc_type,x.expiry_date,x.status||"Active",x.file_name||null,x.file_mime||null,fileBuf,req.user.id);
  res.json({id:r.lastInsertRowid});
}));

app.get("/api/announcements",auth,requireCompany,wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM announcements WHERE company_id=? ORDER BY id DESC").all(req.user.company_id))));
async function publishAnnouncement(companyId,title,body,audience="All"){
  const r=await db.prepare("INSERT INTO announcements(company_id,title,body,audience) VALUES(?,?,?,?)").run(companyId,title,body,audience);
  (async()=>{
    const emails=(await db.prepare("SELECT DISTINCT email FROM employees WHERE company_id=? AND status='Active' AND email IS NOT NULL AND email<>''").all(companyId)).map(r=>r.email);
    if(!emails.length)return;
    const sender=await companySender(companyId);
    const html=layout(title,`<p>${esc(body||"").replace(/\n/g,"<br>")}</p><p style="font-size:12px;color:#64748b">Audience: ${esc(audience)}</p>`);
    for(const email of emails) await sendMail(email,`Announcement: ${title}`,html,sender);
  })().catch(e=>console.error("Announcement email batch failed:",e.message));
  return r.lastInsertRowid;
}
app.post("/api/announcements",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  if(!x.title||!String(x.title).trim())return res.status(400).json({error:"Title is required"});
  const id=await publishAnnouncement(req.user.company_id,String(x.title).trim(),x.body||"",x.audience||"All");
  res.json({id});
}));

/* ---------------- Employee of the Month (data-driven suggestion) ---------------- */
const OFFICE_START_MIN=9*60+30, GRACE_MIN=15;
function lateFromFirstIn(firstIn){
  if(!firstIn||firstIn.length<16)return false;
  const hh=Number(firstIn.slice(11,13)),mm=Number(firstIn.slice(14,16));
  if(isNaN(hh)||isNaN(mm))return false;
  return hh*60+mm>OFFICE_START_MIN+GRACE_MIN;
}
async function computeEom(companyId,month){
  const [y,m]=month.split("-").map(Number);
  const daysInMonth=new Date(y,m,0).getDate();
  const now=new Date();
  const isCurrent=now.getFullYear()===y&&now.getMonth()+1===m;
  const elapsed=isCurrent?now.getDate():daysInMonth;
  let workingDays=0;
  for(let d=1;d<=elapsed;d++) if(new Date(y,m-1,d).getDay()!==0) workingDays++;
  const monthEnd=`${month}-${String(daysInMonth).padStart(2,"0")}`;
  const emps=await db.prepare("SELECT id,employee_code,name,department,designation,joining_date FROM employees WHERE company_id=? AND status='Active'").all(companyId);
  const att=await db.prepare("SELECT employee_id,status,first_in FROM attendance WHERE company_id=? AND work_date LIKE ?").all(companyId,month+"%");
  const leaves=await db.prepare("SELECT employee_id,COALESCE(SUM(days),0) d FROM leave_requests WHERE company_id=? AND status='Approved' AND category='Leave' AND from_date LIKE ? GROUP BY employee_id").all(companyId,month+"%");
  const leaveMap=Object.fromEntries(leaves.map(l=>[l.employee_id,Number(l.d)]));
  const reviews=await db.prepare("SELECT employee_id,rating,cycle FROM performance WHERE company_id=? AND status='Finalized' ORDER BY finalized_at DESC NULLS LAST,id DESC").all(companyId);
  const latestRating={};
  for(const r of reviews) if(latestRating[r.employee_id]==null) latestRating[r.employee_id]={rating:Number(r.rating),cycle:r.cycle};
  const prevMonths=[];
  for(let i=1;i<=3;i++){const d=new Date(y,m-1-i,1);prevMonths.push(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`)}
  const recent=await db.prepare("SELECT employee_id FROM awards WHERE company_id=? AND award_type='Employee of the Month' AND period=ANY(?)").all(companyId,prevMonths);
  const recentSet=new Set(recent.map(r=>r.employee_id));
  const byEmp={};
  for(const a of att){(byEmp[a.employee_id]=byEmp[a.employee_id]||[]).push(a)}
  const out=[];
  for(const e of emps){
    if(e.joining_date && e.joining_date>monthEnd) continue;
    const rows=byEmp[e.id]||[];
    const present=rows.filter(r=>r.status==="Present");
    const half=rows.filter(r=>r.status==="Half Day").length;
    const presentEq=present.length+half*0.5;
    const eligible=Math.max(1,workingDays-(leaveMap[e.id]||0));
    const attendance=Math.min(1,presentEq/eligible);
    const lateDays=present.filter(r=>lateFromFirstIn(r.first_in)).length;
    const punctuality=present.length?1-lateDays/present.length:0;
    const rv=latestRating[e.id];
    const perf=rv?rv.rating/5:null;
    const wP=perf==null?0:0.4,wA=perf==null?0.58:0.35,wT=perf==null?0.42:0.25;
    const score=+((100*(wP*(perf||0)+wA*attendance+wT*punctuality))).toFixed(1);
    const reasons=[`${Math.round(attendance*100)}% attendance (${presentEq} of ${eligible} working days)`,`${lateDays} late arrival${lateDays===1?"":"s"}`];
    reasons.push(rv?`Latest review rating ${rv.rating.toFixed(1)}/5 (${rv.cycle})`:"No finalized performance review");
    out.push({employee_id:e.id,employee_code:e.employee_code,name:e.name,department:e.department,designation:e.designation,
      score,attendance_pct:Math.round(attendance*100),late_days:lateDays,rating:rv?rv.rating:null,
      has_attendance_data:rows.length>0,recent_winner:recentSet.has(e.id),reasons});
  }
  out.sort((a,b)=>b.score-a.score);
  return {month,working_days:workingDays,candidates:out};
}
app.get("/api/eom",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const month=req.query.month||new Date().toISOString().slice(0,7);
  if(!/^\d{4}-\d{2}$/.test(month))return res.status(400).json({error:"Month must be in YYYY-MM format"});
  const r=await computeEom(req.user.company_id,month);
  const declared=await db.prepare("SELECT a.*,e.name,e.employee_code FROM awards a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=? AND a.award_type='Employee of the Month' AND a.period=?").get(req.user.company_id,month);
  res.json({...r,candidates:r.candidates.slice(0,10),declared:declared||null});
}));
app.post("/api/eom/declare",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const {month,employee_id,note}=req.body||{};
  if(!/^\d{4}-\d{2}$/.test(month||""))return res.status(400).json({error:"Month must be in YYYY-MM format"});
  const emp=await db.prepare("SELECT id,name FROM employees WHERE id=? AND company_id=? AND status='Active'").get(employee_id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  if(await db.prepare("SELECT id FROM awards WHERE company_id=? AND award_type='Employee of the Month' AND period=?").get(req.user.company_id,month))
    return res.status(400).json({error:"Employee of the Month is already declared for this month"});
  const r=await computeEom(req.user.company_id,month);
  const c=r.candidates.find(x=>x.employee_id===emp.id);
  await db.prepare("INSERT INTO awards(company_id,employee_id,award_type,period,score,note,declared_by) VALUES(?,?,?,?,?,?,?)")
    .run(req.user.company_id,emp.id,"Employee of the Month",month,c?c.score:null,note||"",req.user.username);
  const label=new Date(month+"-01").toLocaleDateString("en-IN",{month:"long",year:"numeric"});
  await publishAnnouncement(req.user.company_id,`Employee of the Month — ${label}`,`Congratulations to ${emp.name} for being recognised as Employee of the Month for ${label}.${note?"\n\n"+note:""}`,"All");
  await audit(req,"DECLARE","EOM",`${month}: ${emp.name}`);
  res.json({ok:true});
}));
/* ---------------- MIS (management information) ---------------- */
app.get("/api/mis",auth,requireCompany,roles("Super Admin","HR Admin","Director","Finance"),wrap(async(req,res)=>{
  const cid=req.user.company_id;
  const month=req.query.month||new Date().toISOString().slice(0,7);
  if(!/^\d{4}-\d{2}$/.test(month))return res.status(400).json({error:"Month must be in YYYY-MM format"});
  const N=v=>Number(v)||0;
  const dept="COALESCE(NULLIF(e.department,''),'Unassigned')";
  const today=new Date();
  const d30=new Date(today.getTime()-30*86400000).toISOString().slice(0,10);
  const d365=new Date(today.getTime()-365*86400000).toISOString().slice(0,10);

  const headcount=await db.prepare("SELECT status,COUNT(*) c FROM employees WHERE company_id=? GROUP BY status").all(cid);
  const active=N(headcount.find(h=>h.status==="Active")?.c);
  const byDept=(await db.prepare(`SELECT ${dept} department,COUNT(*) c FROM employees e WHERE e.company_id=? AND e.status='Active' GROUP BY 1 ORDER BY 2 DESC`).all(cid)).map(r=>({department:r.department,count:N(r.c)}));
  const newJoiners=N((await db.prepare("SELECT COUNT(*) c FROM employees WHERE company_id=? AND status='Active' AND joining_date>=?").get(cid,d30)).c);
  const exits12=N((await db.prepare("SELECT COUNT(*) c FROM exit_requests WHERE company_id=? AND status='Approved' AND last_working_date>=?").get(cid,d365)).c);
  const openExits=N((await db.prepare("SELECT COUNT(*) c FROM exit_requests WHERE company_id=? AND status='Pending'").get(cid)).c);

  const attRows=await db.prepare(`SELECT ${dept} department,a.status,a.first_in FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=? AND a.work_date LIKE ?`).all(cid,month+"%");
  const attDept={};
  let tPresent=0,tAbsent=0,tHalf=0,tLate=0;
  for(const r of attRows){
    const d=attDept[r.department]=attDept[r.department]||{department:r.department,present:0,absent:0,half_day:0,late:0};
    if(r.status==="Present"){d.present++;tPresent++;if(lateFromFirstIn(r.first_in)){d.late++;tLate++}}
    else if(r.status==="Absent"){d.absent++;tAbsent++}
    else if(r.status==="Half Day"){d.half_day++;tHalf++}
  }

  const leaveByType=(await db.prepare("SELECT leave_type,COALESCE(SUM(days),0) d FROM leave_requests WHERE company_id=? AND status='Approved' AND category='Leave' AND from_date LIKE ? GROUP BY leave_type ORDER BY 2 DESC").all(cid,month+"%")).map(r=>({leave_type:r.leave_type,days:N(r.d)}));
  const wfhDays=N((await db.prepare("SELECT COALESCE(SUM(days),0) d FROM leave_requests WHERE company_id=? AND status='Approved' AND category='WFH' AND from_date LIKE ?").get(cid,month+"%")).d);
  const pendingApprovals=N((await db.prepare("SELECT COUNT(*) c FROM leave_requests WHERE company_id=? AND status='Pending'").get(cid)).c);

  const pay=await db.prepare("SELECT COUNT(*) n,COALESCE(SUM(gross),0) gross,COALESCE(SUM(net),0) net,COALESCE(SUM(pf_employee),0) pf,COALESCE(SUM(esic_employee),0) esic,COALESCE(SUM(tds),0) tds,COALESCE(SUM(lop),0) lop FROM payroll WHERE company_id=? AND month=?").get(cid,month);
  const payDept=(await db.prepare(`SELECT ${dept} department,COALESCE(SUM(p.net),0) net FROM payroll p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=? AND p.month=? GROUP BY 1 ORDER BY 2 DESC`).all(cid,month)).map(r=>({department:r.department,net:N(r.net)}));
  const payTrend=(await db.prepare("SELECT month,COALESCE(SUM(net),0) net,COALESCE(SUM(gross),0) gross FROM payroll WHERE company_id=? GROUP BY month ORDER BY month DESC LIMIT 6").all(cid)).reverse().map(r=>({month:r.month,net:N(r.net),gross:N(r.gross)}));

  const cand=(await db.prepare("SELECT status,COUNT(*) c FROM candidates WHERE company_id=? GROUP BY status").all(cid)).map(r=>({status:r.status,count:N(r.c)}));
  const expStatus=(await db.prepare("SELECT status,COUNT(*) c,COALESCE(SUM(amount),0) a FROM expenses WHERE company_id=? AND expense_date LIKE ? GROUP BY status").all(cid,month+"%")).map(r=>({status:r.status,count:N(r.c),amount:N(r.a)}));
  const expCat=(await db.prepare("SELECT COALESCE(NULLIF(category,''),'Other') category,COALESCE(SUM(amount),0) a FROM expenses WHERE company_id=? AND expense_date LIKE ? GROUP BY 1 ORDER BY 2 DESC").all(cid,month+"%")).map(r=>({category:r.category,amount:N(r.a)}));

  const ratings=(await db.prepare("SELECT rating FROM performance WHERE company_id=? AND status='Finalized'").all(cid)).map(r=>Number(r.rating));
  const dist=[{label:"Below 2",count:0},{label:"2 to 3",count:0},{label:"3 to 4",count:0},{label:"4 to 5",count:0}];
  for(const r of ratings){dist[r<2?0:r<3?1:r<4?2:3].count++}
  const openTickets=N((await db.prepare("SELECT COUNT(*) c FROM tickets WHERE company_id=? AND status<>'Closed'").get(cid)).c);
  const assets=(await db.prepare("SELECT status,COUNT(*) c FROM assets WHERE company_id=? GROUP BY status").all(cid)).map(r=>({status:r.status,count:N(r.c)}));

  res.json({month,
    headcount:{active,inactive:N(headcount.find(h=>h.status==="Inactive")?.c),byDepartment:byDept,newJoiners30:newJoiners,exits12m:exits12,pendingExits:openExits,
      attritionPct:active+exits12?+(exits12/(active+exits12)*100).toFixed(1):0},
    attendance:{present:tPresent,absent:tAbsent,halfDay:tHalf,late:tLate,byDepartment:Object.values(attDept)},
    leave:{byType:leaveByType,wfhDays,pendingApprovals},
    payroll:{processed:N(pay.n),gross:N(pay.gross),net:N(pay.net),pf:N(pay.pf),esic:N(pay.esic),tds:N(pay.tds),lop:N(pay.lop),byDepartment:payDept,trend:payTrend},
    recruitment:cand,
    expenses:{byStatus:expStatus,byCategory:expCat},
    performance:{reviews:ratings.length,average:ratings.length?+(ratings.reduce((a,b)=>a+b,0)/ratings.length).toFixed(2):null,distribution:dist},
    helpdesk:{open:openTickets},assets});
}));

app.get("/api/awards",auth,requireCompany,wrap(async(req,res)=>{
  res.json(await db.prepare("SELECT a.*,e.name,e.employee_code,e.department FROM awards a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=? ORDER BY a.period DESC LIMIT 24").all(req.user.company_id));
}));

app.get("/api/tickets",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT t.*,e.name employee_name,e.employee_code FROM tickets t JOIN employees e ON e.id=t.employee_id WHERE t.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND t.employee_id=?";p.push(req.user.employee_id)}q+=" ORDER BY t.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/tickets",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const x=req.body;const r=await db.prepare("INSERT INTO tickets(company_id,employee_id,subject,description,priority) VALUES(?,?,?,?,?)").run(req.user.company_id,eid,x.subject,x.description,x.priority||"Medium");res.json({id:r.lastInsertRowid});
}));
app.post("/api/tickets/:id/status",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE tickets SET status=?,assigned_to=? WHERE id=? AND company_id=?").run(req.body.status,req.user.username,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Ticket not found"});res.json({ok:true})
}));

app.get("/api/exits",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT x.*,e.name employee_name,e.employee_code FROM exit_requests x JOIN employees e ON e.id=x.employee_id WHERE x.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND x.employee_id=?";p.push(req.user.employee_id)}q+=" ORDER BY x.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/exits",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const x=req.body;
  const r=await db.prepare("INSERT INTO exit_requests(company_id,employee_id,resignation_date,last_working_date,reason) VALUES(?,?,?,?,?)").run(req.user.company_id,eid,x.resignation_date,x.last_working_date,x.reason||"");res.json({id:r.lastInsertRowid});
}));
app.post("/api/exits/:id/status",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  const r=await db.prepare("UPDATE exit_requests SET status=COALESCE(?,status),clearance=COALESCE(?,clearance),fnf_status=COALESCE(?,fnf_status) WHERE id=? AND company_id=?")
    .run(x.status||null,x.clearance||null,x.fnf_status||null,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Exit request not found"});
  if(x.status==="Approved"){
    const ex=await db.prepare("SELECT employee_id FROM exit_requests WHERE id=?").get(req.params.id);
    if(ex)await db.prepare("UPDATE employees SET status='Inactive' WHERE id=? AND company_id=?").run(ex.employee_id,req.user.company_id);
  }
  await audit(req,"STATUS","EXIT",req.params.id);res.json({ok:true});
}));

app.get("/api/biometric/devices",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM biometric_devices WHERE company_id=? ORDER BY id DESC").all(req.user.company_id))));
app.post("/api/biometric/devices",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;const apiKey=crypto.randomBytes(24).toString("hex");
  const r=await db.prepare("INSERT INTO biometric_devices(company_id,name,model,serial_no,branch,ip,port,protocol,api_key) VALUES(?,?,?,?,?,?,?,?,?)").run(req.user.company_id,x.name,x.model,x.serial_no||"",x.branch,x.ip,Number(x.port)||4370,x.protocol||"ZKTeco/eSSL (LAN)",apiKey);res.json({id:r.lastInsertRowid,api_key:apiKey});
}));
app.post("/api/biometric/devices/:id/rotate-key",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const apiKey=crypto.randomBytes(24).toString("hex");
  const r=await db.prepare("UPDATE biometric_devices SET api_key=? WHERE id=? AND company_id=?").run(apiKey,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Device not found"});
  res.json({ok:true,api_key:apiKey});
}));
app.put("/api/biometric/devices/:id",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  const r=await db.prepare("UPDATE biometric_devices SET name=?,model=?,serial_no=?,branch=?,ip=?,port=? WHERE id=? AND company_id=?")
    .run(x.name,x.model,x.serial_no||"",x.branch,x.ip,Number(x.port)||4370,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Device not found"});
  res.json({ok:true});
}));

function pad2(n){return String(n).padStart(2,"0")}
function localISO(d){return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}:${pad2(d.getSeconds())}`}

async function ingestPunch(companyId,deviceId,biometricId,punchTimeStr){
  await db.prepare("INSERT OR IGNORE INTO punches(company_id,biometric_id,punch_time,punch_type,device_id,raw_payload) VALUES(?,?,?,?,?,?)")
    .run(companyId,biometricId,punchTimeStr,"AUTO",deviceId||null,"{}");
  const e=await db.prepare("SELECT * FROM employees WHERE biometric_id=? AND company_id=?").get(biometricId,companyId);
  if(!e)return null;
  const d=punchTimeStr.slice(0,10),a=await db.prepare("SELECT * FROM attendance WHERE employee_id=? AND work_date=?").get(e.id,d);
  if(!a)await db.prepare("INSERT INTO attendance(company_id,employee_id,work_date,first_in,status,source) VALUES(?,?,?,?,?,?)").run(companyId,e.id,d,punchTimeStr,"Present","eSSL");
  else{
    const firstIn=a.first_in&&a.first_in<punchTimeStr?a.first_in:(a.first_in||punchTimeStr);
    const lastOut=(!a.last_out||punchTimeStr>a.last_out)&&punchTimeStr!==firstIn?punchTimeStr:a.last_out;
    await db.prepare("UPDATE attendance SET first_in=?,last_out=?,source='eSSL' WHERE id=?").run(firstIn<punchTimeStr?firstIn:punchTimeStr,lastOut,a.id);
  }
  return e.employee_code;
}

app.post("/api/biometric/punch",auth,requireCompany,wrap(async(req,res)=>{
  const x=req.body;if(!x.biometric_id||!x.punch_time)return res.status(400).json({error:"biometric_id and punch_time required"});
  const code=await ingestPunch(req.user.company_id,x.device_id||null,x.biometric_id,x.punch_time);
  res.json({ok:true,employee:code});
}));

app.post("/api/biometric/devices/:id/sync",auth,requireCompany,roles("Super Admin","HR Admin"),async(req,res)=>{
  const device=await db.prepare("SELECT * FROM biometric_devices WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!device)return res.status(404).json({error:"Device not found"});
  if(!device.ip)return res.status(400).json({error:"Device IP address is not configured"});
  let zk;
  try{
    zk=new ZKLib(device.ip,Number(device.port)||4370,10000,4000);
    await zk.createSocket();
    const result=await zk.getAttendances();
    const allLogs=result?.data||[];
    const days=Math.max(1,Number(req.body?.days)||4);
    const cutoff=new Date();cutoff.setHours(0,0,0,0);cutoff.setDate(cutoff.getDate()-(days-1));
    const logs=allLogs.filter(l=>l.recordTime instanceof Date && l.recordTime>=cutoff);
    let matched=0,unmatched=0;
    for(const log of logs){
      const biometricId=String(log.deviceUserId||"").trim();
      const t=log.recordTime instanceof Date?localISO(log.recordTime):null;
      if(!biometricId||!t)continue;
      const code=await ingestPunch(req.user.company_id,device.id,biometricId,t);
      if(code)matched++;else unmatched++;
    }
    try{await zk.disconnect()}catch{}
    await db.prepare("UPDATE biometric_devices SET status='Connected',last_sync=?,last_error=NULL WHERE id=?").run(new Date().toISOString(),device.id);
    await audit(req,"SYNC","BIOMETRIC",`${device.name}: ${allLogs.length} logs on device, last ${days} day(s) = ${logs.length} logs, ${matched} matched, ${unmatched} unmatched biometric IDs`);
    res.json({ok:true,totalOnDevice:allLogs.length,totalLogs:logs.length,matched,unmatched,days});
  }catch(e){
    try{await zk?.disconnect()}catch{}
    const msg=e?.err?.code?`${e.err.code} (${e.command||"connection"} to ${e.ip||device.ip})`:(e?.message||e?.err?.message||String(e));
    await db.prepare("UPDATE biometric_devices SET status='Not Reachable',last_error=? WHERE id=?").run(msg,device.id);
    res.status(502).json({error:"Could not connect to device: "+msg});
  }
});

// Used by the standalone local sync-agent (sync-agent.js) so it can push punches from a LAN-only
// device to a HRMS server hosted anywhere on the internet, without needing a logged-in session.
app.post("/api/biometric/ingest",wrap(async(req,res)=>{
  const {device_id,api_key,records}=req.body||{};
  if(!device_id||!api_key)return res.status(400).json({error:"device_id and api_key required"});
  const device=await db.prepare("SELECT * FROM biometric_devices WHERE id=?").get(device_id);
  if(!device||!device.api_key||device.api_key!==api_key)return res.status(401).json({error:"Invalid device_id or api_key"});
  if(!Array.isArray(records))return res.status(400).json({error:"records array required"});
  let matched=0,unmatched=0;
  for(const r of records){
    const biometricId=String(r.biometric_id||"").trim();
    const t=r.punch_time;
    if(!biometricId||!t)continue;
    const code=await ingestPunch(device.company_id,device.id,biometricId,t);
    if(code)matched++;else unmatched++;
  }
  await db.prepare("UPDATE biometric_devices SET status='Connected',last_sync=?,last_error=NULL WHERE id=?").run(new Date().toISOString(),device.id);
  await db.prepare("INSERT INTO audit_logs(company_id,user_id,action,module,details) VALUES(?,?,?,?,?)")
    .run(device.company_id,null,"SYNC","BIOMETRIC",`${device.name} (agent push): ${records.length} logs, ${matched} matched, ${unmatched} unmatched biometric IDs`);
  res.json({ok:true,received:records.length,matched,unmatched});
}));

/* ---------------- Data export (Excel / CSV) ---------------- */
function safeCell(v){
  if(v==null)return "";
  if(typeof v==="number")return v;
  const s=String(v);
  return /^[=+\-@\t\r]/.test(s)?"'"+s:s; // neutralise spreadsheet formula injection
}
async function sendTable(res,format,filename,columns,rows){
  const data=rows.map(r=>Object.fromEntries(columns.map(c=>[c.key,safeCell(c.fmt?c.fmt(r[c.key],r):r[c.key])])));
  if(format==="csv"){
    const q=v=>{const s=String(v);return /[",\r\n]/.test(s)?`"${s.replace(/"/g,'""')}"`:s};
    const lines=[columns.map(c=>q(c.header)).join(","),...data.map(r=>columns.map(c=>q(r[c.key])).join(","))];
    res.setHeader("Content-Type","text/csv; charset=utf-8");
    res.setHeader("Content-Disposition",`attachment; filename="${filename}.csv"`);
    return res.send("﻿"+lines.join("\r\n"));
  }
  const wb=new ExcelJS.Workbook();
  const ws=wb.addWorksheet("Data");
  ws.columns=columns.map(c=>({header:c.header,key:c.key,width:c.width||18}));
  ws.addRows(data);
  const head=ws.getRow(1);
  head.font={bold:true,color:{argb:"FFFFFFFF"}};
  head.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF4F46E5"}};
  head.alignment={vertical:"middle"};
  ws.views=[{state:"frozen",ySplit:1}];
  ws.autoFilter={from:{row:1,column:1},to:{row:1,column:columns.length}};
  res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  res.setHeader("Content-Disposition",`attachment; filename="${filename}.xlsx"`);
  await wb.xlsx.write(res);
  res.end();
}
const yn=v=>v?"Yes":"No";
const HR_ROLES=["Super Admin","HR Admin","Director"];
const teamFilter=(req,q,p,alias="e")=>{
  if(req.user.role==="Manager"&&req.user.employee_id){q+=` AND (${alias}.reporting_manager_id=? OR ${alias}.id=?)`;p.push(req.user.employee_id,req.user.employee_id)}
  return q;
};
const EXPORTS={
  employees:{roles:[...HR_ROLES,"Finance","Manager"],
    columns:req=>{
      const base=[{header:"Employee Code",key:"employee_code"},{header:"Name",key:"name",width:26},{header:"Email",key:"email",width:28},{header:"Phone",key:"phone"},{header:"Department",key:"department"},{header:"Designation",key:"designation"},{header:"Reporting Manager",key:"reporting_manager",width:24},{header:"Branch",key:"branch"},{header:"Joining Date",key:"joining_date"},{header:"Status",key:"status"}];
      if(req.user.role==="Manager")return base;
      return [...base,{header:"Date of Birth",key:"date_of_birth"},{header:"Biometric ID",key:"biometric_id"},{header:"Basic (monthly)",key:"basic_salary"},{header:"HRA (monthly)",key:"hra"},{header:"Other Allowances (monthly)",key:"other_allowances",width:24},
        {header:"PF Applicable",key:"pf_applicable",fmt:yn},{header:"ESIC Applicable",key:"esic_applicable",fmt:yn},{header:"PF Number",key:"pf_number"},{header:"UAN",key:"uan_number"},{header:"ESIC Number",key:"esic_number"},{header:"PAN",key:"pan_number"},
        {header:"Bank Name",key:"bank_name"},{header:"Bank Account",key:"bank_account",width:22},{header:"IFSC",key:"ifsc"}];
    },
    query:req=>{let p=[req.user.company_id];let q="SELECT e.*,m.name reporting_manager FROM employees e LEFT JOIN employees m ON m.id=e.reporting_manager_id WHERE e.company_id=?";q=teamFilter(req,q,p);return [q+" ORDER BY e.employee_code",p]}},
  attendance:{roles:[...HR_ROLES,"Manager"],
    columns:()=>[{header:"Date",key:"work_date"},{header:"Employee Code",key:"employee_code"},{header:"Name",key:"name",width:26},{header:"Department",key:"department"},{header:"First In",key:"first_in",width:22},{header:"Last Out",key:"last_out",width:22},{header:"Status",key:"status"},{header:"Late (min)",key:"late_minutes"},{header:"Source",key:"source"}],
    query:req=>{let p=[req.user.company_id];let q="SELECT a.*,e.employee_code,e.name,e.department FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=?";
      if(req.query.from){q+=" AND a.work_date>=?";p.push(req.query.from)}if(req.query.to){q+=" AND a.work_date<=?";p.push(req.query.to)}
      q=teamFilter(req,q,p);return [q+" ORDER BY a.work_date DESC,e.employee_code",p]}},
  leaves:{roles:[...HR_ROLES,"Manager"],
    columns:()=>[{header:"Employee Code",key:"employee_code"},{header:"Name",key:"name",width:26},{header:"Category",key:"category"},{header:"Leave Type",key:"leave_type",width:22},{header:"From",key:"from_date"},{header:"To",key:"to_date"},{header:"Days",key:"days"},{header:"Status",key:"status"},{header:"Approved By",key:"approved_by"},{header:"Reason",key:"reason",width:34}],
    query:req=>{let p=[req.user.company_id];let q="SELECT l.*,e.employee_code,e.name FROM leave_requests l JOIN employees e ON e.id=l.employee_id WHERE l.company_id=?";q=teamFilter(req,q,p);return [q+" ORDER BY l.id DESC",p]}},
  payroll:{roles:[...HR_ROLES,"Finance"],
    columns:()=>[{header:"Month",key:"month"},{header:"Employee Code",key:"employee_code"},{header:"Name",key:"name",width:26},{header:"Gross",key:"gross"},{header:"LOP Days",key:"lop_days"},{header:"LOP Amount",key:"lop"},{header:"PF (Employee)",key:"pf_employee"},{header:"ESIC (Employee)",key:"esic_employee"},{header:"TDS",key:"tds"},{header:"Other Deductions",key:"deductions"},{header:"Overtime",key:"ot"},{header:"Net Pay",key:"net"},{header:"Status",key:"status"},{header:"Payslip No",key:"payslip_no",width:24},{header:"Bank Name",key:"bank_name"},{header:"Bank Account",key:"bank_account",width:22},{header:"IFSC",key:"ifsc"}],
    query:req=>{let p=[req.user.company_id];let q="SELECT p.*,e.employee_code,e.name,e.bank_name,e.bank_account,e.ifsc FROM payroll p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=?";
      if(req.query.month){q+=" AND p.month=?";p.push(req.query.month)}return [q+" ORDER BY p.month DESC,e.employee_code",p]}},
  expenses:{roles:[...HR_ROLES,"Finance","Manager"],
    columns:()=>[{header:"Employee Code",key:"employee_code"},{header:"Name",key:"employee_name",width:26},{header:"Category",key:"category"},{header:"Amount",key:"amount"},{header:"Date",key:"expense_date"},{header:"Description",key:"description",width:36},{header:"Status",key:"status"}],
    query:req=>{let p=[req.user.company_id];let q="SELECT x.*,e.name employee_name,e.employee_code FROM expenses x JOIN employees e ON e.id=x.employee_id WHERE x.company_id=?";q=teamFilter(req,q,p);return [q+" ORDER BY x.id DESC",p]}},
  performance:{roles:[...HR_ROLES,"Manager"],
    columns:()=>[{header:"Employee Code",key:"employee_code"},{header:"Name",key:"name",width:26},{header:"Cycle",key:"cycle"},{header:"Rating (1-5)",key:"rating"},{header:"Status",key:"status"},{header:"Reviewer",key:"reviewer"},{header:"Goals / KRA",key:"goals",width:40},{header:"Manager Comments",key:"manager_comments",width:40}],
    query:req=>{let p=[req.user.company_id];let q="SELECT p.*,e.employee_code,e.name FROM performance p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=?";
      if(req.user.role==="Manager"&&req.user.employee_id){q+=" AND e.reporting_manager_id=?";p.push(req.user.employee_id)}return [q+" ORDER BY p.id DESC",p]}},
  candidates:{roles:[...HR_ROLES,"Manager"],
    columns:()=>[{header:"Name",key:"name",width:26},{header:"Email",key:"email",width:28},{header:"Phone",key:"phone"},{header:"Position",key:"position",width:24},{header:"Status",key:"status"},{header:"Interview Date",key:"interview_date"},{header:"Notes",key:"notes",width:40}],
    query:req=>["SELECT * FROM candidates WHERE company_id=? ORDER BY id DESC",[req.user.company_id]]},
  assets:{roles:HR_ROLES,
    columns:()=>[{header:"Asset Code",key:"asset_code"},{header:"Asset",key:"name",width:26},{header:"Category",key:"category"},{header:"Serial No",key:"serial_no"},{header:"Status",key:"status"},{header:"Assigned To",key:"employee_name",width:26},{header:"Issued Date",key:"issued_date"}],
    query:req=>["SELECT a.*,e.name employee_name FROM assets a LEFT JOIN employees e ON e.id=a.employee_id WHERE a.company_id=? ORDER BY a.id DESC",[req.user.company_id]]},
  increments:{roles:[...HR_ROLES,"Finance"],
    columns:()=>[{header:"Employee Code",key:"employee_code"},{header:"Name",key:"name",width:26},{header:"Cycle",key:"cycle"},{header:"Rating",key:"rating"},{header:"Increment %",key:"percent"},{header:"Old Monthly CTC",key:"old_ctc"},{header:"New Monthly CTC",key:"new_ctc"},{header:"Effective Date",key:"effective_date"},{header:"Applied By",key:"applied_by"}],
    query:req=>["SELECT i.*,e.employee_code,e.name FROM increments i JOIN employees e ON e.id=i.employee_id WHERE i.company_id=? ORDER BY i.id DESC",[req.user.company_id]]}
};
app.get("/api/export/:dataset",auth,requireCompany,wrap(async(req,res)=>{
  const cfg=EXPORTS[req.params.dataset];
  if(!cfg)return res.status(404).json({error:"Unknown dataset"});
  if(!cfg.roles.includes(req.user.role))return res.status(403).json({error:"You do not have permission to download this data"});
  const [sql,params]=cfg.query(req);
  const rows=await db.prepare(sql).all(...params);
  const company=await db.prepare("SELECT code FROM companies WHERE id=?").get(req.user.company_id);
  const format=req.query.format==="csv"?"csv":"xlsx";
  await audit(req,"EXPORT","DATA",`${req.params.dataset} (${rows.length} rows, ${format})`);
  await sendTable(res,format,`${(company?.code||"company").toLowerCase()}_${req.params.dataset}_${new Date().toISOString().slice(0,10)}`,cfg.columns(req),rows);
}));

app.get("/api/audit",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>res.json(await db.prepare(`SELECT a.*,u.username FROM audit_logs a LEFT JOIN users u ON u.id=a.user_id WHERE a.company_id=? ORDER BY a.id DESC LIMIT 300`).all(req.user.company_id))));

app.get("/api/settings",auth,requireCompany,roles("Super Admin","HR Admin"),(req,res)=>res.json({
  office_start:"09:30",grace_minutes:15,weekly_off:"Sunday",half_day_after:"13:30",
  payroll_cutoff:25,leave_approval:"Manager → HR",biometric_provider:"eSSL"
}));
app.use((req,res)=>res.sendFile(path.join(__dirname,"..","public","index.html")));

initSchema().then(seed).then(()=>{
  app.listen(PORT,()=>console.log(`BMS HRMS running on http://localhost:${PORT}`));
}).catch(e=>{
  console.error("Failed to initialize database:",e);
  process.exit(1);
});
