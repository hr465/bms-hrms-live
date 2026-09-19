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
 smtp_user TEXT, smtp_pass TEXT, policy_agreement_text TEXT, increment_policy TEXT, letter_template TEXT, custom_domain TEXT
);
CREATE TABLE IF NOT EXISTS letters(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER NOT NULL,letter_type TEXT DEFAULT 'Appointment Letter',
 ref_no TEXT,content TEXT,issued_by TEXT,issued_at TEXT DEFAULT CURRENT_TIMESTAMP,emailed_at TEXT
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
CREATE TABLE IF NOT EXISTS images(
 kind TEXT NOT NULL,ref_id INTEGER NOT NULL,company_id INTEGER,mime TEXT,data TEXT,updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
 PRIMARY KEY(kind,ref_id)
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
for(const col of ["serial_no TEXT","last_error TEXT","api_key TEXT","last_seen TEXT","adms_stamp TEXT"]){
  try{await db.exec(`ALTER TABLE biometric_devices ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["email TEXT"]){
  try{await db.exec(`ALTER TABLE users ADD COLUMN ${col}`)}catch(e){}
}
for(const col of ["last_digest TEXT","work_timing TEXT","smtp_user TEXT","smtp_pass TEXT","policy_agreement_text TEXT","increment_policy TEXT","letter_template TEXT","custom_domain TEXT"]){
  try{await db.exec(`ALTER TABLE companies ADD COLUMN ${col}`)}catch(e){}
}
try{await db.exec("ALTER TABLE employees ADD COLUMN work_timing TEXT")}catch(e){}
for(const col of ["in_loc TEXT","out_loc TEXT"]){try{await db.exec(`ALTER TABLE attendance ADD COLUMN ${col}`)}catch(e){}}
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

// A company can be served from its own domain (for example hr.example.com). The request's host
// decides which company's branding is shown and which users may sign in.
async function companyByHost(req){
  const h=String(req.hostname||"").toLowerCase();
  if(!h)return null;
  return (await db.prepare("SELECT id,name,code,industry FROM companies WHERE LOWER(custom_domain)=?").get(h))||null;
}
app.get("/api/branding",wrap(async(req,res)=>{
  const c=await companyByHost(req);
  res.json({company:c?{id:c.id,name:c.name,code:c.code,industry:c.industry}:null});
}));
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
  const hostCo=u?await companyByHost(req):null;
  if(!u||!verify(req.body.password||"",u.password_hash)||(hostCo&&(u.role==="Super Admin"||u.company_id!==hostCo.id))){
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
// Super Admin: list a company's logins and reset a password (a strong one is generated and shown once).
/* ---------------- Images: company logo and employee profile photo ---------------- */
const IMG_RE=/^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+\/=]+)$/;
async function storeImage(kind,refId,companyId,dataUrl){
  const m=IMG_RE.exec(String(dataUrl||""));
  if(!m)throw new Error("Please upload a PNG, JPG or WebP image");
  if(Buffer.byteLength(m[2],"base64")>400*1024)throw new Error("The image is too large (maximum 400 KB)");
  await db.prepare("INSERT INTO images(kind,ref_id,company_id,mime,data) VALUES(?,?,?,?,?) ON CONFLICT(kind,ref_id) DO UPDATE SET mime=excluded.mime,data=excluded.data,company_id=excluded.company_id,updated_at=CURRENT_TIMESTAMP")
    .run(kind,refId,companyId||null,m[1],m[2]);
}
async function sendImage(res,kind,refId){
  const r=await db.prepare("SELECT mime,data FROM images WHERE kind=? AND ref_id=?").get(kind,refId);
  if(!r)return res.status(404).end();
  res.set({"Content-Type":r.mime,"Cache-Control":"private, max-age=300","X-Content-Type-Options":"nosniff","Content-Security-Policy":"default-src 'none'"});
  res.send(Buffer.from(r.data,"base64"));
}
app.get("/api/images/company/:id",wrap(async(req,res)=>sendImage(res,"company",Number(req.params.id))));
app.post("/api/images/company/:id",auth,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const id=Number(req.params.id);
  if(req.user.role!=="Super Admin"&&req.user.company_id!==id)return res.status(403).json({error:"Permission denied"});
  if(!await db.prepare("SELECT id FROM companies WHERE id=?").get(id))return res.status(404).json({error:"Company not found"});
  try{await storeImage("company",id,id,req.body.data)}catch(e){return res.status(400).json({error:e.message})}
  await audit(req,"UPDATE","COMPANY_LOGO",String(id));res.json({ok:true});
}));
app.delete("/api/images/company/:id",auth,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const id=Number(req.params.id);
  if(req.user.role!=="Super Admin"&&req.user.company_id!==id)return res.status(403).json({error:"Permission denied"});
  await db.prepare("DELETE FROM images WHERE kind='company' AND ref_id=?").run(id);res.json({ok:true});
}));
app.get("/api/images/employee/:id",auth,wrap(async(req,res)=>{
  const e=await db.prepare("SELECT id,company_id FROM employees WHERE id=?").get(Number(req.params.id));
  const cid=req.user.company_id;
  if(!e||(req.user.role!=="Super Admin"&&e.company_id!==cid))return res.status(404).end();
  return sendImage(res,"employee",e.id);
}));
async function canEditPhoto(req,empId){
  const e=await db.prepare("SELECT id,company_id FROM employees WHERE id=?").get(empId);
  if(!e)return null;
  const hr=["Super Admin","HR Admin"].includes(req.user.role)&&e.company_id===(req.user.company_id||e.company_id);
  const self=req.user.employee_id===e.id;
  return hr||self?e:false;
}
app.post("/api/images/employee/:id",auth,wrap(async(req,res)=>{
  const e=await canEditPhoto(req,Number(req.params.id));
  if(e===null)return res.status(404).json({error:"Employee not found"});
  if(!e)return res.status(403).json({error:"Permission denied"});
  try{await storeImage("employee",e.id,e.company_id,req.body.data)}catch(err){return res.status(400).json({error:err.message})}
  res.json({ok:true});
}));
app.delete("/api/images/employee/:id",auth,wrap(async(req,res)=>{
  const e=await canEditPhoto(req,Number(req.params.id));
  if(e===null)return res.status(404).json({error:"Employee not found"});
  if(!e)return res.status(403).json({error:"Permission denied"});
  await db.prepare("DELETE FROM images WHERE kind='employee' AND ref_id=?").run(e.id);res.json({ok:true});
}));

app.get("/api/companies/:id/users",auth,roles("Super Admin"),wrap(async(req,res)=>{
  res.json(await db.prepare("SELECT u.id,u.username,u.role,u.active,e.name AS employee_name FROM users u LEFT JOIN employees e ON e.id=u.employee_id WHERE u.company_id=? ORDER BY u.id").all(req.params.id));
}));
app.post("/api/users/:id/reset-password",auth,roles("Super Admin"),wrap(async(req,res)=>{
  const u=await db.prepare("SELECT id,username,role,company_id FROM users WHERE id=?").get(req.params.id);
  if(!u||u.role==="Super Admin"||!u.company_id)return res.status(404).json({error:"User not found"});
  const A="ABCDEFGHJKMNPQRSTUVWXYZabcdefghjkmnpqrstuvwxyz23456789";
  let pw="";for(const b of crypto.randomBytes(10))pw+=A[b%A.length];pw+="@7";
  await db.prepare("UPDATE users SET password_hash=? WHERE id=?").run(hash(pw,crypto.randomBytes(16).toString("hex")),u.id);
  await db.prepare("DELETE FROM sessions WHERE user_id=?").run(u.id);
  await audit(req,"RESET_PASSWORD","USER",u.username);
  res.json({username:u.username,password:pw});
}));
app.get("/api/companies",auth,roles("Super Admin"),wrap(async(req,res)=>{
  res.json(await db.prepare(`SELECT c.id,c.name,c.code,c.industry,c.address,c.contact_email,c.contact_phone,c.status,c.created_at,c.custom_domain,c.smtp_user,
    (SELECT COUNT(*) FROM employees e WHERE e.company_id=c.id AND e.status='Active') employee_count
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
    if(x.logo){try{await storeImage("company",companyId,companyId,x.logo)}catch(e){console.warn("logo skipped:",e.message)}}
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
const EMAIL_RE=/^[^\s@]+@[^\s@]+\.[^\s@]+$/;
app.put("/api/companies/:id",auth,roles("Super Admin"),wrap(async(req,res)=>{
  const x=req.body||{};
  if(!String(x.name||"").trim())return res.status(400).json({error:"Company name is required"});
  if(x.contact_email&&!EMAIL_RE.test(x.contact_email))return res.status(400).json({error:"Enter a valid contact email address"});
  const r=await db.prepare("UPDATE companies SET name=?,industry=?,address=?,contact_email=?,contact_phone=? WHERE id=?")
    .run(String(x.name).trim(),x.industry||"",x.address||"",x.contact_email||"",x.contact_phone||"",req.params.id);
  if(r.changes===0)return res.status(404).json({error:"Company not found"});
  await audit(req,"UPDATE","COMPANY",req.params.id);res.json({ok:true});
}));
app.post("/api/companies/:id/email-settings",auth,roles("Super Admin"),wrap(async(req,res)=>{
  const {smtp_user,smtp_pass}=req.body||{};
  if(smtp_user&&!EMAIL_RE.test(smtp_user))return res.status(400).json({error:"Enter a valid sender email address"});
  const r=smtp_pass
    ?await db.prepare("UPDATE companies SET smtp_user=?,smtp_pass=? WHERE id=?").run(smtp_user||null,smtp_pass,req.params.id)
    :await db.prepare("UPDATE companies SET smtp_user=? WHERE id=?").run(smtp_user||null,req.params.id);
  if(r.changes===0)return res.status(404).json({error:"Company not found"});
  await audit(req,"UPDATE","EMAIL_SETTINGS",String(req.params.id));res.json({ok:true});
}));
app.get("/api/company-profile",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  res.json(await db.prepare("SELECT name,code,industry,address,contact_email,contact_phone,custom_domain FROM companies WHERE id=?").get(req.user.company_id));
}));
app.post("/api/company-profile",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body||{};
  if(x.contact_email&&!EMAIL_RE.test(x.contact_email))return res.status(400).json({error:"Enter a valid contact email address"});
  await db.prepare("UPDATE companies SET industry=?,address=?,contact_email=?,contact_phone=? WHERE id=?")
    .run(x.industry||"",x.address||"",x.contact_email||"",x.contact_phone||"",req.user.company_id);
  await audit(req,"UPDATE","COMPANY_PROFILE","");res.json({ok:true});
}));
app.post("/api/companies/:id/domain",auth,roles("Super Admin"),wrap(async(req,res)=>{
  let d=String(req.body?.domain||"").trim().toLowerCase().replace(/^https?:\/\//,"").replace(/\/.*$/,"");
  if(d){
    if(!/^(?!-)[a-z0-9-]{1,63}(\.[a-z0-9-]{1,63})+$/.test(d))return res.status(400).json({error:"Enter a valid domain such as hr.example.com"});
    if(d.endsWith(".onrender.com"))return res.status(400).json({error:"Use the company's own domain, not the platform address"});
    const taken=await db.prepare("SELECT id FROM companies WHERE LOWER(custom_domain)=?").get(d);
    if(taken&&String(taken.id)!==String(req.params.id))return res.status(400).json({error:"This domain is already assigned to another company"});
  }
  const r=await db.prepare("UPDATE companies SET custom_domain=? WHERE id=?").run(d||null,req.params.id);
  if(r.changes===0)return res.status(404).json({error:"Company not found"});
  await audit(req,"DOMAIN","COMPANY",`${req.params.id}: ${d||"(cleared)"}`);
  res.json({ok:true,domain:d||null});
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
// When an employee is created or given a Biometric ID after the device has already sent punches,
// rebuild their attendance from the punches already stored (last 45 days).
async function backfillFromPunches(companyId,biometricId){
  const bid=String(biometricId||"").trim();
  if(!bid)return;
  try{
    const cut=new Date(Date.now()-45*86400000).toISOString().slice(0,19);
    const rows=await db.prepare("SELECT biometric_id,punch_time FROM punches WHERE company_id=? AND biometric_id=? AND punch_time>=?").all(companyId,bid,cut);
    if(rows.length)await ingestBatch(companyId,null,rows);
  }catch(e){console.error("Attendance backfill failed:",e.message)}
}
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
    await backfillFromPunches(req.user.company_id,x.biometric_id);
    await audit(req,"CREATE","EMPLOYEE",x.employee_code);res.json({id:r.lastInsertRowid});
    notifyEmployeeWelcome(req.user.company_id,x).catch(e=>console.error("welcome mail",e.message));
  }catch(e){res.status(400).json({error:friendlyDupError(e)})}
}));
app.put("/api/employees/:id",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await db.prepare(`UPDATE employees SET employee_code=?,name=?,email=?,phone=?,department=?,designation=?,manager=?,reporting_manager_id=?,branch=?,joining_date=?,status=?,biometric_id=?,salary=?,bank_name=?,bank_account=?,ifsc=?,pf_number=?,esic_number=?,uan_number=?,pan_number=?,date_of_birth=?,basic_salary=?,hra=?,other_allowances=?,pf_applicable=?,esic_applicable=? WHERE id=? AND company_id=?`)
      .run(x.employee_code,x.name,x.email,x.phone,x.department,x.designation,x.manager,x.reporting_manager_id||null,x.branch,x.joining_date,x.status||"Active",x.biometric_id||null,x.salary||0,x.bank_name,x.bank_account,x.ifsc,x.pf_number||null,x.esic_number||null,x.uan_number||null,x.pan_number||null,x.date_of_birth||null,x.basic_salary||0,x.hra||0,x.other_allowances||0,+!!x.pf_applicable,+!!x.esic_applicable,req.params.id,req.user.company_id);
    if(r.changes===0)return res.status(404).json({error:"Employee not found"});
    await backfillFromPunches(req.user.company_id,x.biometric_id);
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
  notifyAgreementStep(req.user.company_id,emp.id,"Pending Employee").catch(e=>console.error("agreement mail",e.message));
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
  if(updated.status!=="Completed")notifyAgreementStep(req.user.company_id,updated.employee_id,updated.status).catch(e=>console.error("agreement mail",e.message));
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

/* ---------------- Letters (appointment letter with PDF + email) ---------------- */
const PDFDocument=require("pdfkit");
const LETTER_PLACEHOLDERS=["employee_name","employee_code","designation","department","branch","reporting_manager","joining_date","monthly_ctc","annual_ctc","basic_monthly","hra_monthly","other_allowances_monthly","company_name","company_address","issue_date","ref_no"];
const DEFAULT_LETTER_TEMPLATE=`Ref: {{ref_no}}
Date: {{issue_date}}

To,
{{employee_name}}
Employee Code: {{employee_code}}

Subject: Letter of Appointment

Dear {{employee_name}},

We are pleased to appoint you at {{company_name}} on the terms set out below.

## Position details
Designation: {{designation}}
Department: {{department}}
Work location: {{branch}}
Reporting to: {{reporting_manager}}
Date of joining: {{joining_date}}

## Compensation
Your monthly gross compensation will be Rs. {{monthly_ctc}} (Rs. {{annual_ctc}} per annum), structured as follows:
Basic salary: Rs. {{basic_monthly}} per month
House rent allowance: Rs. {{hra_monthly}} per month
Other allowances: Rs. {{other_allowances_monthly}} per month
Statutory deductions such as Provident Fund, ESIC and income tax (TDS) will be applied as per applicable law.

## Terms of employment
1. You are expected to follow the company's policies, code of conduct and confidentiality obligations.
2. This appointment is subject to satisfactory verification of the documents and information you have provided.
3. Either party may end the employment by giving notice as per company policy.

Please sign the duplicate copy of this letter as a token of your acceptance of these terms.

We welcome you to {{company_name}} and look forward to a long and successful association.

Yours sincerely,
For {{company_name}}


Authorized Signatory`;
const fmtDate=d=>{const x=new Date(d);return isNaN(x)?String(d||""):x.toLocaleDateString("en-GB",{day:"2-digit",month:"long",year:"numeric"})};
const fmtNum=n=>Math.round(Number(n)||0).toLocaleString("en-IN");
async function getLetterTemplate(companyId){
  const c=await db.prepare("SELECT letter_template FROM companies WHERE id=?").get(companyId);
  return (c?.letter_template&&c.letter_template.trim())?c.letter_template:DEFAULT_LETTER_TEMPLATE;
}
async function buildLetter(companyId,emp,refNo){
  const company=await db.prepare("SELECT name,code,address FROM companies WHERE id=?").get(companyId);
  const missing=[];
  if(!emp.designation)missing.push("designation");
  if(!emp.joining_date)missing.push("joining date");
  const basic=Number(emp.basic_salary)||0,hra=Number(emp.hra)||0,other=Number(emp.other_allowances)||0;
  if(basic+hra+other<=0)missing.push("salary structure (Basic, HRA, Other Allowances)");
  let manager="";
  if(emp.reporting_manager_id){const m=await db.prepare("SELECT name FROM employees WHERE id=?").get(emp.reporting_manager_id);manager=m?.name||""}
  const ctc=basic+hra+other;
  const values={employee_name:emp.name,employee_code:emp.employee_code,designation:emp.designation,department:emp.department,branch:emp.branch,
    reporting_manager:manager||emp.manager,joining_date:fmtDate(emp.joining_date),monthly_ctc:fmtNum(ctc),annual_ctc:fmtNum(ctc*12),
    basic_monthly:fmtNum(basic),hra_monthly:fmtNum(hra),other_allowances_monthly:fmtNum(other),company_name:company?.name,company_address:company?.address,
    issue_date:fmtDate(new Date()),ref_no:refNo};
  const tpl=await getLetterTemplate(companyId);
  const content=tpl.replace(/\{\{\s*(\w+)\s*\}\}/g,(m,k)=>{const v=values[k];return v==null||String(v).trim()===""?"-":String(v)});
  return {content,missing,company};
}
function letterToHtml(content){
  return content.split("\n").map(line=>{
    if(line.startsWith("## "))return `<div style="font-weight:700;margin:14px 0 4px">${esc(line.slice(3))}</div>`;
    return line.trim()===""?`<div style="height:8px"></div>`:`<div>${esc(line)}</div>`;
  }).join("");
}
function buildLetterPdf(company,content){
  return new Promise((resolve,reject)=>{
    const doc=new PDFDocument({size:"A4",margins:{top:56,bottom:56,left:64,right:64}});
    const bufs=[];doc.on("data",b=>bufs.push(b));doc.on("end",()=>resolve(Buffer.concat(bufs)));doc.on("error",reject);
    doc.font("Helvetica-Bold").fontSize(17).fillColor("#312e81").text(company?.name||"");
    if(company?.address)doc.font("Helvetica").fontSize(9).fillColor("#64748b").text(company.address);
    doc.moveDown(0.6);
    const y=doc.y;doc.moveTo(64,y).lineTo(531,y).lineWidth(1.2).strokeColor("#4f46e5").stroke();
    doc.moveDown(1);
    for(const line of content.split("\n")){
      if(line.startsWith("## ")){doc.moveDown(0.4).font("Helvetica-Bold").fontSize(11).fillColor("#0f172a").text(line.slice(3));}
      else if(line.trim()===""){doc.moveDown(0.5);}
      else{doc.font("Helvetica").fontSize(10.5).fillColor("#0f172a").text(line,{lineGap:2});}
    }
    doc.end();
  });
}
async function nextRefNo(companyId,code){
  const yr=new Date().getFullYear();
  const n=Number((await db.prepare("SELECT COUNT(*) c FROM letters WHERE company_id=? AND issued_at LIKE ?").get(companyId,yr+"%")).c)+1;
  return `APT/${(code||"CO").toUpperCase()}/${yr}/${String(n).padStart(4,"0")}`;
}
async function emailLetter(companyId,emp,letter,pdf){
  if(!emp.email)return false;
  const sender=await companySender(companyId);
  const ok=await sendMail(emp.email,`${letter.letter_type} — ${sender?.name||"Company"}`,layout(letter.letter_type,
    `<p>Dear ${esc(emp.name)},</p><p>Please find attached your <b>${esc(letter.letter_type)}</b> (Ref: ${esc(letter.ref_no)}).</p>
     <p>Kindly review it, sign the copy and return it to the HR team. Reach out to HR if you have any questions.</p><p>Regards,<br>HR Team, ${esc(sender?.name||"")}</p>`),
    sender,[{filename:`${letter.ref_no.replace(/\//g,"-")}.pdf`,content:pdf,contentType:"application/pdf"}]);
  return !!ok;
}
app.get("/api/letter-template",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  res.json({template:await getLetterTemplate(req.user.company_id),defaultTemplate:DEFAULT_LETTER_TEMPLATE,placeholders:LETTER_PLACEHOLDERS});
}));
app.post("/api/letter-template",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const t=String(req.body?.template||"").trim();
  if(!t)return res.status(400).json({error:"The letter template cannot be empty"});
  if(t.length>20000)return res.status(400).json({error:"The letter template is too long"});
  const unknown=[...new Set([...t.matchAll(/\{\{\s*(\w+)\s*\}\}/g)].map(m=>m[1]).filter(k=>!LETTER_PLACEHOLDERS.includes(k)))];
  if(unknown.length)return res.status(400).json({error:"Unknown placeholder(s): "+unknown.map(u=>`{{${u}}}`).join(", ")});
  await db.prepare("UPDATE companies SET letter_template=? WHERE id=?").run(t===DEFAULT_LETTER_TEMPLATE?null:t,req.user.company_id);
  await audit(req,"UPDATE","LETTER_TEMPLATE","");res.json({ok:true});
}));
app.post("/api/employees/:id/letters/preview",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found"});
  const co=await db.prepare("SELECT code FROM companies WHERE id=?").get(req.user.company_id);
  const {content,missing}=await buildLetter(req.user.company_id,emp,await nextRefNo(req.user.company_id,co?.code));
  res.json({html:letterToHtml(content),missing,hasEmail:!!emp.email});
}));
app.post("/api/employees/:id/letters",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found"});
  const co=await db.prepare("SELECT code FROM companies WHERE id=?").get(req.user.company_id);
  const ref=await nextRefNo(req.user.company_id,co?.code);
  const {content,missing,company}=await buildLetter(req.user.company_id,emp,ref);
  if(missing.length)return res.status(400).json({error:"Please complete the employee's "+missing.join(", ")+" before issuing the letter."});
  const r=await db.prepare("INSERT INTO letters(company_id,employee_id,letter_type,ref_no,content,issued_by) VALUES(?,?,?,?,?,?)").run(req.user.company_id,emp.id,"Appointment Letter",ref,content,req.user.username);
  const letter={id:r.lastInsertRowid,letter_type:"Appointment Letter",ref_no:ref};
  let emailed=false;
  if(req.body?.send_email!==false){
    emailed=await emailLetter(req.user.company_id,emp,letter,await buildLetterPdf(company,content));
    if(emailed)await db.prepare("UPDATE letters SET emailed_at=? WHERE id=?").run(new Date().toISOString(),letter.id);
  }
  await audit(req,"ISSUE","LETTER",`${ref} → ${emp.employee_code}`);
  res.json({ok:true,id:letter.id,ref_no:ref,emailed,hasEmail:!!emp.email});
}));
app.get("/api/letters",auth,requireCompany,wrap(async(req,res)=>{
  let q="SELECT l.id,l.employee_id,l.letter_type,l.ref_no,l.issued_by,l.issued_at,l.emailed_at,e.employee_code,e.name,e.email FROM letters l JOIN employees e ON e.id=l.employee_id WHERE l.company_id=?";const p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND l.employee_id=?";p.push(req.user.employee_id)}
  else if(!["Super Admin","HR Admin","Director"].includes(req.user.role))return res.status(403).json({error:"Permission denied"});
  res.json(await db.prepare(q+" ORDER BY l.id DESC").all(...p));
}));
app.get("/api/letters/:id/pdf",auth,requireCompany,wrap(async(req,res)=>{
  const l=await db.prepare("SELECT * FROM letters WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!l)return res.status(404).json({error:"Letter not found"});
  if(req.user.role==="Employee"&&l.employee_id!==req.user.employee_id)return res.status(403).json({error:"Permission denied"});
  if(!["Employee","Super Admin","HR Admin","Director"].includes(req.user.role))return res.status(403).json({error:"Permission denied"});
  const company=await db.prepare("SELECT name,address FROM companies WHERE id=?").get(req.user.company_id);
  const pdf=await buildLetterPdf(company,l.content);
  res.setHeader("Content-Type","application/pdf");
  res.setHeader("Content-Disposition",`attachment; filename="${l.ref_no.replace(/\//g,"-")}.pdf"`);
  res.send(pdf);
}));
app.post("/api/letters/:id/resend",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const l=await db.prepare("SELECT * FROM letters WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!l)return res.status(404).json({error:"Letter not found"});
  const emp=await db.prepare("SELECT * FROM employees WHERE id=?").get(l.employee_id);
  if(!emp?.email)return res.status(400).json({error:"This employee has no email address on file"});
  const company=await db.prepare("SELECT name,address FROM companies WHERE id=?").get(req.user.company_id);
  const ok=await emailLetter(req.user.company_id,emp,l,await buildLetterPdf(company,l.content));
  if(!ok)return res.status(502).json({error:"The email could not be sent. Check the company's email settings."});
  await db.prepare("UPDATE letters SET emailed_at=? WHERE id=?").run(new Date().toISOString(),l.id);
  await audit(req,"RESEND","LETTER",l.ref_no);res.json({ok:true});
}));

app.get("/api/departments",auth,requireCompany,wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM departments WHERE company_id=? ORDER BY name").all(req.user.company_id))));
app.post("/api/departments",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  try{await db.prepare("INSERT INTO departments(company_id,name) VALUES(?,?)").run(req.user.company_id,req.body.name);res.json({ok:true})}
  catch(e){res.status(400).json({error:"Department already exists"})}
}));

/* ---------------- Work timing (company default + per-employee override) ---------------- */
const DEFAULT_TIMING={start:"09:30",end:"18:30",grace:15,break_on:true,break_start:"13:30",break_end:"14:00",min_hours:8,notify_low_hours:true,notify_late:true,web_clock:true,require_location:false,allow_remote:true,office_lat:null,office_lng:null,office_radius:200};
const distM=(a,b,c,d)=>{const R=6371000,r=x=>x*Math.PI/180,dl=r(c-a),dn=r(d-b),h=Math.sin(dl/2)**2+Math.cos(r(a))*Math.cos(r(c))*Math.sin(dn/2)**2;return 2*R*Math.asin(Math.sqrt(h))};
const HHMM=/^([01]\d|2[0-3]):[0-5]\d$/;
function cleanTiming(x,partial){
  const o={};
  for(const k of ["start","end","break_start","break_end"])if(x?.[k]!==undefined&&x[k]!==""){if(!HHMM.test(x[k]))throw new Error("Time must be in HH:MM format");o[k]=x[k]}
  if(x?.grace!==undefined&&x.grace!==""){const g=Number(x.grace);if(!(g>=0&&g<=120))throw new Error("Grace minutes must be between 0 and 120");o.grace=g}
  if(x?.min_hours!==undefined&&x.min_hours!==""){const h=Number(x.min_hours);if(!(h>=0&&h<=16))throw new Error("Minimum hours must be between 0 and 16");o.min_hours=h}
  for(const k of ["office_lat","office_lng"])if(x?.[k]!==undefined){
    if(x[k]===""||x[k]===null)o[k]=null;
    else{const v=Number(x[k]);if(!Number.isFinite(v)||Math.abs(v)>(k==="office_lat"?90:180))throw new Error("Office location is not valid");o[k]=v}
  }
  if(x?.office_radius!==undefined&&x.office_radius!==""){const v=Number(x.office_radius);if(!(v>=20&&v<=5000))throw new Error("Office radius must be between 20 and 5000 metres");o.office_radius=v}
  for(const k of ["break_on","notify_low_hours","notify_late","web_clock","require_location","allow_remote"])if(x?.[k]!==undefined)o[k]=!!x[k];
  if(!partial&&(!o.start||!o.end))throw new Error("Start and close time are required");
  return o;
}
function parseJSON(t){try{return JSON.parse(t||"")||{}}catch{return {}}}
function timingFor(company,emp){return {...DEFAULT_TIMING,...parseJSON(company?.work_timing),...parseJSON(emp?.work_timing)}}
const toMin=t=>{const [h,m]=String(t).split(":").map(Number);return h*60+m};
function dayMetrics(t,firstIn,lastOut){
  const r={late_minutes:0,worked_minutes:0,overtime_minutes:0};
  if(!firstIn||firstIn.length<16)return r;
  const inM=toMin(firstIn.slice(11,16));
  r.late_minutes=Math.max(0,inM-(toMin(t.start)+Number(t.grace||0)));
  if(lastOut&&lastOut.length>=16){
    let w=toMin(lastOut.slice(11,16))-inM;
    const bLen=t.break_on?Math.max(0,toMin(t.break_end)-toMin(t.break_start)):0;
    if(t.break_on&&inM<toMin(t.break_start)&&toMin(lastOut.slice(11,16))>toMin(t.break_end))w-=bLen;
    r.worked_minutes=Math.max(0,w);
    const std=toMin(t.end)-toMin(t.start)-bLen;
    r.overtime_minutes=Math.max(0,w-std);
  }
  return r;
}
/* ---------------- Email notifications: leave, late arrival, minimum hours ---------------- */
function esc2(v){return String(v??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]))}
async function seniorEmails(companyId){
  const rows=await db.prepare("SELECT e.email FROM users u JOIN employees e ON e.id=u.employee_id WHERE u.company_id=? AND u.active=1 AND u.role IN ('HR Admin','Director') AND e.email IS NOT NULL AND e.email<>''").all(companyId);
  const c=await db.prepare("SELECT contact_email FROM companies WHERE id=?").get(companyId);
  const direct=await db.prepare("SELECT email FROM users WHERE company_id=? AND active=1 AND role IN ('HR Admin','Director') AND email IS NOT NULL AND email<>''").all(companyId);
  return [...new Set([...rows.map(r=>r.email),...direct.map(r=>r.email),c?.contact_email].filter(Boolean))];
}
async function notifyEmployeeWelcome(companyId,x){
  if(!x.email)return;
  const co=await db.prepare("SELECT name FROM companies WHERE id=?").get(companyId);
  const sender=await companySender(companyId);
  const row=(k,v)=>v?`<tr><td style="padding:4px 12px 4px 0;color:#64748b">${k}</td><td>${esc2(v)}</td></tr>`:"";
  await sendMail(x.email,`Welcome to ${co?.name||"the team"}`,layout(`Welcome to ${esc2(co?.name||"the team")}`,
    `<p>Hi ${esc2(x.name)},</p><p>Welcome aboard! Your employee profile has been created in the HR portal.</p>
     <table style="border-collapse:collapse;font-size:14px">${row("Employee code",x.employee_code)}${row("Designation",x.designation)}${row("Department",x.department)}${row("Joining date",x.joining_date)}</table>
     <p style="margin-top:14px">Your HR team will share your portal login separately. Once you sign in you can view attendance, apply for leave, download payslips and sign your onboarding agreement.</p>`),sender);
}
async function notifyAgreementStep(companyId,empId,status){
  const emp=await db.prepare("SELECT name,email FROM employees WHERE id=?").get(empId);
  if(!emp)return;
  const sender=await companySender(companyId);
  if(status==="Pending Employee"){
    if(emp.email)await sendMail(emp.email,"Please sign your onboarding agreement",layout("Your agreement is ready",`<p>Hi ${esc2(emp.name)},</p><p>Your onboarding agreement is ready. Please sign in to the HR portal, open <b>Agreements</b> and sign it.</p>`),sender);
    return;
  }
  const label=status==="Pending HR"?"HR":"Director";
  const rows=status==="Pending HR"
    ?await db.prepare("SELECT e.email FROM users u JOIN employees e ON e.id=u.employee_id WHERE u.company_id=? AND u.active=1 AND u.role='HR Admin' AND e.email<>''").all(companyId)
    :await db.prepare("SELECT e.email FROM users u JOIN employees e ON e.id=u.employee_id WHERE u.company_id=? AND u.active=1 AND u.role='Director' AND e.email<>''").all(companyId);
  const to=[...new Set([...rows.map(r=>r.email),...(await db.prepare("SELECT email FROM users WHERE company_id=? AND active=1 AND role=? AND email IS NOT NULL AND email<>''").all(companyId,status==="Pending HR"?"HR Admin":"Director")).map(r=>r.email)])];
  for(const t of to)await sendMail(t,`Agreement awaiting your signature — ${emp.name}`,layout("Agreement awaiting your signature",`<p>The onboarding agreement of <b>${esc2(emp.name)}</b> has been signed by the previous party and now needs the <b>${label}</b> signature.</p><p>Please sign in to the HR portal and open <b>Agreements</b>.</p>`),sender);
}
async function notifyLeaveApplied(companyId,empId,x){
  const emp=await db.prepare("SELECT name,employee_code,reporting_manager_id FROM employees WHERE id=?").get(empId);
  if(!emp)return;
  const mgr=emp.reporting_manager_id?await db.prepare("SELECT email FROM employees WHERE id=?").get(emp.reporting_manager_id):null;
  const to=[...new Set([mgr?.email,...await seniorEmails(companyId)].filter(Boolean))];
  if(!to.length)return;
  const sender=await companySender(companyId);
  const label=x.category==="WFH"?"work from home":x.category==="Permission"?"permission":"leave";
  const row=(k,v)=>`<tr><td style="padding:4px 12px 4px 0;color:#64748b">${k}</td><td>${esc2(v)}</td></tr>`;
  const html=layout(`New ${label} request`,`<p><b>${esc2(emp.name)}</b> (${esc2(emp.employee_code)}) has applied for ${label}.</p>
   <table style="border-collapse:collapse;font-size:14px">${row("Type",x.leave_type||x.category)}${row("From",x.from_date)}${row("To",x.to_date)}${row("Days",x.days||1)}${row("Reason",x.reason||"-")}</table>
   <p style="margin-top:14px">Please review it in the HR portal.</p>`);
  for(const t of to)await sendMail(t,`${emp.name} applied for ${label}`,html,sender);
}
async function notifyLeaveDecision(companyId,leaveId,status,by){
  const l=await db.prepare("SELECT l.*,e.email,e.name FROM leave_requests l JOIN employees e ON e.id=l.employee_id WHERE l.id=? AND l.company_id=?").get(leaveId,companyId);
  if(!l?.email)return;
  const sender=await companySender(companyId);
  await sendMail(l.email,`Your request was ${String(status).toLowerCase()}`,layout(`Request ${status}`,
    `<p>Hi ${esc2(l.name)}, your ${esc2(l.leave_type||l.category)} request (${esc2(l.from_date)} to ${esc2(l.to_date)}) has been <b>${esc2(status)}</b> by ${esc2(by)}.</p>`),sender);
}
const fmtHM=m=>Math.floor(m/60)+"h "+String(m%60).padStart(2,"0")+"m";
function istNow(){
  const p=Object.fromEntries(new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Kolkata",year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",hour12:false}).formatToParts(new Date()).map(x=>[x.type,x.value]));
  return {date:`${p.year}-${p.month}-${p.day}`,minutes:(+p.hour%24)*60+ +p.minute};
}
// Builds the day's late / short-hours lists and emails them to HR and directors.
async function sendDailyDigest(companyId,date,force){
  const co=await db.prepare("SELECT id,name,work_timing,smtp_user,smtp_pass FROM companies WHERE id=?").get(companyId);
  const base=timingFor(co,null);
  if(!force&&!base.notify_low_hours&&!base.notify_late)return {sent:0,reason:"Notifications are off"};
  const rows=await db.prepare("SELECT a.first_in,a.last_out,e.name,e.employee_code,e.work_timing FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=? AND a.work_date=? AND e.status='Active'").all(companyId,date);
  const late=[],low=[];
  for(const r of rows){
    const t=timingFor(co,r),m=dayMetrics(t,r.first_in,r.last_out);
    if((force||base.notify_late)&&m.late_minutes>0)late.push({...r,m});
    if((force||base.notify_low_hours)&&Number(t.min_hours)>0){
      const need=Math.round(t.min_hours*60);
      if(!r.last_out)low.push({...r,m,note:"No check-out recorded",need});
      else if(m.worked_minutes<need)low.push({...r,m,note:fmtHM(m.worked_minutes),need});
    }
  }
  if(!late.length&&!low.length)return {sent:0,late:0,low:0,reason:"Nothing to report for today"};
  const tbl=(head,items)=>`<table style="border-collapse:collapse;width:100%;font-size:13px;margin:8px 0 18px"><tr>${head.map(h=>`<th style="text-align:left;padding:6px;border:1px solid #e5e7eb;background:#f8fafc">${h}</th>`).join("")}</tr>${items.map(r=>`<tr>${r.map(c=>`<td style="padding:6px;border:1px solid #e5e7eb">${c}</td>`).join("")}</tr>`).join("")}</table>`;
  const html=layout(`Attendance summary — ${date}`,
    (late.length?`<h3 style="margin:10px 0 0">Late arrivals (${late.length})</h3>`+tbl(["Employee","First in","Late by"],late.map(r=>[esc2(r.employee_code+" - "+r.name),esc2((r.first_in||"").slice(11,16)),r.m.late_minutes+" min"])):"")+
    (low.length?`<h3 style="margin:10px 0 0">Minimum working hours not completed (${low.length})</h3>`+tbl(["Employee","Worked","Required"],low.map(r=>[esc2(r.employee_code+" - "+r.name),esc2(r.note),fmtHM(r.need)])):""));
  const to=await seniorEmails(companyId),sender={name:co.name,smtp_user:co.smtp_user,smtp_pass:co.smtp_pass};
  for(const t of to)await sendMail(t,`Attendance summary ${date} — ${co.name}`,html,sender);
  return {sent:to.length,late:late.length,low:low.length};
}
app.post("/api/work-timing/send-digest",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  res.json(await sendDailyDigest(req.user.company_id,istNow().date,true));
}));
// Once a day, 30 minutes after office close (IST), per company.
async function digestTick(){
  const now=istNow();
  const cos=await db.prepare("SELECT id,work_timing,last_digest FROM companies").all();
  for(const c of cos){
    if(c.last_digest===now.date||new Date(now.date+"T00:00:00Z").getUTCDay()===0)continue;
    if(now.minutes<toMin(timingFor(c,null).end)+30)continue;
    const r=await db.prepare("UPDATE companies SET last_digest=? WHERE id=? AND (last_digest IS NULL OR last_digest<>?)").run(now.date,c.id,now.date);
    if(r.changes)await sendDailyDigest(c.id,now.date,false).catch(e=>console.error("digest",e.message));
  }
}
setInterval(()=>digestTick().catch(e=>console.error("digest tick",e.message)),10*60*1000);

app.get("/api/work-timing",auth,requireCompany,wrap(async(req,res)=>{
  const c=await db.prepare("SELECT work_timing FROM companies WHERE id=?").get(req.user.company_id);
  res.json(timingFor(c,null));
}));
app.post("/api/work-timing",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  let o;try{o=cleanTiming(req.body,false)}catch(e){return res.status(400).json({error:e.message})}
  if(toMin(o.end)<=toMin(o.start))return res.status(400).json({error:"Close time must be after start time"});
  const full={...DEFAULT_TIMING,...o};
  if(full.break_on&&toMin(full.break_end)<=toMin(full.break_start))return res.status(400).json({error:"Break end must be after break start"});
  await db.prepare("UPDATE companies SET work_timing=? WHERE id=?").run(JSON.stringify(full),req.user.company_id);
  await audit(req,"UPDATE","WORK_TIMING",String(req.user.company_id));res.json({ok:true});
}));
// Per-employee override: body {use_default:true} clears it; otherwise partial fields are stored.
app.post("/api/employees/:id/work-timing",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  let v=null;
  if(!req.body.use_default){
    try{v=JSON.stringify(cleanTiming(req.body,true))}catch(e){return res.status(400).json({error:e.message})}
  }
  const r=await db.prepare("UPDATE employees SET work_timing=? WHERE id=? AND company_id=?").run(v,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Employee not found"});
  await audit(req,"UPDATE","EMP_TIMING",String(req.params.id));res.json({ok:true});
}));
app.get("/api/employees/:id/work-timing",auth,requireCompany,roles("Super Admin","HR Admin","Director"),wrap(async(req,res)=>{
  const e=await db.prepare("SELECT work_timing FROM employees WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!e)return res.status(404).json({error:"Employee not found"});
  const c=await db.prepare("SELECT work_timing FROM companies WHERE id=?").get(req.user.company_id);
  res.json({custom:!!e.work_timing,effective:timingFor(c,e),override:parseJSON(e.work_timing)});
}));

/* ---------------- Web clock in / clock out (for employees without a biometric punch) ---------------- */
function istStamp(){
  const p=Object.fromEntries(new Intl.DateTimeFormat("en-CA",{timeZone:"Asia/Kolkata",year:"numeric",month:"2-digit",day:"2-digit",hour:"2-digit",minute:"2-digit",second:"2-digit",hour12:false}).formatToParts(new Date()).map(x=>[x.type,x.value]));
  return `${p.year}-${p.month}-${p.day}T${String(+p.hour%24).padStart(2,"0")}:${p.minute}:${p.second}`;
}
app.get("/api/attendance/clock",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.employee_id;
  const co=await db.prepare("SELECT work_timing FROM companies WHERE id=?").get(req.user.company_id);
  const enabled=!!timingFor(co,null).web_clock;
  if(!eid)return res.json({enabled:false,linked:false});
  const a=await db.prepare("SELECT first_in,last_out FROM attendance WHERE employee_id=? AND work_date=?").get(eid,istStamp().slice(0,10));
  const t=timingFor(co,null);
  res.json({enabled,linked:true,first_in:a?.first_in||null,last_out:a?.last_out||null,require_location:!!t.require_location});
}));
app.post("/api/attendance/clock",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.employee_id;
  if(!eid)return res.status(400).json({error:"Your login is not linked to an employee record"});
  const co=await db.prepare("SELECT work_timing FROM companies WHERE id=?").get(req.user.company_id);
  if(!timingFor(co,null).web_clock)return res.status(403).json({error:"Web clock in/out is turned off for your company"});
  const now=istStamp(),d=now.slice(0,10);
  const t=timingFor(co,null);
  const lat=Number(req.body.lat),lng=Number(req.body.lng),hasLoc=req.body.lat!=null&&req.body.lng!=null&&Number.isFinite(lat)&&Number.isFinite(lng)&&Math.abs(lat)<=90&&Math.abs(lng)<=180;
  if(t.require_location&&!hasLoc)return res.status(400).json({error:"Location access is required to clock in or out. Please allow location in your browser."});
  let loc=null;
  if(hasLoc){
    loc={lat:+lat.toFixed(6),lng:+lng.toFixed(6),acc:Math.round(Number(req.body.acc)||0)};
    if(t.office_lat!=null&&t.office_lng!=null){
      const dist=Math.round(distM(lat,lng,t.office_lat,t.office_lng));
      loc.dist=dist;loc.mode=dist<=Number(t.office_radius||200)?"Office":"Remote";
      if(loc.mode==="Remote"&&!t.allow_remote)return res.status(403).json({error:`You are ${dist} m away from the office. Clocking in from outside the office is not allowed.`});
    }else loc.mode="Recorded";
  }
  const locJson=loc?JSON.stringify(loc):null;
  const a=await db.prepare("SELECT id,first_in,last_out FROM attendance WHERE employee_id=? AND work_date=?").get(eid,d);
  if(req.body.action==="in"){
    if(a?.first_in)return res.status(400).json({error:"You have already clocked in today"});
    if(a)await db.prepare("UPDATE attendance SET first_in=?,status='Present',in_loc=? WHERE id=?").run(now,locJson,a.id);
    else await db.prepare("INSERT INTO attendance(company_id,employee_id,work_date,first_in,status,source,in_loc) VALUES(?,?,?,?,?,?,?)").run(req.user.company_id,eid,d,now,"Present","Web",locJson);
    return res.json({ok:true,time:now,mode:loc?.mode||null});
  }
  if(req.body.action==="out"){
    if(!a?.first_in)return res.status(400).json({error:"Clock in first"});
    await db.prepare("UPDATE attendance SET last_out=?,out_loc=? WHERE id=?").run(now,locJson,a.id);
    return res.json({ok:true,time:now,mode:loc?.mode||null});
  }
  res.status(400).json({error:"Invalid action"});
}));

app.get("/api/attendance",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT a.*,e.employee_code,e.name,e.department,e.work_timing AS emp_timing FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=?`;
  let params=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND e.id=?";params.push(req.user.employee_id)}
  else if(req.user.role==="Manager" && req.user.employee_id){q+=" AND (e.reporting_manager_id=? OR e.id=?)";params.push(req.user.employee_id,req.user.employee_id)}
  q+=" ORDER BY a.work_date DESC,a.first_in DESC";
  const co=await db.prepare("SELECT work_timing FROM companies WHERE id=?").get(req.user.company_id);
  const rows=await db.prepare(q).all(...params);
  res.json(rows.map(({emp_timing,...a})=>{
    if(a.source==="Manual"&&(a.late_minutes||a.overtime_minutes))return a;
    const t=timingFor(co,{work_timing:emp_timing});
    return {...a,...dayMetrics(t,a.first_in,a.last_out),min_minutes:Math.round(Number(t.min_hours||0)*60)};
  }));
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
  notifyLeaveApplied(req.user.company_id,eid,{category,...req.body}).catch(e=>console.error("leave notify",e.message));
}));
app.post("/api/leaves/:id/status",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>{
  const r=await db.prepare("UPDATE leave_requests SET status=?,approved_by=? WHERE id=? AND company_id=?").run(req.body.status,req.user.username,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Leave request not found"});
  await audit(req,req.body.status,"LEAVE",req.params.id);res.json({ok:true});
  notifyLeaveDecision(req.user.company_id,req.params.id,req.body.status,req.user.username).catch(e=>console.error("leave notify",e.message));
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
async function serialTaken(sn,exceptId){
  if(!sn)return false;
  const r=await db.prepare("SELECT id FROM biometric_devices WHERE UPPER(serial_no)=UPPER(?)").get(sn);
  return !!r&&String(r.id)!==String(exceptId||"");
}
app.post("/api/biometric/devices",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;const apiKey=crypto.randomBytes(24).toString("hex");
  const mode=x.mode==="adms"?"adms":"lan";
  const serial=String(x.serial_no||"").trim();
  if(mode==="adms"&&!serial)return res.status(400).json({error:"The device serial number is required for ADMS devices"});
  if(await serialTaken(serial))return res.status(400).json({error:"This serial number is already registered"});
  const r=await db.prepare("INSERT INTO biometric_devices(company_id,name,model,serial_no,branch,ip,port,protocol,api_key) VALUES(?,?,?,?,?,?,?,?,?)").run(req.user.company_id,x.name,x.model,serial,x.branch,x.ip||"",Number(x.port)||4370,mode==="adms"?"ADMS (push)":"ZKTeco/eSSL (LAN)",apiKey);res.json({id:r.lastInsertRowid,api_key:apiKey});
}));
app.post("/api/biometric/devices/:id/rotate-key",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const apiKey=crypto.randomBytes(24).toString("hex");
  const r=await db.prepare("UPDATE biometric_devices SET api_key=? WHERE id=? AND company_id=?").run(apiKey,req.params.id,req.user.company_id);
  if(r.changes===0)return res.status(404).json({error:"Device not found"});
  res.json({ok:true,api_key:apiKey});
}));
app.put("/api/biometric/devices/:id",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  const serial=String(x.serial_no||"").trim();
  if(x.mode==="adms"&&!serial)return res.status(400).json({error:"The device serial number is required for ADMS devices"});
  if(await serialTaken(serial,req.params.id))return res.status(400).json({error:"This serial number is already registered"});
  const protocol=x.mode==="adms"?"ADMS (push)":x.mode==="lan"?"ZKTeco/eSSL (LAN)":null;
  const r=await db.prepare("UPDATE biometric_devices SET name=?,model=?,serial_no=?,branch=?,ip=?,port=?,protocol=COALESCE(?,protocol) WHERE id=? AND company_id=?")
    .run(x.name,x.model,serial,x.branch,x.ip||"",Number(x.port)||4370,protocol,req.params.id,req.user.company_id);
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

// Bulk ingestion: stores raw punches and rebuilds first-in / last-out per employee per day
// in a handful of queries, so thousands of punches (backfill) process in seconds.
async function ingestBatch(companyId,deviceId,records){
  const seen=new Set(),rows=[];
  for(const r of records||[]){
    const b=String(r.biometric_id||"").trim(),t=String(r.punch_time||"").trim().replace(" ","T");
    if(!b||!/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}(:\d{2})?$/.test(t))continue;
    const k=b+"|"+t;if(seen.has(k))continue;seen.add(k);rows.push({b,t});
  }
  if(!rows.length)return {received:0,matched:0,unmatched:0,newPunches:0};
  let newPunches=0;
  for(let i=0;i<rows.length;i+=1000){
    const c=rows.slice(i,i+1000);
    const r=await db.prepare("INSERT INTO punches(company_id,biometric_id,punch_time,punch_type,device_id,raw_payload) SELECT ?::int,x.b,x.t,'AUTO',?::int,'{}' FROM unnest(?::text[],?::text[]) AS x(b,t) ON CONFLICT DO NOTHING")
      .run(companyId,deviceId||null,c.map(x=>x.b),c.map(x=>x.t));
    newPunches+=r.changes||0;
  }
  const bids=[...new Set(rows.map(r=>r.b))];
  const emps=await db.prepare("SELECT id,biometric_id FROM employees WHERE company_id=? AND biometric_id=ANY(?::text[])").all(companyId,bids);
  const empMap=new Map(emps.map(e=>[e.biometric_id,e.id]));
  const groups=new Map();let matched=0,unmatched=0;
  for(const r of rows){
    const eid=empMap.get(r.b);
    if(!eid){unmatched++;continue}
    matched++;
    const d=r.t.slice(0,10),k=eid+"|"+d,g=groups.get(k)||{emp:eid,d,min:r.t,max:r.t};
    if(r.t<g.min)g.min=r.t;if(r.t>g.max)g.max=r.t;groups.set(k,g);
  }
  if(groups.size){
    const gl=[...groups.values()];
    const ex=await db.prepare("SELECT employee_id,work_date,first_in,last_out FROM attendance WHERE employee_id=ANY(?::int[]) AND work_date=ANY(?::text[])")
      .all([...new Set(gl.map(g=>g.emp))],[...new Set(gl.map(g=>g.d))]);
    const exMap=new Map(ex.map(a=>[a.employee_id+"|"+a.work_date,a]));
    const up=gl.map(g=>{
      const a=exMap.get(g.emp+"|"+g.d);let first=g.min,last=g.max;
      if(a){if(a.first_in&&a.first_in<first)first=a.first_in;if(a.last_out&&a.last_out>last)last=a.last_out}
      return {e:g.emp,d:g.d,f:first,l:last===first?null:last};
    });
    for(let i=0;i<up.length;i+=500){
      const c=up.slice(i,i+500);
      await db.prepare("INSERT INTO attendance(company_id,employee_id,work_date,first_in,last_out,status,source) SELECT ?::int,x.e,x.d,x.f,x.l,'Present','eSSL' FROM unnest(?::int[],?::text[],?::text[],?::text[]) AS x(e,d,f,l) ON CONFLICT(employee_id,work_date) DO UPDATE SET first_in=excluded.first_in,last_out=excluded.last_out,source='eSSL'")
        .run(companyId,c.map(x=>x.e),c.map(x=>x.d),c.map(x=>x.f),c.map(x=>x.l));
    }
  }
  return {received:rows.length,matched,unmatched,newPunches};
}

/* ---------------- ADMS / iClock push receiver (device pushes attendance to the portal) ---------------- */
app.use("/iclock",express.text({type:()=>true,limit:"10mb"}));
const ADMS_SEEN=new Map(),ADMS_UNKNOWN=new Map();
const ADMS_MAX_AGE_DAYS=45;
async function admsDevice(req,res){
  const sn=String(req.query.SN||req.query.sn||"").trim();
  if(!sn){res.status(400).type("text/plain").send("ERROR: SN required");return null}
  const d=await db.prepare("SELECT * FROM biometric_devices WHERE UPPER(serial_no)=UPPER(?)").get(sn);
  if(!d){
    ADMS_UNKNOWN.set(sn,{serial_no:sn,ip:req.ip,last_contact:new Date().toISOString(),model:String(req.query.DeviceType||req.query.PushVersion||"")});
    res.status(403).type("text/plain").send("ERROR: device not registered");return null;
  }
  const now=Date.now();
  if(!ADMS_SEEN.has(d.id)||now-ADMS_SEEN.get(d.id)>60000){
    ADMS_SEEN.set(d.id,now);
    await db.prepare("UPDATE biometric_devices SET last_seen=?,status='Connected' WHERE id=?").run(new Date(now).toISOString(),d.id);
  }
  return d;
}
function parseAttLog(body){
  const out=[];
  for(const line of String(body||"").split(/\r?\n/)){
    const m=line.match(/^\s*(\S+)\s+(\d{4}-\d{2}-\d{2})[ T](\d{2}:\d{2}:\d{2})/);
    if(m)out.push({biometric_id:m[1],punch_time:`${m[2]}T${m[3]}`});
  }
  return out;
}
app.get(["/iclock/cdata","/iclock/cdata.aspx"],wrap(async(req,res)=>{
  const d=await admsDevice(req,res);if(!d)return;
  res.type("text/plain").send([
    `GET OPTION FROM: ${req.query.SN}`,`ATTLOGStamp=${d.adms_stamp||"None"}`,"OPERLOGStamp=9999","ATTPHOTOStamp=None",
    "ErrorDelay=30","Delay=10","TransTimes=00:00;14:05","TransInterval=1","TransFlag=TransData AttLog","Realtime=1","Encrypt=None"
  ].join("\n")+"\n");
}));
app.post(["/iclock/cdata","/iclock/cdata.aspx"],wrap(async(req,res)=>{
  const d=await admsDevice(req,res);if(!d)return;
  if(String(req.query.table||"").toUpperCase()==="ATTLOG"){
    const parsed=parseAttLog(req.body);
    const cut=new Date(Date.now()-ADMS_MAX_AGE_DAYS*86400000).toISOString().slice(0,19);
    const r=await ingestBatch(d.company_id,d.id,parsed.filter(p=>p.punch_time>=cut));
    const now=new Date().toISOString();
    await db.prepare("UPDATE biometric_devices SET last_sync=?,last_seen=?,last_error=NULL,adms_stamp=COALESCE(?,adms_stamp) WHERE id=?").run(now,now,req.query.Stamp?String(req.query.Stamp):null,d.id);
    if(parsed.length>20)await db.prepare("INSERT INTO audit_logs(company_id,user_id,action,module,details) VALUES(?,?,?,?,?)")
      .run(d.company_id,null,"SYNC","BIOMETRIC",`${d.name} (ADMS push): ${parsed.length} logs received, ${r.matched} matched, ${r.unmatched} unmatched biometric IDs`);
  }
  res.type("text/plain").send("OK");
}));
app.all(["/iclock/getrequest","/iclock/getrequest.aspx","/iclock/devicecmd","/iclock/devicecmd.aspx","/iclock/ping"],wrap(async(req,res)=>{
  const d=await admsDevice(req,res);if(!d)return;
  res.type("text/plain").send("OK");
}));
app.get("/api/biometric/unregistered",auth,roles("Super Admin"),(req,res)=>{
  const dayAgo=Date.now()-86400000;
  res.json([...ADMS_UNKNOWN.values()].filter(x=>new Date(x.last_contact).getTime()>dayAgo).sort((a,b)=>b.last_contact.localeCompare(a.last_contact)));
});

app.post("/api/biometric/punch",auth,requireCompany,wrap(async(req,res)=>{
  const x=req.body;if(!x.biometric_id||!x.punch_time)return res.status(400).json({error:"biometric_id and punch_time required"});
  const code=await ingestPunch(req.user.company_id,x.device_id||null,x.biometric_id,x.punch_time);
  res.json({ok:true,employee:code});
}));

app.post("/api/biometric/devices/:id/sync",auth,requireCompany,roles("Super Admin","HR Admin"),async(req,res)=>{
  const device=await db.prepare("SELECT * FROM biometric_devices WHERE id=? AND company_id=?").get(req.params.id,req.user.company_id);
  if(!device)return res.status(404).json({error:"Device not found"});
  if(device.protocol==="ADMS (push)")return res.status(400).json({error:"This device pushes attendance to the portal automatically (ADMS). Manual sync is not needed."});
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
    const {matched,unmatched}=await ingestBatch(req.user.company_id,device.id,logs.map(l=>({biometric_id:String(l.deviceUserId||"").trim(),punch_time:localISO(l.recordTime)})));
    try{await zk.disconnect()}catch{}
    await db.prepare("UPDATE biometric_devices SET status='Connected',last_sync=?,last_error=NULL WHERE id=?").run(new Date().toISOString(),device.id);
    await audit(req,"SYNC","BIOMETRIC",`${device.name}: ${allLogs.length} logs on device, last ${days} day(s) = ${logs.length} logs, ${matched} matched, ${unmatched} unmatched biometric IDs`);
    res.json({ok:true,totalOnDevice:allLogs.length,totalLogs:logs.length,matched,unmatched,days});
  }catch(e){
    try{await zk?.disconnect()}catch{}
    const msg=e?.err?.code?`${e.err.code} (${e.command||"connection"} to ${e.ip||device.ip})`:(e?.message||e?.err?.message||String(e));
    await db.prepare("UPDATE biometric_devices SET last_error=? WHERE id=?").run(msg,device.id);
    res.status(502).json({error:"Could not connect to the device directly from the server ("+msg+"). If the device is on another network, this is expected: use the Sync Agent installed at that office, which uploads attendance automatically."});
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
  const {matched,unmatched}=await ingestBatch(device.company_id,device.id,records);
  await db.prepare("UPDATE biometric_devices SET status='Connected',last_sync=?,last_seen=?,last_error=NULL WHERE id=?").run(new Date().toISOString(),new Date().toISOString(),device.id);
  if(records.length)await db.prepare("INSERT INTO audit_logs(company_id,user_id,action,module,details) VALUES(?,?,?,?,?)")
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
