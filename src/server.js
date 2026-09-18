const express=require("express");
const path=require("path");
const crypto=require("crypto");
const db=require("./db");
const ZKLib=require("node-zklib");
const {sendMail,layout}=require("./mail");

const app=express();
const PORT=process.env.PORT||3000;
app.set("trust proxy",1);
app.use(express.json({limit:"5mb"}));
app.use(express.urlencoded({extended:true}));
app.use(express.static(path.join(__dirname,"..","public")));

async function initSchema(){
await db.exec(`
CREATE TABLE IF NOT EXISTS companies(
 id SERIAL PRIMARY KEY, name TEXT NOT NULL, code TEXT UNIQUE NOT NULL, industry TEXT,
 address TEXT, contact_email TEXT, contact_phone TEXT, status TEXT DEFAULT 'Active', created_at TEXT DEFAULT CURRENT_TIMESTAMP,
 smtp_user TEXT, smtp_pass TEXT
);
CREATE TABLE IF NOT EXISTS users(
 id SERIAL PRIMARY KEY, company_id INTEGER, username TEXT UNIQUE, password_hash TEXT,
 role TEXT, employee_id INTEGER, active INTEGER DEFAULT 1, created_at TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE TABLE IF NOT EXISTS sessions(token TEXT PRIMARY KEY,user_id INTEGER,expires_at BIGINT,active_company_id INTEGER);
CREATE TABLE IF NOT EXISTS employees(
 id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL, employee_code TEXT, name TEXT, email TEXT, phone TEXT,
 department TEXT, designation TEXT, manager TEXT, branch TEXT, joining_date TEXT, status TEXT DEFAULT 'Active',
 biometric_id TEXT, salary REAL DEFAULT 0, bank_name TEXT, bank_account TEXT, ifsc TEXT,
 created_at TEXT DEFAULT CURRENT_TIMESTAMP, UNIQUE(company_id,employee_code)
);
CREATE TABLE IF NOT EXISTS departments(id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT,UNIQUE(company_id,name));
CREATE TABLE IF NOT EXISTS leave_types(id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,name TEXT,annual_balance REAL DEFAULT 0,UNIQUE(company_id,name));
CREATE TABLE IF NOT EXISTS leave_requests(
 id SERIAL PRIMARY KEY, company_id INTEGER NOT NULL, employee_id INTEGER, leave_type TEXT, from_date TEXT, to_date TEXT,
 days REAL, reason TEXT, status TEXT DEFAULT 'Pending', approved_by TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP
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
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,cycle TEXT,goals TEXT,rating REAL,manager_comments TEXT,status TEXT DEFAULT 'Open'
);
CREATE TABLE IF NOT EXISTS assets(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,asset_code TEXT,name TEXT,category TEXT,serial_no TEXT,status TEXT DEFAULT 'Available',
 employee_id INTEGER,issued_date TEXT,return_date TEXT,UNIQUE(company_id,asset_code)
);
CREATE TABLE IF NOT EXISTS expenses(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,category TEXT,amount REAL,expense_date TEXT,description TEXT,status TEXT DEFAULT 'Pending'
);
CREATE TABLE IF NOT EXISTS documents(
 id SERIAL PRIMARY KEY,company_id INTEGER NOT NULL,employee_id INTEGER,name TEXT,doc_type TEXT,expiry_date TEXT,status TEXT DEFAULT 'Active'
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
for(const col of ["smtp_user TEXT","smtp_pass TEXT"]){
  try{await db.exec(`ALTER TABLE companies ADD COLUMN ${col}`)}catch(e){}
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
  for(const [n,b] of [["Casual Leave",12],["Sick Leave",12],["Earned Leave",18],["Unpaid Leave",0]]){
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

app.post("/api/login",wrap(async(req,res)=>{
  const u=await db.prepare("SELECT * FROM users WHERE username=? AND active=1").get(req.body.username||"");
  if(!u||!verify(req.body.password||"",u.password_hash))return res.status(401).json({error:"Invalid username or password"});
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
  res.json(rows);
}));
app.post("/api/employees",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await db.prepare(`INSERT INTO employees(company_id,employee_code,name,email,phone,department,designation,manager,branch,joining_date,status,biometric_id,salary,bank_name,bank_account,ifsc)
      VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(req.user.company_id,x.employee_code,x.name,x.email,x.phone,x.department,x.designation,x.manager,x.branch,x.joining_date,x.status||"Active",x.biometric_id,x.salary||0,x.bank_name,x.bank_account,x.ifsc);
    await db.prepare("INSERT OR IGNORE INTO onboarding(company_id,employee_id) VALUES(?,?)").run(req.user.company_id,r.lastInsertRowid);
    await audit(req,"CREATE","EMPLOYEE",x.employee_code);res.json({id:r.lastInsertRowid});
  }catch(e){res.status(400).json({error:/duplicate key|unique/i.test(e.message)?"Employee code already exists":e.message})}
}));
app.put("/api/employees/:id",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;
  try{
    const r=await db.prepare(`UPDATE employees SET employee_code=?,name=?,email=?,phone=?,department=?,designation=?,manager=?,branch=?,joining_date=?,status=?,biometric_id=?,salary=?,bank_name=?,bank_account=?,ifsc=? WHERE id=? AND company_id=?`)
      .run(x.employee_code,x.name,x.email,x.phone,x.department,x.designation,x.manager,x.branch,x.joining_date,x.status||"Active",x.biometric_id,x.salary||0,x.bank_name,x.bank_account,x.ifsc,req.params.id,req.user.company_id);
    if(r.changes===0)return res.status(404).json({error:"Employee not found"});
    await audit(req,"UPDATE","EMPLOYEE",x.employee_code);res.json({ok:true});
  }catch(e){res.status(400).json({error:e.message})}
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

app.get("/api/departments",auth,requireCompany,wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM departments WHERE company_id=? ORDER BY name").all(req.user.company_id))));
app.post("/api/departments",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  try{await db.prepare("INSERT INTO departments(company_id,name) VALUES(?,?)").run(req.user.company_id,req.body.name);res.json({ok:true})}
  catch(e){res.status(400).json({error:"Department already exists"})}
}));

app.get("/api/attendance",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT a.*,e.employee_code,e.name,e.department FROM attendance a JOIN employees e ON e.id=a.employee_id WHERE a.company_id=?`;
  let params=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND e.id=?";params.push(req.user.employee_id)}
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
  q+=" ORDER BY l.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/leaves",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const r=await db.prepare(`INSERT INTO leave_requests(company_id,employee_id,leave_type,from_date,to_date,days,reason) VALUES(?,?,?,?,?,?,?)`)
    .run(req.user.company_id,eid,req.body.leave_type,req.body.from_date,req.body.to_date,Number(req.body.days)||1,req.body.reason||"");
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
  const used=await db.prepare(`SELECT leave_type,COALESCE(SUM(days),0) d FROM leave_requests WHERE employee_id=? AND company_id=? AND status='Approved' AND from_date LIKE ? GROUP BY leave_type`).all(eid,req.user.company_id,year+"%");
  const usedMap=Object.fromEntries(used.map(u=>[u.leave_type,u.d]));
  res.json(types.map(t=>({leave_type:t.name,annual_balance:t.annual_balance,used:usedMap[t.name]||0,remaining:t.annual_balance-(usedMap[t.name]||0)})));
}));

app.get("/api/payroll",auth,requireCompany,roles("Super Admin","HR Admin","Finance","Manager"),wrap(async(req,res)=>{
  res.json(await db.prepare(`SELECT p.*,e.employee_code,e.name FROM payroll p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=? ORDER BY p.id DESC`).all(req.user.company_id));
}));
app.post("/api/payroll",auth,requireCompany,roles("Super Admin","HR Admin","Finance"),wrap(async(req,res)=>{
  const x=req.body;
  const emp=await db.prepare("SELECT * FROM employees WHERE id=? AND company_id=?").get(x.employee_id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const gross=Number(x.gross)||0,deductions=Number(x.deductions)||0,lop=Number(x.lop)||0,ot=Number(x.ot)||0,net=gross-deductions-lop+ot;
  const no="BMS-"+Date.now();
  await db.prepare(`INSERT INTO payroll(company_id,employee_id,month,gross,deductions,lop,ot,net,status,payslip_no) VALUES(?,?,?,?,?,?,?,?,?,?)
    ON CONFLICT(employee_id,month) DO UPDATE SET gross=excluded.gross,deductions=excluded.deductions,lop=excluded.lop,ot=excluded.ot,net=excluded.net,status=excluded.status,payslip_no=excluded.payslip_no`)
    .run(req.user.company_id,x.employee_id,x.month,gross,deductions,lop,ot,net,x.status||"Processed",no);
  await audit(req,"UPSERT","PAYROLL",x.month);
  if(emp.email){
    const company=await db.prepare("SELECT name,smtp_user,smtp_pass FROM companies WHERE id=?").get(req.user.company_id);
    const inr=n=>"₹"+Number(n).toLocaleString("en-IN");
    sendMail(emp.email,`Payslip for ${x.month} — ${company?.name||"BMS HRMS"}`,layout(`Payslip — ${x.month}`,
      `<p>Hi ${emp.name},</p><p>Your payslip for <b>${x.month}</b> has been processed.</p>
       <table style="width:100%;border-collapse:collapse;margin-top:10px">
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Payslip No</td><td style="padding:6px;border:1px solid #e5e7eb">${no}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Gross</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(gross)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Deductions</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(deductions)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">LOP</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(lop)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb">Overtime</td><td style="padding:6px;border:1px solid #e5e7eb">${inr(ot)}</td></tr>
       <tr><td style="padding:6px;border:1px solid #e5e7eb"><b>Net Pay</b></td><td style="padding:6px;border:1px solid #e5e7eb"><b>${inr(net)}</b></td></tr>
       </table>
       <p style="font-size:12px;color:#64748b">Log in to the HRMS to view or print your full payslip.</p>`),
      {smtp_user:company?.smtp_user,smtp_pass:company?.smtp_pass,name:company?.name}).catch(()=>{});
  }
  res.json({ok:true,net,payslip_no:no});
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

app.get("/api/performance",auth,requireCompany,wrap(async(req,res)=>{
  let q=`SELECT p.*,e.employee_code,e.name FROM performance p JOIN employees e ON e.id=p.employee_id WHERE p.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND e.id=?";p.push(req.user.employee_id)}q+=" ORDER BY p.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/performance",auth,requireCompany,roles("Super Admin","HR Admin","Manager"),wrap(async(req,res)=>{
  const x=req.body;
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(x.employee_id,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const r=await db.prepare("INSERT INTO performance(company_id,employee_id,cycle,goals,rating,manager_comments,status) VALUES(?,?,?,?,?,?,?)").run(req.user.company_id,x.employee_id,x.cycle,x.goals,x.rating||0,x.manager_comments,x.status||"Open");
  res.json({id:r.lastInsertRowid});
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
  if(req.user.role==="Employee"){q+=" AND x.employee_id=?";p.push(req.user.employee_id)}q+=" ORDER BY x.id DESC";res.json(await db.prepare(q).all(...p));
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
  let q=`SELECT d.*,e.name employee_name,e.employee_code FROM documents d JOIN employees e ON e.id=d.employee_id WHERE d.company_id=?`;let p=[req.user.company_id];
  if(req.user.role==="Employee"){q+=" AND d.employee_id=?";p.push(req.user.employee_id)}q+=" ORDER BY d.id DESC";res.json(await db.prepare(q).all(...p));
}));
app.post("/api/documents",auth,requireCompany,wrap(async(req,res)=>{
  const eid=req.user.role==="Employee"?req.user.employee_id:Number(req.body.employee_id);
  const emp=await db.prepare("SELECT id FROM employees WHERE id=? AND company_id=?").get(eid,req.user.company_id);
  if(!emp)return res.status(404).json({error:"Employee not found in this company"});
  const x=req.body;const r=await db.prepare("INSERT INTO documents(company_id,employee_id,name,doc_type,expiry_date,status) VALUES(?,?,?,?,?,?)").run(req.user.company_id,eid,x.name,x.doc_type,x.expiry_date,x.status||"Active");res.json({id:r.lastInsertRowid});
}));

app.get("/api/announcements",auth,requireCompany,wrap(async(req,res)=>res.json(await db.prepare("SELECT * FROM announcements WHERE company_id=? ORDER BY id DESC").all(req.user.company_id))));
app.post("/api/announcements",auth,requireCompany,roles("Super Admin","HR Admin"),wrap(async(req,res)=>{
  const x=req.body;const r=await db.prepare("INSERT INTO announcements(company_id,title,body,audience) VALUES(?,?,?,?)").run(req.user.company_id,x.title,x.body,x.audience||"All");
  res.json({id:r.lastInsertRowid});
  (async()=>{
    const emails=(await db.prepare("SELECT DISTINCT email FROM employees WHERE company_id=? AND status='Active' AND email IS NOT NULL AND email<>''").all(req.user.company_id)).map(r=>r.email);
    if(!emails.length)return;
    const sender=await companySender(req.user.company_id);
    const html=layout(x.title,`<p>${(x.body||"").replace(/\n/g,"<br>")}</p><p style="font-size:12px;color:#64748b">Audience: ${x.audience||"All"}</p>`);
    for(const email of emails) await sendMail(email,`Announcement: ${x.title}`,html,sender);
  })().catch(e=>console.error("Announcement email batch failed:",e.message));
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
