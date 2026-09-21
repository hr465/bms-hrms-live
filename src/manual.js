// Builds the HR portal training manual as a PDF. Wording is generic; the cover and addresses are filled in per company.
const PDFDocument=require("pdfkit");
function buildManualPdf(opts){
return new Promise((resolve,reject)=>{
const O={company:"Your Company",portalUrl:"",platformUrl:"",serverIp:"15.252.60.243",platform:"HR Platform",...opts};
const doc=new PDFDocument({size:"A4",margins:{top:60,bottom:60,left:56,right:56},bufferPages:true,info:{Title:"HR Portal Training Manual",Author:O.platform}});
const bufs=[];doc.on("data",b=>bufs.push(b));doc.on("end",()=>resolve(Buffer.concat(bufs)));doc.on("error",reject);
const C={ink:"#0f172a",mut:"#475569",brand:"#312e81",acc:"#4f46e5",soft:"#eef2ff",line:"#e2e8f0",warn:"#fef3c7",warnInk:"#92400e",ok:"#dcfce7",okInk:"#166534"};
const W=doc.page.width-112;
const need=h=>{if(doc.y+h>doc.page.height-doc.page.margins.bottom)doc.addPage()};
let chapterNo=0;
function chapter(title,sub){doc.addPage();chapterNo++;
  doc.rect(0,0,doc.page.width,86).fill(C.brand);
  doc.font("Helvetica").fontSize(10).fillColor("#c7d2fe").text("CHAPTER "+chapterNo,56,26);
  doc.font("Helvetica-Bold").fontSize(22).fillColor("#fff").text(title,56,42,{width:W});
  doc.y=110;doc.x=56;if(sub){doc.font("Helvetica-Oblique").fontSize(10.5).fillColor(C.mut).text(sub,{width:W});doc.moveDown(0.8)}}
function h2(t){need(50);doc.moveDown(0.6);doc.font("Helvetica-Bold").fontSize(13).fillColor(C.brand).text(t,{width:W});
  const y=doc.y+2;doc.moveTo(56,y).lineTo(56+W,y).lineWidth(.8).strokeColor(C.line).stroke();doc.moveDown(0.5)}
function p(t){need(30);doc.font("Helvetica").fontSize(10.5).fillColor(C.ink).text(t,{width:W,lineGap:2.5});doc.moveDown(0.5)}
function bullets(a){for(const t of a){need(24);const y=doc.y;doc.font("Helvetica").fontSize(10.5).fillColor(C.acc).text("•",64,y);doc.fillColor(C.ink).text(t,80,y,{width:W-26,lineGap:2.5});doc.x=56}doc.moveDown(0.4)}
function steps(a){a.forEach((t,i)=>{need(30);const y=doc.y;
  doc.circle(70,y+7,9).fill(C.acc);doc.font("Helvetica-Bold").fontSize(9).fillColor("#fff").text(String(i+1),62,y+3,{width:16,align:"center"});
  doc.font("Helvetica").fontSize(10.5).fillColor(C.ink).text(t,88,y,{width:W-34,lineGap:2.5});doc.x=56;doc.moveDown(0.35)});doc.moveDown(0.3)}
function box(t,kind){const bg=kind==="ok"?C.ok:kind==="warn"?C.warn:C.soft,ink=kind==="ok"?C.okInk:kind==="warn"?C.warnInk:C.brand;
  doc.font("Helvetica").fontSize(10);const h=doc.heightOfString(t,{width:W-24,lineGap:2})+18;need(h+6);const y=doc.y;
  doc.roundedRect(56,y,W,h,5).fill(bg);doc.rect(56,y,3.5,h).fill(ink);doc.fillColor(ink).text(t,70,y+9,{width:W-26,lineGap:2});doc.x=56;doc.y=y+h+8}
function table(head,rows,widths){const tw=widths.reduce((a,b)=>a+b,0),sc=W/tw,ws=widths.map(w=>w*sc);
  const draw=(cells,hd)=>{doc.font(hd?"Helvetica-Bold":"Helvetica").fontSize(9.5);
    const h=Math.max(...cells.map((c,i)=>doc.heightOfString(String(c),{width:ws[i]-12,lineGap:1.5})))+12;need(h+2);
    const y=doc.y;let x=56;if(hd)doc.rect(56,y,W,h).fill(C.soft);
    cells.forEach((c,i)=>{doc.fillColor(hd?C.brand:C.ink).text(String(c),x+6,y+6,{width:ws[i]-12,lineGap:1.5});x+=ws[i]});
    doc.moveTo(56,y+h).lineTo(56+W,y+h).lineWidth(.5).strokeColor(C.line).stroke();doc.x=56;doc.y=y+h};
  draw(head,true);rows.forEach(r=>draw(r,false));doc.moveDown(0.7)}

// ---------------- Cover ----------------
doc.rect(0,0,doc.page.width,doc.page.height).fill(C.brand);
doc.rect(0,300,doc.page.width,4).fill("#818cf8");
doc.font("Helvetica").fontSize(12).fillColor("#c7d2fe").text("TRAINING MANUAL",56,200);
doc.font("Helvetica-Bold").fontSize(34).fillColor("#fff").text("HR Portal",56,225);
doc.font("Helvetica").fontSize(16).fillColor("#e0e7ff").text("Step-by-step guide for every role",56,325);
doc.font("Helvetica-Bold").fontSize(14).fillColor("#fff").text(O.company,56,372);
doc.font("Helvetica").fontSize(11).fillColor("#c7d2fe").text("Portal address: "+O.portalUrl,56,395);
doc.fontSize(10).text("Powered by "+O.platform,56,760);

// ---------------- Contents ----------------
doc.addPage();doc.font("Helvetica-Bold").fontSize(22).fillColor(C.brand).text("Contents",56,70);doc.moveDown(0.8);
const toc=[["1","How the portal works and who can do what"],["2","Platform setup for a new company (Super Admin)"],["3","First-day setup by the company HR Admin"],["4","Adding employees and their logins"],["5","Connecting the biometric device"],["6","Daily work: attendance, clock in and out, leave"],["7","Overtime and working hours reports"],["8","Payroll, payslips and salary payment"],["9","Agreements, letters and onboarding"],["10","Performance, increments and recognition"],["11","Announcements, helpdesk, exits and other modules"],["12","Guide for the Manager role"],["13","Guide for the Finance role"],["14","Guide for the Director role"],["15","Guide for the Employee role"],["16","Reports and downloads"],["17","Troubleshooting and frequently asked questions"],["18","Go-live checklist"],["19","Putting the portal on your own web address"]];
toc.forEach(([n,t])=>{const y=doc.y;doc.font("Helvetica-Bold").fontSize(11).fillColor(C.acc).text(n,64,y,{width:26,lineBreak:false});doc.font("Helvetica").fillColor(C.ink).text(t,96,y,{width:W-60});doc.x=56;doc.moveDown(0.35)});

// ---------------- 1 ----------------
chapter("How the portal works","Read this first. It explains the roles and how the pieces fit together.");
h2("What the portal does");
p("The HR portal keeps everything about your people in one place: employee records, attendance from the biometric device, leave, payroll and payslips, agreements and letters, performance and increments. It sends email notifications automatically and every company sees only its own data.");
h2("Roles and what each can do");
table(["Role","Who","Main responsibilities"],[
["Super Admin","Platform team","Onboards companies, edits company details, resets logins. Not used by company staff."],
["HR Admin","Company HR","Full control of the company: employees, attendance, leave, payroll, letters, settings, biometric, logins."],
["Director","Management","Views reports, MIS, performance and increments. Gives the final signature on agreements."],
["Manager","Team leaders","Sees own team's attendance, approves leave, reviews performance, sees the overtime report of the team."],
["Finance","Accounts","Runs payroll, marks salary as paid with the bank UTR, exports payroll data, handles expenses."],
["Employee","Every staff member","Own attendance, clock in and out, leave, payslips, letters, agreements, helpdesk."]],[1.1,1.3,4]);
h2("The usual order of work");
steps(["The platform team onboards the company (Chapter 2).","HR completes the first-day setup (Chapter 3).","HR adds employees and creates their logins (Chapter 4).","HR connects the biometric device (Chapter 5).","Everyday attendance, leave and monthly payroll begin (Chapters 6 to 8)."]);
box("Every list in the portal can be searched from the search box above it. Most pages have Download buttons (Excel or CSV) at the top right.","info");

// ---------------- 2 ----------------
chapter("Platform setup for a new company","For the platform administrator. Sign in at "+O.platformUrl);
h2("Onboard the company");
steps(["Sign in as Super Admin and open Companies in the left menu.","Click + Onboard New Company.","Optional: click Choose image to add the company logo (PNG or JPG).","Enter the company name, a short code (for example MEPL), industry, address, contact email and phone.","Enter the first HR Admin username and a password of at least 8 characters. Share these with the company HR.","Optional: enter the sender Gmail and its app password. This is the address from which the company's emails (payslips, letters, notices) are sent.","Click Onboard Company."]);
h2("Change company details later");
bullets(["Companies, then Edit next to the company. You can change name, industry, address, contact details, logo and the sender Gmail. Leave the app password blank to keep the current one.","Logins button: lists every login of that company and lets you reset a forgotten password. A new password is generated and shown once.","Domain button: gives the company its own web address, for example hr.yourcompany.com. The login page then shows that company's name and logo and only its users can sign in there.","Manage button: switches you into the company workspace to help its HR."]);
h2("Sender Gmail for emails");
steps(["Use a Google Workspace or Gmail account that belongs to the company, for example hr@company.com.","Turn on 2-Step Verification for that account.","Create an App Password (Google Account, Security, App passwords).","Enter the account and the 16-character app password in Edit company, then save."]);
box("Company HR can check that the sender works from Settings, Email log, Send test email. Every sent or failed email is listed there with the reason.","info");

// ---------------- 3 ----------------
chapter("First-day setup by the company HR Admin","Sign in at "+O.portalUrl+" with the username and password given to you.");
h2("1. Change your password");
p("Click Password at the top right, enter the current password and a new one of at least 8 characters.");
h2("2. Settings (menu: HR Policies)");
p("Open HR Policies in the Administration section. It has these panels:");
table(["Panel","What to do"],[
["Company profile","Enter industry, address, contact email and phone. These appear on letters and emails."],
["Letterhead","Upload the company letterhead as a full A4 page image. Letters and agreements print on it. Keep the default top margin 130 and bottom margin 170 unless the text touches the header or footer."],
["Working hours","Set start time, close time, grace minutes, the break, the minimum daily hours and the overtime rule. See Chapter 7."],
["Clock In / Out location","Choose whether location is required and whether staff may clock in from outside the office. Use Use my current location as the office while standing in the office."],
["Email log","Shows every email sent and whether it worked. Use Send test email once after setup."],
["Email Settings","The sender Gmail and app password (also editable by the platform team)."],
["Policy Agreement","The text of the agreement given to every new employee. See Chapter 9."]],[1.3,4.5]);
h2("3. Departments and leave types");
p("Departments and leave types (Casual, Sick, Privilege, Unpaid) are created automatically for every new company. Employees pick a department when they are added.");
h2("4. Basic policies to set up first");
bullets(["Working hours: start and close time, grace minutes, break, minimum daily hours (Settings, Working hours).","Overtime rule: turn it on or off and set the smallest overtime that counts.","Clock In / Out and location rules for staff without a biometric punch.","Leave types and yearly balances (Leave page and Settings).","Policy Agreement text with placeholders such as {{employee_name}}, so it fills in for every employee.","Appointment letter template (Letters page) and the letterhead image.","Email sender and a test email, so payslips, letters and notices reach your staff."]);
h2("5. Logins for management and staff");
steps(["Open Team in the Administration section.","Add a login for each Director, Finance user, Manager or extra HR user: username, password, role and email.","To link the login to a person in the employee list, choose the employee. Managers must be linked to their employee record."]);
box("Managers, Finance and Directors receive email alerts (leave requests, agreements waiting for signature, daily attendance summary) only if their email address is filled in.","warn");

// ---------------- 4 ----------------
chapter("Adding employees and their logins","HR Admin");
h2("Add an employee");
steps(["Open Employees and click + Add Employee.","Fill in employee code (unique, for example MEPL-027), name, email, phone, department, designation, branch and joining date.","Biometric ID: enter exactly the user ID that the person has on the biometric device. If it does not match, the punches will not be linked.","Choose the reporting manager. Leave requests go to this person first.","Enter the salary structure: basic, HRA and other allowances. Tick PF or ESIC where applicable, and fill bank details, PAN and UAN.","Click Save Employee. A welcome email is sent to the employee's email address."]);
h2("Buttons in the employee list");
table(["Button","What it does"],[
["Edit","Change any detail of the employee."],
["Photo","Upload the employee's profile photo."],
["Timing","Set a different start, close, break or minimum hours for this person. Tick Use company default to go back to the company timing."],
["Create Login","Creates the employee's own login. The username is the employee code in small letters (MEPL-027 becomes mepl-027). A temporary password is generated and emailed with the login link."],
["Agreement","Generates the onboarding agreement for signing."],
["Deactivate / Activate","Stops or restores the employee. Deactivated staff are ignored in payroll and reports."]],[1.4,4.4]);
box("Create the login while you are signed in at your own portal address ("+O.portalUrl+"), so that the emailed link points to it.","warn");
h2("If the employee did not get the login email");
bullets(["Check that the email address on the employee record is correct.","Open Settings, Email log and look at the status and reason.","The popup shows the temporary password when no email could be sent, so you can share it manually.","A super admin can reset any login password from Companies, Logins."]);

// ---------------- 5 ----------------
chapter("Connecting the biometric device","HR Admin. Works with eSSL and ZKTeco devices that support ADMS (cloud push).");
h2("What happens");
p("The device sends every punch to the portal over the internet by itself. Nobody needs to keep a computer running at the device's office, and HR can see the data from any city.");
h2("Step 1: register the device in the portal");
steps(["Open Biometric / eSSL in the Administration section.","Choose the connection type ADMS push.","Enter the device name (for example Udaipur Office), the model (for example X2008), the branch and the device serial number.","The serial number is on the device: Menu, System Info, Device Info.","Click Add Device."]);
h2("Step 2: set the device");
steps(["On the device open Menu, Comm., Cloud Server Setting.","Server Mode: ADMS.","Server Address: "+O.serverIp+" (the portal server).","Server Port: 80.","Enable Proxy Server: off. Save and restart the device."]);
p("Within a minute or two the device shows as Online on the Biometric page. Punches then arrive in real time and the first contact also loads recent history.");
h2("Step 3: match people");
p("Each employee's Biometric ID in the portal must be the same as the user ID on the device. Add users on the device first, then create the employee with that ID. If you add the ID later, past punches are linked automatically.");
table(["Status","Meaning"],[["Online","The device has contacted the portal in the last few minutes."],["Delayed","No contact for a while. Check the device's internet."],["Offline","No contact for a long time. Check power, internet and the server settings on the device."]],[1.2,4.6]);
box("Each device serial number can be registered to only one company. Give each new company its own device settings; the server address and port are the same for all.","info");

// ---------------- 6 ----------------
chapter("Daily work: attendance and leave","HR, Managers and Employees");
h2("Attendance page");
p("Shows one row per person per day: first in, last out, location, late minutes, worked time, overtime, status and source (biometric, web or manual). Hours shown in red mean the minimum hours were not completed. HR can add or correct a day with the manual entry form.");
h2("Clock In and Clock Out (Employee)");
steps(["Open Attendance. The Today box shows your status.","Click Clock In when you start work. Allow location when the browser asks.","Click Clock Out at the end of the day.","The location is saved and marked Office or Remote. The time is taken from the server, so it cannot be changed."]);
box("This is for staff who have no biometric punch, for example people working remotely. HR can switch it off in Settings, Working hours.","info");
h2("Applying for leave, work from home or permission (Employee)");
steps(["Open Leave and choose the category: Leave, WFH or Permission.","Choose the leave type, dates and days, and give a reason.","Submit. The reporting manager, HR and Directors are emailed. You are emailed when the request is approved or rejected."]);
h2("Approving requests (Manager or HR)");
steps(["Open Leave. Pending requests are listed.","Click Approved or Rejected. The employee is informed by email.","Approved unpaid leave reduces salary in payroll (loss of pay)."]);
table(["Statuses","Effect on salary"],[["Present","Full pay for the day"],["Half Day","Half a day of loss of pay"],["Absent","Full day of loss of pay"],["Unpaid leave (approved)","Loss of pay for those days"]],[2,3.8]);

// ---------------- 7 ----------------
chapter("Overtime and working hours","HR Admin sets the rules. Directors, Finance, Managers and HR see the report.");
h2("Working hours (Settings, HR Policies)");
table(["Setting","Meaning"],[
["Start time and close time","The official shift of the company."],
["Grace (minutes)","Minutes after the start time that are not counted as late."],
["Deduct break time","When on, the break length is removed from worked hours."],
["Minimum working hours per day","Days with less than this are shown in red and reported to HR."],
["Email HR alerts","Choose whether HR is emailed about late arrivals and about people who did not complete the minimum hours."],
["Overtime","Turn overtime counting on or off, and set the smallest overtime (in minutes) that is counted."]],[2,3.8]);
p("A summary of late arrivals and short hours is emailed to HR and Directors every working day, thirty minutes after the close time. Use Send today's summary now to test it.");
h2("How overtime is counted");
p("Overtime is the time worked after the shift close time. For a shift ending at 18:30, a last punch at 19:45 gives 1 hour 15 minutes of overtime, if it is at least the smallest overtime you set. Each employee can have a personal shift and overtime on or off under Employees, Timing.");
h2("Overtime Report");
steps(["Open Overtime Report in the Overview section.","Choose the month.","The top table shows each employee's overtime days and total time. The lower table shows every day.","Use Summary (Excel), Daily (Excel) or CSV to download."]);

// ---------------- 8 ----------------
chapter("Payroll, payslips and salary payment","HR Admin and Finance");
h2("Monthly run");
steps(["Make sure attendance and approved leave for the month are complete.","Open Payroll. Under Run payroll for the whole company, choose the month and click Run Payroll for Month.","The portal calculates gross, loss of pay days from attendance and unpaid leave, PF (12% of basic) and ESIC (0.75% of gross when gross is up to Rs. 21,000).","Each employee is emailed a PDF payslip.","For one person, use the individual payroll form to enter TDS, other deductions or overtime pay."]);
box("Basic and HRA appear on the payslip only if the salary structure is filled in on the employee record. Fill it in before running payroll.","warn");
h2("Marking salary as paid (with the bank UTR)");
steps(["After the bank transfer, open Payroll.","For one person click Mark paid on their row, enter the bank UTR or reference, the payment date and the mode (NEFT, IMPS, RTGS, UPI, Cheque or Cash).","If one bank batch paid everyone, use Mark the whole month as paid: choose the month, enter the UTR and date, and click Mark paid and email.","Each employee is emailed Salary credited with the UTR, date and the payslip PDF."]);
h2("The payslip PDF");
bullets(["Company name, address and logo at the top.","Employee details: code, designation, department, joining date, PAN, UAN, PF and ESIC numbers, bank and masked account number.","Days in the month, paid days and loss of pay days.","Earnings and deductions with totals, and Net Pay with the amount in words.","Payment details, including the UTR once marked paid.","Use the PDF button in the payroll list to open it. Employees open their own payslips the same way."]);
box("Running payroll again for a month that is already paid clears the Paid status and UTR, because the figures may have changed. Mark it as paid again.","warn");

// ---------------- 9 ----------------
chapter("Agreements, letters and onboarding","HR Admin, Director, Employee");
h2("Policy Agreement text");
steps(["Open HR Policies and find Policy Agreement.","Paste the agreement text. Use placeholders so it fills in for each person.","Save. The text applies to agreements created afterwards."]);
p("Placeholders: {{employee_name}}, {{employee_code}}, {{designation}}, {{department}}, {{branch}}, {{reporting_manager}}, {{joining_date}}, {{company_name}}, {{company_address}}, {{issue_date}}, {{basic_monthly}}, {{hra_monthly}}, {{other_allowances_monthly}}, {{monthly_ctc}}, {{annual_ctc}}. A line starting with # is a centred title and a line starting with ## is a heading.");
h2("Sending an agreement for signing");
steps(["Open Employees and click Agreement next to the person (or use the Agreements page).","The employee is emailed to sign. The employee opens Agreements, clicks Review & Sign and signs.","HR is emailed and signs next. Then the Director is emailed and signs.","When all three have signed, the signed PDF on the letterhead is emailed to the employee and to the company."]);
p("Anyone involved can open the PDF with View PDF on the agreement.");
h2("Appointment letters");
steps(["Open Letters. Edit the letter template if needed. It uses the same placeholders.","Open the employee, preview the letter, then Issue and email.","The letter is generated as a PDF on the letterhead, emailed to the employee and stored under Letters."]);
box("The employee needs a designation, joining date and salary structure filled in before a letter can be issued.","warn");
h2("Onboarding checklist");
p("The Onboarding page tracks each new joiner's steps: offer, documents, verification, assets and policy. Tick each step as it is done.");

// ---------------- 10 ----------------
chapter("Performance, increments and recognition","HR Admin, Managers and Directors");
h2("Performance reviews");
steps(["Open Performance and create a review for an employee for the cycle, with a rating from 1 to 5.","Managers review their own team.","HR finalises the review. A finalised review is locked."]);
h2("Yearly increments");
p("The Increments page suggests an increment for each employee from the finalised rating, using slabs that HR can edit (for example rating 4.5 and above gives 15 percent). HR reviews and applies them, and the employee is emailed the salary revision.");
h2("Employee of the Month");
p("Employee Recognition suggests the top employees of the month from a score that combines rating (40%), attendance (35%) and punctuality (25%). Open the page and choose the month.");

// ---------------- 11 ----------------
chapter("Announcements, helpdesk, exits and other modules");
table(["Page","Use"],[
["Announcements","HR posts notices. Every employee with an email address also receives the notice by email."],
["HR Helpdesk","Employees raise questions or issues. HR replies and closes them."],
["Exit / Separation","Records resignations, notice period and exit steps."],
["Recruitment","Track candidates and interview status."],
["Assets","Record company assets and who holds them."],
["Documents","Store and share employee documents."],
["Expenses","Employees submit expense claims for approval."],
["Calendar","Holidays and events."],
["Audit Log","A record of who did what, for HR and the platform team."]],[1.6,4.2]);

// ---------------- 12-15 ----------------
chapter("Guide for the Manager role","For team leaders. Your login must be created under Team with your employee record linked.");
steps(["Sign in. Your menu shows your team only.","Attendance: see your team's punches, late marks and hours.","Leave: approve or reject requests from your team. You are emailed when someone applies.","Performance: review your team members.","Overtime Report: see overtime for your team."]);
chapter("Guide for the Finance role");
steps(["Open Payroll. Run the monthly payroll or review the run done by HR.","After the bank transfer, click Mark paid on each row, or use Mark the whole month as paid with the bank UTR.","Use the Download buttons to export payroll for the bank or accounts.","Expenses: approve or reject employee expense claims.","MIS Reports and Overtime Report: check totals."]);
chapter("Guide for the Director role");
steps(["MIS Reports: headcount, attendance, leave, payroll trend, hiring and performance at a glance.","Agreements: you are emailed when an agreement needs the final signature. Open Agreements, Review & Sign.","Increments and Performance: review ratings and proposed increments.","Overtime Report and Downloads: view and export."]);
chapter("Guide for the Employee role");
h2("Signing in");
p("Open "+O.portalUrl+" and sign in with the username and temporary password from your welcome email. Change your password using the Password button at the top right. Click your round photo icon at the top to add or change your profile photo.");
h2("What you can do");
table(["Task","How"],[
["Mark attendance","Attendance, Today box: Clock In and Clock Out (if your company allows it). Otherwise the biometric device records you."],
["Apply for leave, WFH or permission","Leave, choose the category, dates and reason."],
["See leave balance","Leave page, the balance cards at the top."],
["Download payslips","Payroll, PDF button next to the month. You also receive it by email."],
["Sign the agreement","Agreements, Review & Sign. Draw your signature and type your name."],
["Get your letters","Letters, download the PDF."],
["Raise an issue","HR Helpdesk, create a ticket."],
["Submit an expense","Expenses, add the claim with the amount and date."]],[1.9,3.9]);
h2("Forgot your password");
p("On the sign-in page click Forgot password, enter your username and follow the link sent to your email. If it does not arrive, ask HR to reset it.");

// ---------------- 16 ----------------
chapter("Reports and downloads");
p("Most pages show Download buttons (Excel or CSV) for the data on that page: employees, attendance (with a date range), leave, payroll (by month), expenses, performance, recruitment, assets and increments. The Overtime Report has its own summary and daily downloads. MIS Reports can be printed.");
box("Downloads contain personal and salary data. Share them only with people who need them.","warn");

// ---------------- 17 ----------------
chapter("Troubleshooting and frequently asked questions");
table(["Problem","What to check"],[
["An employee did not receive an email","Check the email address on the employee record. Open Settings, Email log and read the status and reason. Use Send test email to confirm the sender works. Ask the employee to check the spam folder."],
["The device shows Offline","Check the device's power and internet. Check Server Address and Port on the device. Check the serial number registered in the portal is the same as on the device."],
["Punches are not linked to an employee","The employee's Biometric ID must equal the user ID on the device. Correct it under Employees, Edit."],
["Attendance shows late for everyone","Check the start time, close time and grace in Settings, Working hours."],
["Payslip has no Basic or HRA","Fill in the salary structure on the employee record and run payroll again."],
["Cannot create the login: username already taken","Another company or person already uses that employee code as a username. Use a different employee code."],
["Employee cannot sign in","Ask HR or the platform team to reset the password (Companies, Logins). Make sure they use the company address "+O.portalUrl+"."],
["Letter cannot be issued","The employee needs a designation, joining date and salary structure."],
["Text looks wrong in a PDF","Use English letters and normal punctuation. Special symbols are replaced. The rupee sign prints as Rs."],
["Logo does not show after upload","Refresh the page with Ctrl+Shift+R."]],[2,3.8]);

// ---------------- 18 ----------------
chapter("Go-live checklist","Tick each item before you announce the portal to the team.");
const chk=["Company onboarded, logo added and HR password changed","Company profile, letterhead and working hours saved","Sender Gmail set and a test email received","Departments checked and Team logins (Director, Finance, Managers) created with email addresses","All employees added with correct biometric IDs, reporting managers and salary structure","Employee logins created and login emails received","Biometric device registered, showing Online, and a test punch visible in Attendance","Policy Agreement text saved and one test agreement signed by all three parties","Appointment letter template checked and one test letter issued","Leave types and balances confirmed","One test payroll run and one payslip PDF checked, then Mark paid tried with a test UTR","Overtime rule and minimum hours confirmed with management","Training given to HR, managers and employees using this manual","Optional: own web address added (Chapter 19)"];
chk.forEach(t=>{need(24);const y=doc.y;doc.roundedRect(58,y+1,11,11,2).lineWidth(1).strokeColor(C.acc).stroke();doc.font("Helvetica").fontSize(10.5).fillColor(C.ink).text(t,80,y,{width:W-30});doc.x=56;doc.moveDown(0.5)});

chapter("Putting the portal on your own web address","Optional. Staff open your company's own address, for example https://hr.yourcompany.com, instead of the shared portal address.");
h2("What you get");
p("The login page shows your company name and logo, and only your own staff can sign in there. The shared address "+O.platformUrl+" keeps working too.");
h2("Steps");
steps(["Choose the address, normally a subdomain of your website such as hr.yourcompany.com.","Open the DNS settings of your domain (GoDaddy, Cloudflare, your hosting panel) and add ONE record: Type A, Name hr, Value "+O.serverIp+", TTL 600. If a record named hr already exists (for example a CNAME), remove it first.","Do not change the records of the main domain or www. Your website and email keep working.","Tell the platform team the address. They add it under Companies, Domain and switch on the free HTTPS certificate (one command on the server).","After a few minutes open https://hr.yourcompany.com. The login page shows your company. Share this address with your staff."]);
box("The address works only after the DNS record is in place and the platform team has activated it. Until then use "+O.platformUrl+".","info");
h2("Then the biometric device");
p("The device does not need the domain. Use the server address and port from Chapter 5. It stays the same when you add or change your web address.");

// footer with page numbers
const range=doc.bufferedPageRange();
for(let i=1;i<range.count;i++){doc.switchToPage(i);doc.page.margins.bottom=0;
  doc.font("Helvetica").fontSize(8.5).fillColor("#94a3b8").text(O.company+" • HR Portal Training Manual",56,doc.page.height-38,{width:W-60,lineBreak:false});
  doc.text("Page "+(i+1),56+W-60,doc.page.height-38,{width:60,align:"right",lineBreak:false});}
doc.end();
});
}
module.exports={buildManualPdf};
