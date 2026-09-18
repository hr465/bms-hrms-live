# BMS Enterprise HRMS — Live v1

This is a locally runnable enterprise HRMS starter/demo for Bhartiya Management Solutions.

## Included working areas
- Login/logout and server-side sessions
- Role-based access: Super Admin, HR Admin, Manager, Finance, Employee
- Employee self-service scope
- Employee master
- Dashboard / MIS foundation
- Attendance and manual attendance
- eSSL-ready biometric device registry + punch ingestion API
- Leave management
- Payroll processing foundation
- Recruitment / ATS
- Digital onboarding
- Performance / KRA/KPI foundation
- Asset management
- Expense/reimbursement
- Employee documents
- Announcements
- HR helpdesk/tickets
- Exit/separation
- Audit log
- Configurable HR policy foundation

## Demo credentials
admin / Admin@12345
hr / HR@12345
manager / Manager@12345
finance / Finance@12345
employee / Employee@12345

## Run on Windows
Open CMD in this folder:
npm install
npm start

Then open:
http://localhost:3000

## Important
This is a functional local v1/starter, not a production-certified HR/payroll product. Before production deployment, add PostgreSQL, HTTPS, encrypted secret management, backups, CSRF/rate limiting, email/SMS provider, document storage, statutory payroll validation, and a tested eSSL connector for the exact device model.

For eSSL:
POST /api/biometric/punch
{
  "biometric_id":"1001",
  "punch_time":"2026-08-13T09:30:00+05:30",
  "punch_type":"IN",
  "device_id":1,
  "raw_payload":{}
}

The exact eSSL connector must be matched to the model/firmware and its supported API/SDK/ADMS mechanism.
