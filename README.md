# BMS Enterprise HRMS

Multi-company HR management platform for Bhartiya Management Solutions.

## Modules

- Multi-company workspaces with isolated data and role-based access (Super Admin, HR Admin, Director, Manager, Finance, Employee)
- Employee master with statutory details (PF, ESIC, UAN, PAN), salary structure and reporting hierarchy
- Attendance with eSSL / ZKTeco biometric integration (LAN sync and a local sync agent)
- Leave, work from home and permission requests with balances
- Payroll with automatic LOP, PF and ESIC calculation, bulk monthly run and emailed payslips
- Performance reviews, rating-based yearly increments and Employee of the Month recognition
- Onboarding agreements with sequential e-signatures (Employee, HR, Director)
- Recruitment, assets, expenses, documents, announcements, helpdesk and exit management
- MIS reports and Excel / CSV downloads
- Audit log

## Configuration

Set these environment variables:

| Variable | Purpose |
| --- | --- |
| `DATABASE_URL` | PostgreSQL connection string |
| `GMAIL_USER`, `GMAIL_APP_PASSWORD` | Default outgoing email account (companies can configure their own in Settings) |
| `MAIL_FROM_NAME` | Sender display name |
| `NODE_ENV` | Set to `production` to enable secure cookies |

## Run locally

```
npm install
npm start
```

The application listens on `PORT` (default 3000). On first start it creates the schema and a platform administrator account; change its password immediately after signing in.

## Biometric sync agent

For devices on a private network, run `sync-agent.js` on a PC in the same LAN. See `sync-agent.config.example.json` and the Biometric page in the application.
