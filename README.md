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

## Biometric sync agent (device in a different city than HR)

The portal cannot reach a device that sits on a private office network, so each office runs a small agent that uploads attendance to the portal. HR users in any city then see the data without doing anything.

1. Pick an always-on Windows PC in the same network as the device (for example in the Udaipur office) with internet access. Give the device a fixed IP address.
2. Install Node.js 22 or newer, copy this project folder to the PC and run `npm install`.
3. In the portal open Biometric, register the device and choose "Sync Agent Config". Save the shown JSON as `sync-agent.config.json` next to `sync-agent.js`.
4. Double-click `start-agent.bat`. It restarts itself if it stops. Put a shortcut to it in the Windows Startup folder so it starts on boot.

The first run imports the last 30 days. After that only new punches are sent, every 10 minutes. If the PC or internet is down, the device keeps its logs and the agent uploads everything missed once it is back. The Biometric page shows each device as Online, Delayed or Offline.

Employee Biometric IDs must match the user IDs enrolled on the device.
