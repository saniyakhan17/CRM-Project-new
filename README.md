# scripts
GS scripts
# CRM Script

A full-stack CRM system built on Google Apps Script, designed to manage leads, track sales performance, and streamline BDA workflows for a B2B SaaS company.

---

## Overview

This project is a custom CRM solution built entirely using Google Apps Script and Google Sheets as the data layer. It includes a web app portal with role-based access, automated lead management, Power BI integration, and email notification triggers.

---

## Tech Stack

| Layer | Technology |
|---|---|
| Backend Logic | Google Apps Script (GAS) |
| Frontend Portal | HTML, CSS, JavaScript (GAS Web App) |
| Data Storage | Google Sheets |
| BI Dashboard | Power BI (via Google Sheets CSV export) |
| Automation | Google Apps Script Triggers |

---

## File Structure

```
├── Code.gs         # Main backend logic — lead management, triggers, role handling
├── WebApp.gs       # Web app routing and server-side rendering
├── Index.html      # Main portal UI (dashboard, leads table, performance charts)
├── Login.html      # Login page with session token authentication
├── appsscript.json # Project manifest and deployment config
```

---

## Key Features

- **Role-Based Access Control** — Separate views and permissions for Admin, Team Leaders, and BDAs
- **Automated Lead Assignment** — Round-robin assignment to BDAs based on lowest lead count
- **Lead Pipeline Tracking** — Full pipeline from New Lead → Won with status color coding
- **Follow-up Management** — Follow-up date tracking with overdue highlighting
- **WhatsApp Integration** — One-click WhatsApp button with pre-filled message
- **Stale Lead Email Notifications** — Automated email alerts for leads with no activity
- **Power BI Dashboard** — Real-time performance tracking connected to Google Sheets
- **CRM Maintenance Functions** — Data cleaning, Lead ID generation, BDA portal sync

---

## Data Architecture

| Sheet | Purpose |
|---|---|
| Leads_Master | Central lead database |
| Ad_Leads_Sync | Incoming leads from Google Form |
| Lookups | Role config, passwords, dropdown values |

---

## Deployment

The web app is deployed via Google Apps Script as a public web app:
- **Execute as:** User deploying (Admin)
- **Access:** Anyone (authenticated via internal login system)

---

## Documentation

Full SOP documentation (v1.2) covering BDA workflows, Team Leader workflows, Admin controls, onboarding guide, and Power BI maintenance is maintained separately.

---

## Author

**Saniya Khan** — Developer & Admin  
[LinkedIn](http://www.linkedin.com/in/saniya--khan) | saniyakhan1709@gmail.com
