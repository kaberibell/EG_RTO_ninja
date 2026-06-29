## 🥷 RTO Ninja
Automated RTO Compliance Tracker for Expedia Group Employees

Total cost: $3 (Tasker app) • Everything else is free

RTO Ninja automatically tracks the days you go into the office, calculates your compliance against Expedia's rolling RTO policy, and sends you a clear email every Sunday morning telling you exactly where you stand and how many days you need next week to stay on track. If you have a vacation coming up, it also tells you how many days per week to work beforehand so you're still compliant when you get back.

No spreadsheet math. No guessing. Just a clear answer every Sunday morning.


## How It Works
Three pieces work together:

Tasker (Android) — detects when you arrive at your Expedia office and silently logs the day to your Google Sheet. Nothing to do on normal days.
Google Sheet — stores all your data: office visits, manual corrections, and upcoming travel.
Google Apps Script — runs every Sunday at 9am PT, reads the sheet, calculates compliance, and emails the report.

Everything runs in Google's cloud. Once set up, it's fully automatic.


## The Compliance Math
Expedia's RTO policy uses a rolling window — not a simple monthly count:

Look at the last 12 completed work weeks (Mon–Fri only)
Take the best 8 of those 12 weeks
Sum the office days in those 8 weeks
Target: 24 days

A bad week here or there doesn't hurt you — only your best weeks count. RTO Ninja's recommendation accounts for how the window will shift next week, not just where you stand today.


## What the Email Shows You
Every Sunday morning you'll receive:

Compliance total — your current best-8-of-12 score vs. the 24-day target, with a ✅ or ⚠️
Next week's recommendation — the minimum days to go in next week, in plain English
12-week breakdown — a table of every week in the window, with the counted weeks highlighted
Vacation forecast (when applicable) — if a trip is within 5 weeks, the minimum days per week to work beforehand so you're still compliant when you return


## Setup
1. Google Sheet
Create a new Google Sheet with two tabs:

Tab 1: Sheet1 — attendance log

Column
Header
Format
Notes
A
Start Date
yyyy-MM-dd
Required
B
End Date
yyyy-MM-dd
Same as Start for single days
C
Vacation
Yes or blank
Rows marked Yes are excluded from counts
D
Manual Entry
1 or blank
Must be 1 to count. Tasker writes this automatically.


Tab 2: vacations — upcoming trips

Column
Header
Format
Notes
A
Start Date
yyyy-MM-dd
First day of travel
B
End Date
yyyy-MM-dd
Last day of travel
C
Trip Name
Text
Label shown in the email (e.g. "Greece Trip")

2. Google Apps Script
Open script.google.com and create a new project
Paste the contents of rto_ninja.gs into the editor
Fill in your config values at the top of the file:

var SHEET_ID   = "YOUR_GOOGLE_SHEET_ID";   // from your sheet's URL

var EMAIL_TO   = "YOUR_EMAIL_ADDRESS";

In Project Settings, set the timezone to America/Los_Angeles
Run setupTrigger() once to create the Sunday 9am schedule
Run testReport() to send a test email and verify everything works
Note the Web App URL from Deploy → Manage deployments — you'll need it for Tasker

3. Tasker (Android)
Install Tasker ($3 one-time) and create a profile:

Trigger: Location — your Expedia office address (enter event)
Condition: Day of week = Mon/Tue/Wed/Thu/Fri
Action: HTTP GET → your Apps Script Web App URL

The script also blocks weekends and deduplicates server-side, so double-fires won't cause duplicate entries.


## Manual Entries
If Tasker missed a day (phone died, left early, etc.), add a row to Sheet1:

Start Date = End Date = the date in yyyy-MM-dd
Vacation = blank
Manual Entry = 1

For multi-day stretches (e.g. an all-hands or conference), use a date range in columns A and B — the script counts each weekday in the range.


## Why Tasker?
The core challenge is detecting office arrival automatically without any action required on normal days. Here's what we considered:

IFTTT — has location triggers and URL actions, but geofencing reliability is poor (late fires, duplicates, missed arrivals). Now also requires a paid subscription for multiple applets.

MacroDroid — free Tasker alternative, easier to configure, but location trigger reliability is lower and the free tier caps the number of macros.

Microsoft Power Automate — can trigger on location, but requires signing in with your work Microsoft account, adding a dependency on IT policies and raising questions about data flowing through a work account.

Manual logging — rejected immediately. The whole point of RTO Ninja is zero effort on normal days.

Badge swipe data — the ideal source, but Expedia doesn't expose an employee-facing API for building access data.

Tasker won because it's the most reliable Android automation tool available. Its geofencing combines GPS, WiFi, and cell towers for consistent detection, it makes HTTP requests natively, it runs entirely on-device with no external service dependency, and the $3 one-time cost means no subscription to manage.

iPhone users: Tasker is Android-only. You can replicate the location trigger using iOS Shortcuts, which supports URL actions — setup is slightly different but the rest of the project is identical.


## Contributing
This project was built for personal use at Expedia Group Seattle. If you're adapting it for a different office or a different company's RTO policy, the key config values to change are at the top of rto_ninja.gs:

var TARGET_DAYS        = 24;   // your company's day target

var WEEKS_COUNTED      = 8;    // how many weeks count toward the target

var WEEKS_WINDOW       = 12;   // rolling window size in weeks

var VACATION_LOOKAHEAD = 35;   // days ahead to start showing trip forecasts


#### License
MIT — free to use, adapt, and share.

