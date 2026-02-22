# SS Reporting Tool — Weekly Update Builder

A dark-themed web app for building weekly CS team reports. Fill in metrics, meetings, accomplishments, and sync items; preview live; copy as plain text or Slack-formatted markdown.

## Features

- **Report Period** — Team name and date range
- **CS Router Metrics** — Total conversations, median response time, goal tracking with status indicator
- **Meetings Attended** — Add/remove meetings
- **Accomplishments / Updates** — Recurring presets + one-off items with carry-forward to next week
- **Front Office Sync** — Meeting title and agenda items
- **Private Notes** — Personal context (not in report)
- **Live Preview** — Real-time preview with char/word count
- **Copy** — Plain text or Slack bold format
- **History** — Week navigation and sparkline of conversations per week
- **Auto-save** — Data stored in `localStorage` by week

## Quick Start

### Option 1: Open directly
Open `index.html` in a browser. No server required.

### Option 2: Local dev server
```bash
npm start
# or: npm run dev
```
Then open http://localhost:3000 (or the port shown).

## Project Structure

```
SS Reporting Tool/
├── index.html    # Main HTML
├── styles.css    # All styles
├── app.js        # App logic
├── package.json  # Optional dev scripts
└── README.md     # This file
```

## Data

- Saved in browser `localStorage` under key `fountain_report_v2`
- No backend; data stays local
- Copy to Slack or elsewhere as needed
