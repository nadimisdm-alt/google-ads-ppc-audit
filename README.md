# 📊 Google Ads PPC Audit Script

A free, open-source Google Ads Script that automatically pulls key performance insights into a beautifully formatted Google Sheets dashboard — in one click.

Built for PPC managers, agency teams, and marketers who want fast, actionable audits without manual reporting.

---

## 🚀 What It Does

Runs a full account audit and generates a multi-tab Google Sheets report covering:

| Tab | Description |
|-----|-------------|
| 🏠 Overview | Account KPIs (30-day snapshot) + Week-over-Week summary + Health Scorecard |
| 💀 Zombie Campaigns | Campaigns spending budget with zero conversions in the last 14 days |
| ❌ Non-Converting Terms | Search terms burning budget without a single conversion |
| 💰 High CPC Terms | Search terms above your CPC threshold with no conversions |
| 🛑 Negative KW Candidates | Suggested terms to add as negatives to save budget |
| 📈 Week-over-Week | Campaign-level performance comparison: last 7 days vs prior 7 days |

---

## ⚙️ Setup Instructions

### Step 1 — Prepare Your Google Sheet
1. Create a new Google Sheet
2. Copy its URL from the browser address bar

### Step 2 — Add the Script to Google Ads
1. In Google Ads, go to **Tools & Settings > Bulk Actions > Scripts**
2. Click the **+** button to create a new script
3. Delete any default code and paste the contents of `google_ads_ppc_audit_fixed.js`

### Step 3 — Configure the Script
Edit the `CONFIG` block at the top of the script:

```javascript
var CONFIG = {
  SPREADSHEET_URL: 'YOUR_GOOGLE_SHEET_URL_HERE', // Paste your sheet URL
  HIGH_CPC_THRESHOLD: 3.00,                       // Flag terms above this CPC (no conversions)
  MIN_IMPRESSIONS_FOR_AUDIT: 10,                  // Minimum impressions to include a term
  MIN_ZOMBIE_SPEND: 1.00,                         // Minimum spend to flag a zombie campaign
};
```

### Step 4 — Authorize and Run
1. Click **Authorize** and grant the required permissions
2. Click **Run** — the sheet will be populated and auto-renamed to:
   `[Account Name] — Google Ads PPC Audit — YYYY-MM-DD`

### Step 5 — Schedule (Optional)
Set the script to run weekly under **Frequency** in the Scripts dashboard so your audit stays up to date automatically.

---

## 📋 Requirements

- A Google Ads account (Manager or direct account access)
- A blank Google Sheet shared with your Google Ads account email
- No coding experience needed

---

## 🛠️ Configuration Options

| Setting | Default | Description |
|---------|---------|-------------|
| `SPREADSHEET_URL` | *(required)* | URL of your Google Sheet |
| `HIGH_CPC_THRESHOLD` | `3.00` | CPC value above which a term is flagged (in your account currency) |
| `MIN_IMPRESSIONS_FOR_AUDIT` | `10` | Minimum impressions required for a search term to appear in audit tabs |
| `MIN_ZOMBIE_SPEND` | `1.00` | Minimum spend for a campaign to be flagged as a zombie |

---

## 🐛 Bug Fixes (v2)

### CTR Calculation Fix
In the original version, CTR values were displaying incorrectly (e.g., showing `169` instead of `0.17%`).

**Root cause:** The Google Ads API returns the `Ctr` field as a decimal fraction (e.g., `0.0017` = 0.17%). The script was multiplying by 100 in multiple places across three functions, causing the value to inflate.

**Fixed in:**
- `getAccountTotals()` — Overview KPI cards
- `runNonConvertingSearchTerms()` — CTR column in Non-Converting Terms tab
- `fetchCampaignData()` inside `runWeekOverWeek()` — CTR columns for This Week and Last Week

Thanks to **Gopinath Govindaraj** and **David Hill** for catching and reporting this issue.

---

## 📸 Report Preview

The generated report includes:
- Color-coded KPI cards for quick health checks
- Green/red delta highlighting for Week-over-Week changes
- A Health Scorecard with an overall pass/fail score
- Sortable data tables across all audit tabs

---

## 🤝 Contributing

Found a bug or have a feature request? Feel free to:
- Open an **Issue**
- Submit a **Pull Request**
- Leave a comment with feedback

All contributions are welcome.

---

## 📄 License

MIT License — free to use, modify, and distribute.

---

## 👤 Author

**Nadim Mahmud Sizan**
Helping brands turn Google Ads from a black box into a glass box.

> If this script saved you time, consider sharing it with your network or leaving a ⭐ on the repo!
