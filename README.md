# 💰 Smart Expense Tracker

A free, automated expense tracker built on **Google Sheets + Google Apps Script**.
No paid tools, no servers, no subscriptions. It runs inside your own Google account.

Built by **Pablo Chomcobar**.

---

## ✨ Features

| Feature | What it does |
|---|---|
| ⚙️ One-click setup | Creates the **Expenses**, **Settings** and **Dashboard** sheets with formatting and dropdowns |
| ➕ Sidebar entry | Add an expense inside the sheet, with quick-amount buttons (₹50, ₹100, ₹500…) |
| 📱 Mobile form | Creates a Google Form so you can add expenses from your phone in 5 seconds |
| 📊 Live dashboard | Today, this month, budget left, % used, category pie chart, last-6-months bar chart |
| ⚠️ Budget alerts | Emails you when you cross your alert % (default 80%) and when you cross 100% |
| 📧 Monthly report | Emails a full report on the 1st of every month (by category, by payment mode, biggest expense) |
| 🇮🇳 Indian format | ₹ currency with Indian number grouping (₹1,23,456.00), India time zone |

---

## 🚀 Setup (10 minutes)

### 1. Create the Google Sheet
1. Go to [sheets.new](https://sheets.new). A new blank sheet opens.
2. Rename it to **Smart Expense Tracker**.

### 2. Add the code
1. In the sheet, click **Extensions → Apps Script**.
2. Delete everything in `Code.gs` and paste the full contents of `src/Code.gs`.
3. Click **+** next to *Files* → **HTML** → name it exactly `AddExpense` (no `.html`).
   Paste the contents of `src/AddExpense.html`.
4. *(Optional but recommended)* Click ⚙️ **Project Settings** → tick
   **"Show appsscript.json manifest file in editor"**. Open `appsscript.json` and paste
   the contents of `src/appsscript.json`.
5. Click 💾 **Save** (Ctrl + S).

### 3. Run setup
1. Go back to the sheet and **refresh the page**.
2. A new menu appears: **💰 Expense Tracker**.
3. Click **💰 Expense Tracker → ⚙️ Set up tracker (run once)**.
4. Google asks for permission the first time:
   - Click **Continue**, then choose your account.
   - You'll see *"Google hasn't verified this app"*. This is normal for your own scripts.
     Click **Advanced → Go to Smart Expense Tracker (unsafe) → Allow**.
5. Run **Set up tracker** once more if it stopped at the permission screen.

### 4. Personalise
Open the **Settings** sheet:
- `E2`: your monthly budget
- `E3`: email for reports and alerts
- `E4`: the budget % at which you get a warning
- Columns A and B: edit categories and payment modes. The dropdowns update automatically.

### 5. Turn on the extras
- **💰 Expense Tracker → 📱 Create mobile entry form** gives you a link to open on your phone.
  Use "Add to Home screen" for one-tap entry.
- **💰 Expense Tracker → ⏰ Turn on automations** turns on the daily budget check and the monthly report email.
- **💰 Expense Tracker → 📧 Email me this month's report now** sends a test report.

---

## 🗂️ Project structure

```
smart-expense-tracker/
├── README.md
└── src/
    ├── Code.gs            # All logic: setup, dashboard, form, alerts, reports
    ├── AddExpense.html    # Sidebar UI for adding expenses
    └── appsscript.json    # Manifest (time zone + V8 runtime)
```

## 🧰 Tech
Google Sheets · Google Apps Script (V8) · Google Forms · MailApp · Time-driven triggers

## 📝 Free limits
Free Gmail accounts can send about **100 emails per day** from Apps Script. This tracker
sends at most a few per month, so you'll never hit the limit.

---

© Pablo Chomcobar
