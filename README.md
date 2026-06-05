# 📈 MF Tracker — Mutual Fund SIP Tracker

A **Next.js** app to track your monthly SIP investments across mutual funds.

## ✨ Key Features

- **Excel-based storage** — your data lives in `mf-tracker.xlsx`, not on any server
- **No accounts, no cookies, no database** — 100% private
- **Add funds** with expected CAGR, SIP amount, category
- **Log monthly payments** with optional NAV and notes
- **Projection calculator** — see what your SIPs grow to in 5/10/20/30 years
- **Visual charts** — portfolio allocation, monthly investments
- **Full CRUD** — add, edit, delete funds and payments

## 🔄 How It Works

```
Start fresh or upload mf-tracker.xlsx
         ↓
Add funds & log monthly payments
         ↓
Click "Save Excel" to download updated file
         ↓
Next time → Upload the file to continue
```

## 🚀 Setup

```bash
cd mf-tracker
npm install
npm run dev
```

Open [http://localhost:3000](http://localhost:3000)

## 📁 Excel File Structure

The app reads/writes `mf-tracker.xlsx` with 3 sheets:

| Sheet | Contents |
|-------|----------|
| **Funds** | Fund name, category, SIP amount, start date, expected return |
| **Payments** | Monthly payment log with fund, date, amount, NAV |
| **Summary** | Auto-generated overview (read-only) |

You can also open the Excel file in Excel/Google Sheets to view or edit directly.

## 💡 Tips

- **Always download** after adding data — browser memory resets on refresh
- Use the **Projection** button on Dashboard to forecast future value
- The Excel file is human-readable — open it in Excel anytime
- Back up your `mf-tracker.xlsx` to Google Drive / Dropbox for safety
