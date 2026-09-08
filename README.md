# 📊 Automated Daily GMB Rating & Review Tracker

Automatically tracks Google Business Profile (GMB) **Live Ratings** and **Review Counts** for all store outlets every day at **11:00 AM IST**, updates the Excel sheet with today's date, and emails the updated workbook directly to your inbox.

---

## 📁 What Each File Does

| File | Description |
| :--- | :--- |
| **`scraper.py`** | Launches headless Chrome (Playwright), visits each store link, extracts rating & reviews, appends today's date column to Excel, and saves. |
| **`send_email.py`** | Sends an email with an HTML summary table and attaches the updated Excel workbook. |
| **`.github/workflows/daily_gmb_update.yml`** | GitHub Actions schedule that runs automatically at **11:00 AM IST** every day (even when your PC is turned off). |
| **`Outlet name and Link.xlsx`** | The master Excel workbook with all 24 stores and daily columns. |
| **`requirements.txt`** | Python dependencies (`playwright`, `openpyxl`, `python-dotenv`). |
| **`.env.example`** | Example environment variables for local testing. |

---

## 🚀 Quick Setup Guide (One-Time)

### Step 1: Create a GitHub Repository

1. Go to [github.com](https://github.com) and sign in.
2. Click the **+** icon in the top right > **New repository**.
3. Name it: `gmb-outlet-tracker` (or any name you like).
4. Choose **Private** (recommended, since it contains store business data).
5. Click **Create repository**.

---

### Step 2: Upload Files to the Repository

You can upload directly through your browser:
1. In your new repository on GitHub, click **Upload an existing file** (or drag and drop).
2. Drag and drop all the files from this folder:
   - `.github/` folder (with `workflows/daily_gmb_update.yml`)
   - `scraper.py`
   - `send_email.py`
   - `Outlet name and Link.xlsx`
   - `requirements.txt`
   - `.gitignore`
   - `README.md`
3. Click **Commit changes**.

---

### Step 3: Get a Google App Password (for Email Sending)

To allow GitHub to send emails to your inbox from Gmail:
1. Go to your **Google Account** settings: [myaccount.google.com/security](https://myaccount.google.com/security).
2. Ensure **2-Step Verification** is turned **ON**.
3. Search for **App passwords** in the top search bar (or visit [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)).
4. Enter an App name (e.g. `GMB Tracker`) and click **Create**.
5. Copy the 16-character password generated (it looks like: `abcd efgh ijkl mnop`).

---

### Step 4: Add Secrets in GitHub

In your GitHub repository:
1. Go to **Settings** (tab at the top).
2. In the left menu, expand **Secrets and variables** > click **Actions**.
3. Click **New repository secret** for each of the following 3 secrets:

| Secret Name | Secret Value | Example |
| :--- | :--- | :--- |
| **`EMAIL_SENDER`** | The Gmail address sending the report | `your-email@gmail.com` |
| **`EMAIL_PASSWORD`** | The 16-character App Password from Step 3 | `abcdefghijklmnop` |
| **`EMAIL_RECEIVER`** | Where to receive the report (can be same or multiple separated by comma) | `your-email@gmail.com` |

---

### Step 5: Test It Right Now!

You don't need to wait for 11:00 AM to test it:
1. Go to the **Actions** tab in your GitHub repository.
2. Click **Daily GMB Rating & Review Tracker** in the left sidebar.
3. Click the **Run workflow** dropdown on the right > click the green **Run workflow** button.
4. Watch the run complete in ~1–2 minutes.
5. Check your inbox for the email report with the attached Excel file!

---

## ⏰ Daily Schedule

- **Cron Time:** `30 5 * * *` (05:30 UTC = **11:00 AM IST**).
- GitHub's cloud servers will wake up automatically every single day at 11:00 AM IST, scrape all 24 stores, update `Outlet name and Link.xlsx`, push the changes back to GitHub, and email you the report.
- Works 365 days a year without needing your computer turned on.
