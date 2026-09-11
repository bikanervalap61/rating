import os
import smtplib
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

BASE_DIR = Path(__file__).parent.resolve()
load_dotenv(BASE_DIR / ".env")

def get_target_excel_file():
    env_file = os.getenv("EXCEL_FILE")
    if env_file and (BASE_DIR / env_file).exists():
        return BASE_DIR / env_file
    for name in ["Outlet name and Link.xlsx", "Outlet name and Link - Copy.xlsx"]:
        path = BASE_DIR / name
        if path.exists():
            return path
    return None

def build_html_body(data):
    date_str = data.get("date", datetime.now().strftime("%d/%m/%Y"))
    total = data.get("total_stores", 0)
    success = data.get("successful", 0)
    failed = data.get("failed", 0)
    results = data.get("results", [])

    rows_html = ""
    for idx, item in enumerate(results, 1):
        code = item.get("store_code", "")
        address = item.get("address", "")
        rating = item.get("rating")
        reviews = item.get("reviews")
        status = item.get("status", "unknown")

        rating_display = f"{rating} ★" if rating is not None else "<span style='color:#e06666;'>N/A</span>"
        reviews_display = f"{reviews:,}" if reviews is not None else "<span style='color:#e06666;'>N/A</span>"
        
        if status == "success":
            badge_color = "#38761d"
            status_text = "SUCCESS"
        elif status == "partial":
            badge_color = "#b45f06"
            status_text = "PARTIAL"
        else:
            badge_color = "#cc0000"
            status_text = "FAILED"
        status_badge = f"<span style='color:{badge_color}; font-weight:600;'>{status_text}</span>"

        row_bg = "#ffffff" if idx % 2 != 0 else "#f9f9f9"
        rows_html += f"""
        <tr style="background-color: {row_bg}; border-bottom: 1px solid #e0e0e0;">
            <td style="padding: 10px 12px; text-align: center; color: #555;">{idx}</td>
            <td style="padding: 10px 12px; font-weight: bold; color: #333;">{code}</td>
            <td style="padding: 10px 12px; color: #444;">{address}</td>
            <td style="padding: 10px 12px; text-align: center; font-weight: bold; color: #b45f06; background-color: #fff2cc;">{rating_display}</td>
            <td style="padding: 10px 12px; text-align: center; font-weight: bold; color: #134f5c; background-color: #d0e0e3;">{reviews_display}</td>
            <td style="padding: 10px 12px; text-align: center;">{status_badge}</td>
        </tr>
        """

    html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <style>
            body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; line-height: 1.5; margin: 0; padding: 20px; background-color: #f4f5f7; }}
            .container {{ max-width: 820px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 8px rgba(0,0,0,0.08); }}
            .header {{ background: linear-gradient(135deg, #4285F4 0%, #0F9D58 100%); color: white; padding: 24px 30px; }}
            .header h1 {{ margin: 0 0 6px 0; font-size: 22px; font-weight: 700; }}
            .header p {{ margin: 0; opacity: 0.9; font-size: 14px; }}
            .content {{ padding: 24px 30px; }}
            .stats {{ display: flex; gap: 15px; margin-bottom: 24px; }}
            .stat-box {{ flex: 1; padding: 14px 18px; border-radius: 6px; background: #f8f9fa; border: 1px solid #e9ecef; text-align: center; }}
            .stat-val {{ font-size: 22px; font-weight: bold; color: #202124; }}
            .stat-label {{ font-size: 12px; text-transform: uppercase; color: #5f6368; letter-spacing: 0.5px; margin-top: 4px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 13px; }}
            th {{ background-color: #202124; color: white; padding: 11px 12px; text-align: left; font-size: 12px; letter-spacing: 0.3px; }}
            .footer {{ padding: 18px 30px; background: #f8f9fa; border-top: 1px solid #e9ecef; font-size: 12px; color: #70757a; text-align: center; }}
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>📊 Daily GMB Rating & Review Report</h1>
                <p>Date: <strong>{date_str}</strong> | Automated Daily Update at 11:00 AM IST</p>
            </div>
            <div class="content">
                <table style="width: 100%; margin-bottom: 20px; border: none;">
                    <tr>
                        <td style="width: 33%; padding: 12px; background: #e8f0fe; border-radius: 6px; text-align: center;">
                            <div style="font-size: 22px; font-weight: bold; color: #1a73e8;">{total}</div>
                            <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Total Stores</div>
                        </td>
                        <td style="width: 33%; padding: 12px; background: #e6f4ea; border-radius: 6px; text-align: center;">
                            <div style="font-size: 22px; font-weight: bold; color: #137333;">{success}</div>
                            <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Successfully Scraped</div>
                        </td>
                        <td style="width: 33%; padding: 12px; background: #fce8e6; border-radius: 6px; text-align: center;">
                            <div style="font-size: 22px; font-weight: bold; color: #c5221f;">{failed}</div>
                            <div style="font-size: 11px; text-transform: uppercase; color: #5f6368;">Failed / Review Needed</div>
                        </td>
                    </tr>
                </table>

                <p style="font-size: 13px; color: #444; margin-bottom: 12px;">
                    The updated Excel workbook with today's date column (<strong>{date_str.split()[0]}</strong>) has been attached to this email.
                </p>

                <table>
                    <thead>
                        <tr>
                            <th style="text-align: center;">#</th>
                            <th>Store Code</th>
                            <th>Store Address</th>
                            <th style="text-align: center;">Live Rating</th>
                            <th style="text-align: center;">Reviews</th>
                            <th style="text-align: center;">Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        {rows_html}
                    </tbody>
                </table>
            </div>
            <div class="footer">
                This is an automated report generated by the Google Business Profile Daily Tracker via GitHub Actions.
            </div>
        </div>
    </body>
    </html>
    """
    return html

def send_daily_email():
    sender = os.getenv("EMAIL_SENDER")
    password = os.getenv("EMAIL_PASSWORD")
    receivers_raw = os.getenv("EMAIL_RECEIVER")
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))

    if not sender or not password or not receivers_raw:
        print("\n[send_email] EMAIL_SENDER, EMAIL_PASSWORD, or EMAIL_RECEIVER not set.")
        print("[send_email] Skipping email dispatch. (Set these in .env or GitHub Secrets)")
        return

    receivers = [r.strip() for r in receivers_raw.split(",") if r.strip()]
    if not receivers:
        print("[send_email] No valid recipient email addresses found.")
        return

    # Load scraped summary
    summary_file = BASE_DIR / "scraped_results.json"
    data = {}
    if summary_file.exists():
        with open(summary_file, "r", encoding="utf-8") as f:
            data = json.load(f)

    today_str = datetime.now().strftime("%d/%m/%Y")
    subject = f"📊 Daily GMB Ratings & Reviews - {today_str}"

    msg = MIMEMultipart("alternative")
    msg["From"] = f"GMB Daily Tracker <{sender}>"
    msg["To"] = ", ".join(receivers)
    msg["Subject"] = subject

    html_content = build_html_body(data)
    msg.attach(MIMEText(html_content, "html"))

    # Attach Excel file
    excel_file = get_target_excel_file()
    if excel_file and excel_file.exists():
        part = MIMEBase("application", "octet-stream")
        with open(excel_file, "rb") as f:
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename=\"{excel_file.name}\""
        )
        msg.attach(part)
        print(f"[send_email] Attached Excel file: {excel_file.name}")
    else:
        print("[send_email] Warning: Target Excel file not found to attach.")

    print(f"[send_email] Connecting to {smtp_server}:{smtp_port}...")
    try:
        if smtp_port == 465:
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        else:
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()

        server.login(sender, password)
        server.sendmail(sender, receivers, msg.as_string())
        server.quit()
        print(f"[send_email] [SUCCESS] Successfully sent daily report to: {', '.join(receivers)}")
    except Exception as e:
        print(f"[send_email] [ERROR] Failed to send email: {e}")

if __name__ == "__main__":
    send_daily_email()
