import time
import pandas as pd
import openpyxl
import smtplib
from dotenv import load_dotenv
import os
from ping3 import ping
from datetime import datetime
from email.message import EmailMessage

# Load environment variables from .env
load_dotenv()

# Get email credentials from environment variables
EMAIL_ADDRESS = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASS")

# List of websites to ping
WEBSITES = [
    "google.com",
    "amazon.com",
    "github.com",
    "openai.com",
    "youtube.com",
    "wikipedia.com",
    "linkedin.com",
    "zoom.com"
]

# Output file name
OUTPUT_FILE = "ping_log.xlsx"

# Validate website List
for site in WEBSITES:
    if not isinstance(site, str) or "." not in site:
        print(f"⚠ Warning: Invalid website '{site}' - Skipping...")
        WEBSITES.remove(site)


def ping_website(url):
    """Ping a website and return its response time in milliseconds."""
    try:
        response_time = ping(url, timeout=2)  # Get response time in seconds
        if response_time is not None:
            return round(response_time * 1000, 2)  # Convert to milliseconds
        else:
            return "failed"  # If no response
    except Exception as e:
        return f"Error: {e}"


def log_results(results):
    """save the results to an Excel file(appending data instead of overwriting)."""
    df = pd.DataFrame(results, columns=["Timestamp", "Website", "Response Time (ms)"])

    try:
        existing_df = pd.read_excel(OUTPUT_FILE, engine = "openpyxl")
        df = pd.concat([existing_df, df], ignore_index=True)
        print(f"Results saved to {OUTPUT_FILE}")
    except FileNotFoundError:
        pass  # If file doesn't exist, create a new one

    # Save to Excel
    try:
        df.to_excel(OUTPUT_FILE, index=False, engine="openpyxl")
        print(f"Results saved to {OUTPUT_FILE}")
    except Exception as e:
        print(f"Error saving file: {e}")


def send_email():

    #Load environment variables from .env
    load_dotenv()

    """Send the log file via email."""
    if not os.path.exists(OUTPUT_FILE):
        print("Log file not found, skipping email.")
        return

    msg = EmailMessage()
    msg["Subject"] = "Ping Log Report"
    msg["From"] = EMAIL_ADDRESS
    msg["To"] = EMAIL_ADDRESS
    msg.set_content("Attached is the latest ping log report.")

    #Attach the Excel file
    try:
        with open(OUTPUT_FILE, "rb") as f:
            file_data = f.read()
            file_name = os.path.basename(OUTPUT_FILE)
            msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)


        #connect to SMTP server and send the email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
            server.send_message(msg)

        print("Email sent successfully!")


    except smtplib.SMTPAuthenticationError:
        print("❌ Error: Invalid email credentials. Check your email & app password.")

    except smtplib.SMTPConnectError:
        print("❌ Error: Unable to connect to the email server.")

    except FileNotFoundError:
        print(f"❌ Error: Log file '{OUTPUT_FILE}' not found.")

    except Exception as e:
        print(f"❌ Error sending email: {e}")


def main():
    """Main function to ping websites and log results."""
    results = []

    for website in WEBSITES:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")  # Current timestamp
        response_time = ping_website(website)
        results.append([timestamp, website, response_time])

        print(f"{timestamp} | {website} | {response_time} ms", flush=True)
        time.sleep(0.5)  #small delay for readability


    log_results(results)  # Log results to an Excel file
    send_email()  # Send email after logging results

if __name__ == "__main__":
    main()