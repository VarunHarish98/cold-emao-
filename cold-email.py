# # import pandas as pd
# # import smtplib
# # from email.message import EmailMessage
# # import os

# # # Configuration
# # EXCEL_FILE = 'emails.xlsx'  # Ensure this file exists
# # PDF_ATTACHMENT = 'attachment.pdf'  # Ensure this file exists
# # SMTP_SERVER = 'smtp.gmail.com'
# # SMTP_PORT = 587

# # # Your Gmail credentials (Use App Password, not your regular password)
# # GMAIL_USER = 'varunharish98@gmail.com'
# # GMAIL_PASS = 'mdieoyujmpsfahyh'  # Replace with App Password

# # def send_email(to_email, company_name, mail_body, attachment_path):
# #     """Sends an email with a PDF attachment."""
# #     msg = EmailMessage()
# #     msg['Subject'] = f"Message for {company_name}" if company_name else "Message"
# #     msg['From'] = GMAIL_USER
# #     msg['To'] = to_email
# #     msg.set_content(mail_body)

# #     # Check if the attachment exists
# #     if os.path.exists(attachment_path):
# #         with open(attachment_path, 'rb') as f:
# #             file_data = f.read()
# #             file_name = os.path.basename(attachment_path)
# #             msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)
# #     else:
# #         print(f"Warning: Attachment {attachment_path} not found. Sending email without attachment.")

# #     # Connect to SMTP server and send email
# #     try:
# #         with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
# #             server.starttls()  # Secure connection
# #             server.login(GMAIL_USER, GMAIL_PASS)
# #             server.send_message(msg)
# #         print(f"✅ Email sent successfully to {to_email}")
# #         return "Sent"
# #     except Exception as e:
# #         print(f"❌ Failed to send email to {to_email}. Error: {e}")
# #         return f"Failed: {e}"

# # def main():
# #     """Reads the Excel file and sends emails individually."""
# #     if not os.path.exists(EXCEL_FILE):
# #         print(f"Error: {EXCEL_FILE} not found!")
# #         return

# #     try:
# #         df = pd.read_excel(EXCEL_FILE)
# #     except Exception as e:
# #         print(f"Error reading Excel file: {e}")
# #         return

# #     # Ensure "Status" column exists
# #     if 'Status' not in df.columns:
# #         df['Status'] = ""

# #     for index, row in df.iterrows():
# #         company_name = row.get('Company', 'Valued Partner')
# #         to_email = row.get('Email')
# #         mail_body = row.get('MailBody', '')

# #         # Validate email and mail body
# #         if pd.isna(to_email) or pd.isna(mail_body):
# #             df.at[index, 'Status'] = "Skipped (Missing Data)"
# #             continue

# #         # Personalize mail body
# #         personalized_body = mail_body.replace("{company}", company_name) if company_name else mail_body

# #         # Send email and update status
# #         status = send_email(to_email, company_name, personalized_body, PDF_ATTACHMENT)
# #         df.at[index, 'Status'] = status

# #     # Save updated Excel file
# #     df.to_excel(EXCEL_FILE, index=False)
# #     print("✅ Emails sent and status updated in Excel.")

# # if __name__ == '__main__':
# #     main()

# import pandas as pd
# import smtplib
# import imaplib
# import email
# from email.message import EmailMessage
# import os
# import time
# from datetime import datetime, timedelta

# # Configuration
# EXCEL_FILE = 'emails.xlsx'  # Excel file with columns: Company, Email, MailBody, FollowUpBody
# PDF_ATTACHMENT = 'attachment.pdf'
# SMTP_SERVER = 'smtp.gmail.com'
# SMTP_PORT = 587
# IMAP_SERVER = 'imap.gmail.com'
# CHECK_INTERVAL = 4  # Days before sending follow-up

# # Your Gmail credentials
# GMAIL_USER = 'varunharish98@gmail.com'
# GMAIL_PASS = 'mdieoyujmpsfahyh'  # Replace with App Password

# def send_email(to_email, company_name, mail_body, attachment_path):
#     msg = EmailMessage()
#     msg['Subject'] = f"Message for {company_name}"
#     msg['From'] = GMAIL_USER
#     msg['To'] = to_email
#     msg.set_content(mail_body)
    
#     # Attach the PDF file
#     with open(attachment_path, 'rb') as f:
#         file_data = f.read()
#         file_name = os.path.basename(attachment_path)
#         msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

#     # Connect to the Gmail SMTP server and send the email
#     try:
#         with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
#             server.starttls()  # Secure the connection
#             server.login(GMAIL_USER, GMAIL_PASS)
#             server.send_message(msg)
#             print(f"Email sent successfully to {to_email}")
#     except Exception as e:
#         print(f"Failed to send email to {to_email}. Error: {e}")

# def check_reply(to_email):
#     try:
#         mail = imaplib.IMAP4_SSL(IMAP_SERVER)
#         mail.login(GMAIL_USER, GMAIL_PASS)
#         mail.select('inbox')
        
#         status, messages = mail.search(None, f'(FROM "{to_email}")')
#         if messages[0]:
#             return True  # Reply received
#         return False
#     except Exception as e:
#         print(f"Error checking email replies: {e}")
#         return False

# def main():
#     try:
#         df = pd.read_excel(EXCEL_FILE)
#     except Exception as e:
#         print(f"Error reading Excel file: {e}")
#         return

#     last_sent = {}
    
#     for index, row in df.iterrows():
#         company_name = row.get('Company')
#         to_email = row.get('Email')
#         mail_body = row.get('MailBody')
#         followup_body = row.get('FollowUpBody')
        
#         if pd.isna(to_email) or pd.isna(mail_body):
#             print(f"Skipping row {index} due to missing Email or MailBody.")
#             continue
        
#         personalized_body = mail_body.replace("{company}", company_name) if company_name else mail_body
#         send_email(to_email, company_name, personalized_body, PDF_ATTACHMENT)
#         last_sent[to_email] = datetime.now()
    
#     # Wait for 4 days before sending follow-ups
#     while True:
#         time.sleep(86400)  # Sleep for a day
#         for index, row in df.iterrows():
#             to_email = row.get('Email')
#             followup_body = row.get('FollowUpBody')
            
#             if to_email in last_sent and datetime.now() >= last_sent[to_email] + timedelta(days=CHECK_INTERVAL):
#                 if not check_reply(to_email):
#                     print(f"Sending follow-up email to {to_email}")
#                     send_email(to_email, "Follow-up", followup_body, PDF_ATTACHMENT)
#                     last_sent[to_email] = datetime.now()

# if __name__ == '__main__':
#     main()

import pandas as pd
import smtplib
import re
import time
import random
from email.message import EmailMessage



# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------
EXCEL_FILE = 'emails.xlsx'         # Your Excel file with required columns
SMTP_SERVER = 'smtp.gmail.com'       # Gmail SMTP server
SMTP_PORT = 587                      # Gmail SMTP port
GMAIL_USER = 'varunharish98@gmail.com'  # Your Gmail address
GMAIL_PASS = 'mdieoyujmpsfahyh'     # Your Gmail App Password
MAX_EMAILS_PER_RUN = 490             # Maximum number of emails to send per run
DELAY_RANGE_SECONDS = (15, 20)        # Random delay between sends (in seconds)
RESUME_FILE_PATH = "Resume_Varun.pdf"
GITHUB = "github.com/VarunHarish98"
LINKEDIN = "linkedin.com/in/varun-harish1998"
PORTFOLIO = "https://varun-portfolio-one.vercel.app/"

# -----------------------------------------------------------------------------
# Unicode Bold Conversion Functions
# -----------------------------------------------------------------------------
def to_unicode_bold(text: str) -> str:
    """Convert a string to its mathematical bold Unicode equivalent."""
    bold_text = ""
    for char in text:
        if 'A' <= char <= 'Z':
            bold_text += chr(ord(char) - ord('A') + 0x1D400)
        elif 'a' <= char <= 'z':
            bold_text += chr(ord(char) - ord('a') + 0x1D41A)
        else:
            bold_text += char
    return bold_text

def apply_unicode_bold(mail_body: str) -> str:
    """
    Convert text wrapped in ** (e.g., **this is bold**) to its Unicode bold equivalent.
    """
    def replacer(match):
        inner_text = match.group(1)
        return to_unicode_bold(inner_text)
    
    # Replace all occurrences of **text** with its Unicode bold version.
    return re.sub(r'\*\*(.*?)\*\*', replacer, mail_body)

# -----------------------------------------------------------------------------
# Extract Correct Company Domain
# -----------------------------------------------------------------------------
def company_to_domain(company: str) -> str:
    """Extract company domain correctly (e.g., vectra.ai instead of vectra.com)."""
    company_clean = re.sub(r'[^a-z0-9.]', '', company.lower())  # Allow dots for domains
    return company_clean if '.' in company_clean else f"{company_clean}.com"

# -----------------------------------------------------------------------------
# Generate All Possible Email Formats
# -----------------------------------------------------------------------------
def generate_all_emails(name: str, company: str) -> list:
    """Generate all possible corporate email formats."""
    parts = name.strip().split()
    if not parts or not company:
        return []

    first = parts[0].lower()
    last = parts[-1].lower()
    first_initial = first[0]

    domain = company_to_domain(company)

    return [
        f"{first}.{last}@{domain}",
        f"{first_initial}{last}@{domain}",
        f"{first}{last[0]}@{domain}",
        f"{first}@{domain}",
        f"{last}{first_initial}@{domain}",
        f"{first}_{last}@{domain}",
        f"{first}-{last}@{domain}",
        f"{first}{last}@{domain}"
    ]

# -----------------------------------------------------------------------------
# Send Email Function
# -----------------------------------------------------------------------------
def send_email_to(recipient: str, company_name: str, mail_body: str, name: str) -> bool:
    """Send an email with an attached resume."""
    # Replace placeholders with actual values.
    personalized_body = mail_body.replace("{company}", company_name if company_name else "")
    personalized_body = personalized_body.replace("{name}", name if name else "")
    personalized_body = personalized_body.replace("{github}", GITHUB if GITHUB else "")
    personalized_body = personalized_body.replace("{linkedin}", LINKEDIN if LINKEDIN else "")
    personalized_body = personalized_body.replace("{portfolio}", PORTFOLIO if PORTFOLIO else "")
    
    # Apply Unicode bold conversion (text within **...**)
    personalized_body = apply_unicode_bold(personalized_body)
    
    if not company_name:
        raise Exception("Company name is missing.")

    msg = EmailMessage()
    subject = f"Exploring Frontend Engineering opportunity at {company_name}" if company_name else "Software Engineering opportunity"
    msg['Subject'] = subject
    msg['From'] = GMAIL_USER
    msg['To'] = recipient
    msg.set_content(personalized_body)

    # Attach the resume file
    try:
        with open(RESUME_FILE_PATH, "rb") as file:
            msg.add_attachment(file.read(), maintype="application", subtype="pdf", filename="Varun_Harish_Resume.pdf")
    except Exception as e:
        print(f"[ERROR] Could not attach resume. Error: {e}")
        return False

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(GMAIL_USER, GMAIL_PASS)
            server.send_message(msg)
        print(f"[SUCCESS] Email sent to {recipient}")
        return True
    except Exception as e:
        print(f"[ERROR] Failed to send email to {recipient}. Error: {e}")
        return False

# -----------------------------------------------------------------------------
# Send Email to All Variants
# -----------------------------------------------------------------------------
def send_email_to_all_variants(email_variants: list, company: str, mail_body: str, name: str) -> bool:
    """Try sending an email to all generated variations."""
    for email in email_variants:
        if send_email_to(email, company, mail_body, name):
            return True  # Stop after the first successful email
    return False

# -----------------------------------------------------------------------------
# Main Execution
# -----------------------------------------------------------------------------
def main():
    try:
        df = pd.read_excel(EXCEL_FILE)
    except Exception as e:
        print(f"[ERROR] Could not read Excel file '{EXCEL_FILE}'. Error: {e}")
        return

    # Ensure necessary columns exist
    for col in ["Company", "Name", "MailBody", "Status"] + [f"Email {i+1}" for i in range(8)]:
        if col not in df.columns:
            df[col] = ""

    emails_sent = 0

    for index, row in df.iterrows():
        # Skip rows where status is already "Sent"
        status = str(row["Status"]).strip().lower()
        if status == "sent":
            continue  # Skip without delay

        # Check if any email column is already populated
        email_columns = [f"Email {i+1}" for i in range(8)]
        existing_emails = [row[email] for email in email_columns if pd.notna(row[email]) and row[email] != ""]  # Collect existing emails

        if not existing_emails:  # If no email exists, generate new emails
            company = str(row["Company"]).strip()
            name = str(row["Name"]).strip()
            mail_body = str(row["MailBody"]).strip()

            if not mail_body:
                print(f"[WARNING] Row {index} has no MailBody. Skipping.")
                continue

            # Generate email variations
            email_variants = generate_all_emails(name, company)
            for i, email in enumerate(email_variants):
                df.at[index, f"Email {i+1}"] = email

            # Send email to all variants
            if send_email_to_all_variants(email_variants, company, mail_body, name):
                df.at[index, "Status"] = "Sent"
                emails_sent += 1
        else:
            # If emails already exist, send to the first available email
            email_to_send = existing_emails[0]  # You can modify this logic if you want to send to more than one
            company = str(row["Company"]).strip()
            name = str(row["Name"]).strip()
            mail_body = str(row["MailBody"]).strip()

            if send_email_to(email_to_send, company, mail_body, name):
                df.at[index, "Status"] = "Sent"
                emails_sent += 1

        if emails_sent >= MAX_EMAILS_PER_RUN:
            print(f"Reached the maximum limit of {MAX_EMAILS_PER_RUN} emails for this run.")
            break

        if emails_sent < MAX_EMAILS_PER_RUN:  # Apply delay only when necessary
            delay = random.randint(*DELAY_RANGE_SECONDS)
            print(f"Sleeping for {delay} seconds before next send...")
            time.sleep(delay)

    # Save updates back to the Excel file
    try:
        df.to_excel(EXCEL_FILE, index=False)
        print(f"[INFO] Updated Excel file saved to '{EXCEL_FILE}'.")
    except Exception as e:
        print(f"[ERROR] Could not save updates to Excel file. Error: {e}")

if __name__ == '__main__':
    main()