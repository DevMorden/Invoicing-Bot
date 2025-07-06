import gspread
from oauth2client.service_account import ServiceAccountCredentials
from docx import Document
from datetime import datetime
import yagmail
import os

INVOICE_TEMPLATE_PATH = 'invoice_template.docx'
INVOICE_OUTPUT_PATH = 'invoice_output.docx'
INVOICE_TRACKER_PATH = 'invoice_number.txt'

def get_next_invoice_number():
    if not os.path.exists(INVOICE_TRACKER_PATH):
        with open(INVOICE_TRACKER_PATH, 'w') as f:
            f.write('1')
            return 1
    with open(INVOICE_TRACKER_PATH, 'r+') as f:
        num = int(f.read().strip())
        f.seek(0)
        f.write(str(num + 1))
        f.truncate()
    return num

def get_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    client = gspread.authorize(creds)
    return client.open(sheet_name).sheet1

def find_email_by_name(sheet, name):
    records = sheet.get_all_records()
    for row in records:
        if row['Name'].strip().lower() == name.strip().lower():
            return row['Email']
    return None

def generate_invoice(client_name, amount, invoice_number):
    doc = Document(INVOICE_TEMPLATE_PATH)
    for p in doc.paragraphs:
        p.text = p.text.replace('{client_name}', client_name)
        p.text = p.text.replace('{amount}', f"${amount:.2f}")
        p.text = p.text.replace('{invoice_number}', str(invoice_number))
        p.text = p.text.replace('{date}', datetime.now().strftime('%Y-%m-%d'))
    doc.save(INVOICE_OUTPUT_PATH)

def send_email(recipient_email, subject, body, attachment_path):
    yag = yagmail.SMTP("your_email@gmail.com", "your_app_password")  # Use App Password if 2FA is on
    yag.send(to=recipient_email, subject=subject, contents=body, attachments=attachment_path)

def create_and_send_invoice(client_name, amount, custom_message):
    sheet = get_sheet("ClientList")
    email = find_email_by_name(sheet, client_name)
    if not email:
        print("Client not found.")
        return
    invoice_number = get_next_invoice_number()
    generate_invoice(client_name, amount, invoice_number)
    send_email(
        email,
        f"Invoice #{invoice_number} from Lawn Maintenance",
        custom_message,
        INVOICE_OUTPUT_PATH
    )
    print(f"✅ Invoice #{invoice_number} sent to {email}.")

# Example usage:
# create_and_send_invoice("John Smith", 120.00, "Hi John, here's your invoice for this week's service. Thanks as always!")
