import smtplib
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
import tkinter as tk
from tkinter import filedialog
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

def send_email(sender, password, smtp_server, smtp_port, to, subject, body, pdf_file):
    message = MIMEMultipart()
    message['From'] = sender
    message['To'] = to
    message['Subject'] = subject
    message.attach(MIMEText(body, 'plain'))

    with open(pdf_file, 'rb') as f:
        pdf = f.read()
        part = MIMEApplication(pdf, Name=pdf_file)
        part['Content-Disposition'] = f'attachment; filename="{pdf_file}"'
        message.attach(part)

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.login(sender, password)
            server.sendmail(sender, to, message.as_string())
        return f'Email sent to {to}'
    except Exception as e:
        return f'Error sending email to {to}: {e}'


def generate_pdf(html_file, pdf_file):
    pdfkit.from_file(html_file, pdf_file, options={"enable-local-file-access": ""})


def process_emails(sender, password, excel_file, html_template):
    df = pd.read_excel(excel_file)
    env = Environment(loader=FileSystemLoader('.'))
    env = Environment(loader=FileSystemLoader(os.path.dirname(html_file_path)))
    template = env.get_template(os.path.basename(html_file_path))

    results = []
    for index, row in df.iterrows():
        html_file = template.render(
            name=row['Name'],
            date=row['DateDonated'].strftime('%m/%d/%Y'),
            amount=row['Amount'],
            check_number=row['Check Number'],
            receipt_no=row['receipt_no'])

        pdf_file = f'{row["EmailAddress"]}.pdf'
        generate_pdf(html_file, pdf_file)

        to = row['EmailAddress']
        subject = 'Receipt for your payment'
        body = f'Dear {row["Name"]},\n\nThank you for your payment of ${row["Amount"]:.2f} on {row["DateDonated"].strftime("%m/%d/%Y")}.\n\nPlease find attached a receipt for your records.\n\nBest regards,\n\nThe XYZ Organization'
        result = send_email(sender, password, 'smtp.gmail.com', 587, to, subject, body, pdf_file)
        results.append(result)

    return results


def browse_excel_file():
    global excel_file_path
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
    excel_file_label.config(text=excel_file_path)

def browse_html_file():
    global html_file_path
    html_file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html *.htm"), ("All files", "*.*")])
    html_file_label.config(text=html_file_path)


def send_emails():
    email = email_entry.get()
    password = password_entry.get()

    if not email or not password or not excel_file_path or not html_file_path:
        log.insert(tk.END, "Please fill all fields and select files.\n")
        return

    results = process_emails(email, password, excel_file_path, html_file_path)
    for result in results:
        log.insert(tk.END, f'{result}\n')

# Create the main window
root = tk.Tk()
root.title('Email Sender')

# Create input fields and labels
email_label = tk.Label(root, text='Email:')
email_label.grid(row=0, column=0)
email_entry = tk.Entry(root)
email_entry.grid(row=0, column=1)

password_label = tk.Label(root, text='Password:')
password_label.grid(row=1, column=0)
password_entry = tk.Entry(root, show='*')
password_entry.grid(row=1, column=1)

excel_file_label = tk.Label(root, text='Excel file:')
excel_file_label.grid(row=2, column=0)
browse_excel_button = tk.Button(root, text='Browse', command=browse_excel_file)
browse_excel_button.grid(row=2, column=1)

html_file_label = tk.Label(root, text='HTML file:')
html_file_label.grid(row=3, column=0)
browse_html_button = tk.Button(root, text='Browse', command=browse_html_file)
browse_html_button.grid(row=3, column=1)

# Create a send button
send_button = tk.Button(root, text='Send Emails', command=send_emails)
send_button.grid(row=4, columnspan=2)

# Create a status/error log display
log_label = tk.Label(root, text='Status/Error Log:')
log_label.grid(row=5, columnspan=2)
log = tk.Text(root, wrap=tk.WORD, width=40, height=10)
log.grid(row=6, columnspan=2)

# Initialize global variables for file paths
excel_file_path = None
html_file_path = None

# Run the main loop
root.mainloop()