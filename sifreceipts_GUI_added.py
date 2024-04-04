from email.mime.application import MIMEApplication
import smtplib
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit
from tkinter import *
from tkinter import filedialog
import os
import sys

def append_log(text):
    log_text.insert(END, text + '\n')
    log_text.see(END)
    app.update()

class OutputRedirector:
    def __init__(self, func):
        self.func = func

    def write(self, text):
        self.func(text)

    def flush(self):
        pass

sys.stdout = OutputRedirector(append_log)
sys.stderr = OutputRedirector(append_log)


def open_file():
    file_path = filedialog.askopenfilename()
    return file_path

def process_inputs():
    email = email_entry.get()
    password = password_entry.get()

    if not html_file.get():
        html_label.config(text="Select HTML file")
        html_file.set(open_file())
        html_label.config(text=f"Selected: {html_file.get()}")
    elif not excel_file.get():
        excel_label.config(text="Select Excel file")
        excel_file.set(open_file())
        excel_label.config(text=f"Selected: {excel_file.get()}")
    else:
        # Call your main function here with the obtained inputs
        main(email, password, html_file.get(), excel_file.get())

def send_email(sender, password, to, subject, body, pdf_file):
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    
    message = f"""From: {sender}
To: {to}
Subject: {subject}

{body}
"""
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.ehlo()
            server.starttls()
            server.login(sender, password)
            with open(pdf_file, 'rb') as f:
                pdf = f.read()
                message = message.encode('utf-8')
                part = MIMEApplication(pdf, Name=pdf_file)
                part['Content-Disposition'] = f'attachment; filename="{pdf_file}"'
                message.attach(part)
                server.sendmail(sender, to, message.as_string())
        append_log(f'Email sent to {to}')
    except Exception as e:
        append_log(f'Error sending email to {to}: {e}')

def generate_pdf(html_file, pdf_file):
    pdfkit.from_file(html_file, pdf_file, options={"enable-local-file-access":""})

def main(sender, password, html_file, excel_file):
    df = pd.read_excel(excel_file)  
    # Configure Jinja2 environment
    # Get the directory of the provided HTML file
    html_file_dir = os.path.dirname(os.path.abspath(html_file))

    # Configure Jinja2 environment
    env = Environment(loader=FileSystemLoader(html_file_dir))
    template = env.get_template(os.path.basename(html_file))

    for index, row in df.iterrows():
        # Render the HTML template with data from the spreadsheet
        htmlFile = template.render(
            name=row['Name'],
            date=row['DateDonated'].strftime('%m/%d/%Y'),
            amount=row['Amount'],
            check_number=row['Check Number'],
            receipt_no=row['receipt_no'])
        with open("NewReceipthtml.html", 'w') as fh:
            fh.write(htmlFile)
        
        # Generate a unique PDF file name based on the email address
        pdf_file = f'{row["EmailAddress"]}.pdf'
        html = new_func()
        # Generate the PDF file from the HTML template
        generate_pdf(html, pdf_file)
        
        # Send the email with the PDF attachment
        to = row['EmailAddress']
        subject = 'Receipt for your payment'
        body = f'Dear {row["Name"]},\n\nThank you for your payment of ${row["Amount"]:.2f} on {row["DateDonated"].strftime("%m/%d/%Y")}.\n\nPlease find attached a receipt for your records.\n\nBest regards,\n\nThe XYZ Organization'
        send_email(sender, password, to, subject, body, pdf_file)
        

def new_func():
    return 'NewReceipthtml.html'

app = Tk()
app.title("Email Receipts")

email_label = Label(app, text="Email:")
email_label.grid(row=0, column=0)
email_entry = Entry(app)
email_entry.grid(row=0, column=1)

password_label = Label(app, text="Password:")
password_label.grid(row=1, column=0)
password_entry = Entry(app, show="*")
password_entry.grid(row=1, column=1)

html_label = Label(app, text="HTML file not selected")
html_label.grid(row=2, column=0, columnspan=2)
html_button = Button(app, text="Select HTML file", command=process_inputs)
html_button.grid(row=3, column=0, columnspan=2)

excel_label = Label(app, text="Excel file not selected")
excel_label.grid(row=4, column=0, columnspan=2)
excel_button = Button(app, text="Select Excel file", command=process_inputs)
excel_button.grid(row=5, column=0, columnspan=2)

start_button = Button(app, text="Start", command=process_inputs)
start_button.grid(row=6, column=0, columnspan=2)

html_file = StringVar()
excel_file = StringVar()

log_label = Label(app, text="Logs:")
log_label.grid(row=7, column=0, columnspan=2)

log_text = Text(app, width=50, height=10, wrap=WORD)
log_text.grid(row=8, column=0, columnspan=2)

scrollbar = Scrollbar(app, command=log_text.yview)
scrollbar.grid(row=8, column=2, sticky=N+S)
log_text.config(yscrollcommand=scrollbar.set)

app.mainloop()
# In this code, the receipt_data.xlsx file should contain a spreadsheet with columns for Email, Name, Date, Amount, and Check Number. 
# The main() function loads the data from the spreadsheet into a Pandas DataFrame, configures a Jinja2 environment, and iterates over each row in the DataFrame. 
#  For each row, it renders the HTML template with data from the spreadsheet, generates a unique PDF file name based on the email address, generates the PDF file from the HTML template using pdfkit, and sends the email with the PDF attachment using the send_email() function. 
#  The body variable is a string that contains placeholders for the actual name, amount, and date values. 
# These values can be replaced with real data using string formatting, such as `f'Thank you for your payment of ${row["Amount"]:.2f} on {row

