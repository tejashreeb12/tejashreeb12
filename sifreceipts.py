from email.mime.application import MIMEApplication
import smtplib
import pandas as pd
from jinja2 import Environment, FileSystemLoader
import pdfkit

def send_email(to, subject, body, pdf_file):
    sender = '  '
    password = '  '
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
        print(f'Email sent to {to}')
    except Exception as e:
        print(f'Error sending email to {to}: {e}')

def generate_pdf(html_file, pdf_file):
    pdfkit.from_file(html_file, pdf_file, options={"enable-local-file-access":""})

def main():
    # Load the data from the spreadsheet
    df = pd.read_excel('receipt_data.xlsx')
    
    # Configure Jinja2 environment
    env = Environment(loader=FileSystemLoader('.'))
    template = env.get_template('receipt_template.html')
    
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
        send_email(to, subject, body, pdf_file)

def new_func():
    return 'NewReceipthtml.html'

if __name__ == '__main__':
    main()

# In this code, the receipt_data.xlsx file should contain a spreadsheet with columns for Email, Name, Date, Amount, and Check Number. 
# The main() function loads the data from the spreadsheet into a Pandas DataFrame, configures a Jinja2 environment, and iterates over each row in the DataFrame. 
#  For each row, it renders the HTML template with data from the spreadsheet, generates a unique PDF file name based on the email address, generates the PDF file from the HTML template using pdfkit, and sends the email with the PDF attachment using the send_email() function. 
#  The body variable is a string that contains placeholders for the actual name, amount, and date values. 
# These values can be replaced with real data using string formatting, such as `f'Thank you for your payment of ${row["Amount"]:.2f} on {row

