import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Read Excel data
df_contacts = pd.read_excel('final reminder.xlsx')
df_expenses = pd.read_excel('final reminder.xlsx')

# Configure SMTP server
smtp_server = 'inappmail.atrapa.deloitte.com'
smtp_port = 25  # or the appropriate port
sender_email = 'fromdeeptisagar@deloitte.com' #input the sender email
bcc_email = 'inmumappsupport3@deloitte.com'  # Add your BCC email address here

# Read HTML content from file with Windows-1252 encoding
with open('with table.htm', 'r', encoding='windows-1252') as file:
    email_template = file.read()

# Initialize counter for sent emails
sent_count = 0

def send_email(recipient_name, recipient_email, in_progress, completed, pending, grand_total):
    try:
        # Connect to SMTP server
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        
        # Replace placeholders in email template
        email_content = email_template.replace('[Name]', recipient_name)
        email_content = email_content.replace('[In-Progress]', str(in_progress))
        email_content = email_content.replace('[Completed]', str(completed))
        email_content = email_content.replace('[Pending]', str(pending))
        email_content = email_content.replace('[Grand Total]', str(grand_total))

        # Create message
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = '360-feedback form - Status Report'
        msg['Bcc'] = bcc_email

        # Attach email content
        msg.attach(MIMEText(email_content, 'html'))

        # Attach file
        filename = 'Process Note.pdf'  # Fixed filename
        attachment = open(filename, "rb")
        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', "attachment; filename= " + filename)
        msg.attach(part)

        # Send email
        server.sendmail(sender_email, [recipient_email] + [bcc_email], msg.as_string())
        
        # Increment counter and print success message
        global sent_count
        sent_count += 1
        print(f'Sent email to {recipient_email} (BCC to {bcc_email})')

    except Exception as e:
        print(f'Error sending email to {recipient_email}: {str(e)}')

    finally:
        # Close SMTP server connection
        try:
            server.quit()
        except:
            pass

# Send personalized emails
for index, row in df_contacts.iterrows():
    recipient_name = row['Name']
    recipient_email = row['Email Id']
    in_progress = row['In-progress']
    completed = row['Completed']
    pending = row['Pending']
    grand_total = row['Grand Total']
    
    send_email(recipient_name, recipient_email, in_progress, completed, pending, grand_total)

# Print total number of sent emails
print(f'Total emails sent: {sent_count}')
