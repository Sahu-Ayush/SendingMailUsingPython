from string import Template
from pathlib import Path
import os
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

'''
Function to read the contacts from a given contact file and return a
list of names and email addresses
'''

def get_contacts(filename):

    names = list()
    emails = list()

    with open(filename, mode='r', encoding='utf-8') as contacts_file:
        for a_contacts in contacts_file:
            # names, emails = a_contact.split('')
            names.append(a_contacts.split()[0])
            emails.append(a_contacts.split()[1])

    return names, emails

'''
Function to read in a template file (like message.txt)
and return a Template object made from its contents
'''

def read_template(filename):

    with open(filename, 'r', encoding='utf-8') as template_file:
        file_content_read = template_file.read()
        template_file_content = Template(file_content_read)

    return template_file_content


s = None
MY_ADDRESS = '********@gmail.com'

# Set up the SMTP server
def set_smtp():
    global s
    '''
    For outlook
    '''
    #s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    '''
    For Gmail
    For Gmail SMTP server address: smtp.gmail.com. Gmail SMTP port (TLS): 587. SMTP port (SSL): 465.
    s = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
    '''
    s = smtplib.SMTP(host='smtp.gmail.com', port=587)
    s.starttls()
    MY_ADDRESS = '******@gmail.com'
    PASSWORD = '******'
    s.login(MY_ADDRESS, PASSWORD)


def send_separatly_mail(names, emails, message_template, attach_file_name):
    for name, email in zip(names, emails):
        msg = MIMEMultipart()       # create a message

        # Add in the actual person name to the message template
        message = message_template.substitute(PERSON_NAME=name)

        # Prints out the message body for our sake
        print(message)

        # Setup the parameters of the message
        msg['From'] = MY_ADDRESS
        msg['To'] = email
        msg['Subject'] = "This is TEST"
            
        # Add in the message body
        msg.attach(MIMEText(message, 'plain'))
        # Attach_file_name = 'TP_python_prev.pdf'
        attachment = open(attach_file_name, 'rb') # Open the file as binary mode
        p = MIMEBase('application', 'octat-stream')
        p.set_payload((attachment).read())
        encoders.encode_base64(p) #encode the attachment
        # Add payload header with filename
        p.add_header('Content-Disposition', "attachment; filename= "+attach_file_name)
        msg.attach(p)
        # Send the message via the server set up earlier.
        s.send_message(msg)
        del msg
            
    # Terminate the SMTP session and close the connection
    s.quit()

# Main
contacts_file_path = input('Enter contacts_file_path:\n')
assert os.path.exists(contacts_file_path), "I did not find the file at, "+str(contacts_file_path)
message_file_path = input('Enter message_file_path:\n')
assert os.path.exists(message_file_path), "I did not find the file at, "+str(message_file_path)

num_attachment = int(input("If you want to attached file enter 1 else enter 0 \(only 1 file you can attache\)': "))
if num_attachment != 0:
    attach_file_name = input('Enter attachement_file_path:\n')
    assert os.path.exists(attach_file_name), "I did not find the file at, "+str(attach_file_name)

names, emails = get_contacts(contacts_file_path)
message_template = read_template(message_file_path)

print('Sending...')

set_smtp()

send_separatly_mail(names, emails, message_template, attach_file_name)

print('Done!')
