#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys
import imaplib
import smtplib
import email
import email.header
import email.message
from email.mime.text import MIMEText


if len(sys.argv) < 8:
    print 'Too few arguments. Give your Email account, Email folder, Incoming mail server (IMAP), Outgoing mail server (SMTP), Email port (SMTP), Password, sender and receiver.'
    print 'For example: \n python email_forwarding_machine.py test Inbox imap.example.com smtp.example.com 587 paSSw0rd sender@gmail.com receiver@gmail.com'
    sys.exit(0)

EMAIL_ACCOUNT = sys.argv[1]
EMAIL_FOLDER = sys.argv[2]
incoming_server = sys.argv[3]
outgoing_server = sys.argv[4]
email_port = sys.argv[5]
password = sys.argv[6]
sender = sys.argv[7]
to_address = sys.argv[8]

# go through and send all emails in the selected folder
def process_mailbox(M):

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print "No messages found!"
        return

# for each email in the selected folder: get an original email, build a new email to be forwarded and send it

    emails_not_sent = []

    for num in data[0].split():

        try:
            #get an original email from the folder
            rv, data = M.fetch(num, '(RFC822)')
            if rv != 'OK':
                print "ERROR getting message", num
                return
            msg = email.message_from_string(data[0][1])

            # get information of the original email
            originalfrom = get_original_value(msg['From'])
            originalto = get_original_value(msg['To'])
            originalcc = get_original_value(msg['Cc'])
            originalsubject = original_subject(msg)
            originalcontent = separate_parts(msg)
            originaldate = msg['Date']

            # convert the information of the original email to unicode
            originalcontent = convert_to_unicode(originalcontent)
            originalfrom = convert_to_unicode(originalfrom)
            originalto = convert_to_unicode(originalto)
            originalcc = convert_to_unicode(originalcc)
            originalsubject = convert_to_unicode(originalsubject)
            originaldate = convert_to_unicode(originaldate)

            # create a new message including the content and headers of an original email
            if originalcc == None:
                content = "Begin forwarded message:\nFrom: %s\nSubject: %s\nDate: %s\nTo: %s\n\n%s" % (originalfrom, originalsubject, originaldate, originalto, originalcontent)
            else:
                content = "Begin forwarded message:\nFrom: %s\nSubject: %s\nDate: %s\nTo: %s\nCc: %s\n\n%s" % (originalfrom, originalsubject, originaldate, originalto, originalcc, originalcontent)

            message = MIMEText(content.encode('utf-8', 'replace'), _subtype='plain',_charset='utf-8')
            message['From'] = sender
            message['To'] = to_address
            message['Subject'] = "Fwd: " + originalsubject

            #send a new email including the message to be forwarded
            s = smtplib.SMTP(outgoing_server, int(email_port))
            s.starttls()
            s.login(EMAIL_ACCOUNT, password)
            s.sendmail(sender, [to_address], message.as_string())
            s.quit()

            print num

        # Create a list of emails that were not sent due to possible errors
        except:
            emails_not_sent.append(num)

    if not emails_not_sent:
        print "All emails from the selected email folder forwarded successfully"
    else:
        print "Oooops, unable to forward the following %s emails: %s" % (len(emails_not_sent), emails_not_sent)
        print "All other emails from the selected email folder forwarded successfully"


# function that converts text to unicode
def convert_to_unicode(parameter):
    if parameter != None:
        if isinstance(parameter, unicode) == False:
            parameter = unicode(parameter, "utf-8")
            return parameter
        else:
            return parameter
    else:
        return parameter


# separate message parts of an original email content and return content as text/html
def separate_parts(msg):

    if msg.is_multipart():

        html = None
        text =""
        content = ""

        for part in msg.get_payload():
            # ignore images and attachments
            if 'image' in part.get_content_type():
                continue
            elif 'application' in part.get_content_type():
                continue
            elif 'calendar' in part.get_content_type():
                continue
            else:

                # case 1: content charset unavailable
                if part.get_content_charset() is None:
                    content = part.get_payload(decode=False)
                    i = 0
                    while i<len(content):
                        charset = content[i].get_content_charset()
                        if content[i].get_content_type() == 'text/plain':
                            text = unicode(content[i].get_payload(decode=True), str(charset), "ignore")
                        if content[i].get_content_type() == 'text/html':
                            html = unicode(content[i].get_payload(decode=True), str(charset), "ignore")
                        i += 1
                    continue

                # case 2: content charset available
                charset = part.get_content_charset()
                if part.get_content_type() == 'text/plain':
                    text = unicode(part.get_payload(decode=True), str(charset), "ignore")
                if part.get_content_type() == 'text/html':
                    html = unicode(part.get_payload(decode=True), str(charset), "ignore")

        # return text or html part of the email
        if text is not None:
            return text.strip()
        else:
            return html.strip()

    else:
        text = unicode(msg.get_payload(decode=True), msg.get_content_charset(), 'ignore').encode('utf8', 'replace')
        return text.strip()


# return header parameter or make some corrections with encoding/decoding of scandinavian letters and return the parameter then
def get_original_value(parameter):

    original_value = parameter

    try:
        if "=?" in original_value:

            decoded_header = email.header.decode_header(parameter)[0]
            encoding = decoded_header[1]
            if not encoding:
                encoding = 'iso-8859-1'
            decoded_part1 = decoded_header[0].decode(encoding=encoding, errors='replace')

            decoded_header2 = email.header.decode_header(parameter)[1]
            encoding2 = decoded_header2[1]
            if not encoding2:
                encoding2 = 'iso-8859-1'
            decoded_part2 = decoded_header2[0].decode(encoding=encoding, errors='replace')

            original_value = decoded_part1 + str(" ") + decoded_part2

        return original_value

    except:
        return original_value


# get the original subject of an email
def original_subject(msg):

    decoded_header = email.header.decode_header(msg['Subject'])[0]
    encoding = decoded_header[1]
    if not encoding:
        encoding = 'iso-8859-1'
    originalsubject = decoded_header[0].decode(encoding=encoding, errors='replace')
    return originalsubject


# Executable code starts here...

# connect to incoming mail server
M = imaplib.IMAP4_SSL(incoming_server)

try:
    print "Connecting to incoming mail server"
    rv, data = M.login(EMAIL_ACCOUNT, password)
except imaplib.IMAP4.error:
    print "LOGIN FAILED!!! "
    sys.exit(1)

print rv, data

# select an email folder to be processed and go through and send all the emails inside
rv, data = M.select(EMAIL_FOLDER)
if rv == 'OK':
    print "Processing mailbox...\nPrinting an ID of each succesfully sent email:"
    process_mailbox(M)
    M.close()

else:
    print "ERROR: Unable to open mailbox ", rv

M.logout()
