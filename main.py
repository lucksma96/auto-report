import imaplib, ssl, email, os, re, locale, pikepdf, json
from datetime import date
from pathlib import Path
from getpass import getpass

# --- fuctions --- #

def is_valid_file(part):
    return bool(part.get_filename()) and 'application/pdf' in part.get('Content-Type')

def decode_and_parse_message(data):
    raw_email = data[0][1]
    raw_email_string = raw_email.decode('UTF-8')
    email_message = email.message_from_string(raw_email_string)
    return email_message

def clear_sender(sender):
    sender = sender.lower()
    sender = re.search('<(.+)>', sender).group(1)
    return sender

def get_sender_filename(sender):
    sender_attachment_map = json.load(open('sender_filenames.json'))
    return sender_attachment_map[sender] + '.pdf'

def cleanup_pdf(filepath):
    passwords = json.load(open('passwords.json'))

    for passw in passwords:
        try:
            with pikepdf.open(filepath, password=passw, allow_overwriting_input=True) as pdf:
                for p in pdf.pages:
                    if p.index == 0:
                        continue
                    pdf.pages.remove(p)
                pdf.save(filepath)
        except pikepdf.PasswordError:
            continue

# --- main --- #

config = json.load(open('config.json', encoding='UTF-8'))

locale.setlocale(locale.LC_TIME, 'pt_BR')
BASE_FOLDER_PATH = config['basepath']
now = date.today()
CURRENT_YEAR = now.year
LAST_MONTH = now.replace(day=1, month=now.month - 1).strftime('%m-%B')
CURRENT_FOLDER_PATH = BASE_FOLDER_PATH + '\\' + str(CURRENT_YEAR) + '\\' + LAST_MONTH
Path(CURRENT_FOLDER_PATH).mkdir(parents=True, exist_ok=True)

context = ssl.SSLContext(ssl.PROTOCOL_SSLv23)
mailbox = imaplib.IMAP4_SSL(config['imaphost'], ssl_context=context)
user_email = config.get('useremail')
if user_email is None:
    user_email = input('Email: ')
else:
    print(f'Email: {config.get('useremail')}')
user_pass = getpass()
mailbox.login(user_email, user_pass)
mailbox.select('Inbox')

locale.setlocale(locale.LC_TIME, 'en_US')
since = now.replace(month=now.month - 2).strftime('%d-%b-%Y')
text = now.replace(month=now.month - 1).strftime('%m/%Y')
type, data = mailbox.search(None, f'TEXT */{text} SINCE {since} OR (OR FROM contaporemail@ceee.com.br FROM faturadigital@minhaclaro.com.br) (OR FROM contadigital@vivo.com.br FROM contadigitalvivo@vivo.com.br)')

id_list = data[0].split()

for id in id_list:
    type, data = mailbox.fetch(id, '(RFC822)')
    email_message = decode_and_parse_message(data)

    for part in email_message.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue
        if is_valid_file(part):
            sender = clear_sender(email_message['From'])
            filename = get_sender_filename(sender)
            filepath = os.path.join(CURRENT_FOLDER_PATH, filename)
            if not os.path.isfile(filepath):
                fp = open(filepath, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()

            cleanup_pdf(filepath)