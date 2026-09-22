import datetime
from email.header import decode_header
from email.message import EmailMessage
import os
from imaplib import IMAP4_SSL
import smtplib
from openpyxl import load_workbook
import email

# fetch secret user and password from a file
file = open("usr_pass.txt", 'r')
info = file.read().split('\n')

# fetch encrypted mail username and password of school's mail
sch_user: str = info[0]
sch_pass: str = info[1]
file.close()

# which sender we are looking for and from which date we are looking
sender: str = info[2]
date: datetime.datetime = datetime.datetime.now()
two_days_ago: datetime.datetime = date - datetime.timedelta(days=140)

# what imap and smtp server we are using
imap_server: str = 'imap.gmail.com'
smtp_server: str = 'smtp.gmail.com'

# function for sending a mail with attachment to the relevant people
def send_mail(file_data, file_name: str, classroom: str, week_num: int, student: bool):

    # create a mail
    msg = EmailMessage()
    msg['Subject'] = f'Ukeplannr. {week_num} for klasse {classroom}'
    msg['From'] = 'Salahaddin Skole'
    msg['To'] = 'undisclosed-recipients:;'

    msg.set_content(f"""\
Vedlagt ukeplannr. {week_num}

Mvh
Salahaddin skoleadministrasjon
    """)
    msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

    # getting recipents
    bcc_addresses = []

    # check if we need to retrieve from student mails or teacher mails
    wanted = "student" if student else "teacher"
    ws = load_workbook("mails/mails.xlsx").active

    # parsing through the mails, sending to mail from corresponding classroom
    for ml, cr, role in ws.iter_rows(min_row=2, values_only=True):  # min_row=2 skips header
        if role == wanted and cr.lower() == classroom.lower():
            bcc_addresses.append(ml)

    # sending the mail
    print("sending to", bcc_addresses)
    with smtplib.SMTP_SSL(smtp_server, 465) as smtp:
        smtp.login(sch_user, sch_pass)
        smtp.send_message(msg, to_addrs=bcc_addresses)


# mail server i am connecting to is outlook/hotmail
imap: IMAP4_SSL = IMAP4_SSL(imap_server)

# log in to the mail
imap.login(sch_user, sch_pass)

# looking for mails from sender within the last 5 days in the inbox category
imap.select("Inbox")
formatted = two_days_ago.strftime('%d-%b-%Y')
status, tot_msgs = imap.search(None, f'FROM "{sender}" SINCE {formatted}')

# if it finds a mail that matches the criteria
if status == 'OK':
    print("scanner mail...\n")
    # the latest mail

    # we should be matching with excactly 2 mails
    tot_mails = len(tot_msgs[0].split())
    if tot_mails != 2:
        print(f"{tot_mails} mails funnet, burde være 2.")
        # closing and logging out for safe measure
        imap.close()
        imap.logout()
        input("Trykk på Enter for å avslutte...")
        os._exit(1)

    for msg in tot_msgs[0].split():

        # fetch each mail as bytes
        _, data = imap.fetch(msg, "(RFC822)")   
        # turn into message class
        msg = email.message_from_bytes(data[0][1])
        subject = msg.get('Subject')
        decoded_subject = decode_header(subject)

        # see if its weekly schedule for teachers or studfents
        sub_list = decoded_subject[0][0].split()
        if len(sub_list) != 3:
            print(f"Emnet burde se ut som eksempelet: 'Vs: elev ukeplannr.14', men ser slik ut: {subject}")
            input("Trykk på Enter for å avslutte...")
            os._exit(1)
        student = False
        elev = sub_list[1].lower()
        if elev == 'elev':
            student = True

        # get the week number
        week_num = str(sub_list[-1]).split('.')[1]

        # iterates through all the parts of the mail
        for part in msg.walk():

            # looks for all the weekly schedules 
            file_name = part.get_filename()

            if file_name:

                # retrieve which class the pdf is for
                classroom = file_name.split()[-1].split('.')[0]

                # send the schedule to the relevant people
                file_data = part.get_payload(decode=True)
                print(f'Sender {file_name}...')
                send_mail(file_data, file_name, classroom, week_num, student)
                print(f'{file_name} sendt!\n')


else:
    print("Fant ingen e-mail")

# closing and logging out for safe measure
imap.close()
imap.logout()

input("Trykk på Enter for å avslutte...")