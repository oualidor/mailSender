import smtplib
import xlrd

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

MY_ADDRESS = 'walid.khial@gmail.com'
PASSWORD = 'bpwrhmiusayggsnj'
HOST = 'smtp.gmail.com'
PORT = 587

def readList(file):
    #Open the document, on_demand option used to load the file partially
    workbook = xlrd.open_workbook(file, on_demand=True)
    #Opening the first page on the classer, use workbook.sheet_by_name('page name') to open by page Name
    worksheet = workbook.sheet_by_index(0)
    i = 1
    mails = []
    try:
        while True:

            mails.append(worksheet.cell(i, 9).value)
            i = i+1
    except Exception as e:
        print(e)

    return mails

def main():
    mails = readList('data.xlsx')
    if len(mails) != 0:

        # set up the SMTP server
        s = smtplib.SMTP(host=HOST, port=PORT)
        s.starttls()
        s.login(MY_ADDRESS, PASSWORD)

        # For each contact, send the email:
        for mail in mails:
            # create a message
            msg = MIMEMultipart()


            message = "hi, this is test"

            # setup the parameters of the message
            msg['From'] = MY_ADDRESS
            msg['To'] = "lidoo.kh@gmail.com"
            msg['Subject'] = "This is TEST"

            # add in the message body
            msg.attach(MIMEText(message, 'plain'))

            # send the message via the server set up earlier.
            s.send_message(msg)
            s.quit()

if __name__ == '__main__':
    #main()
    print('start')
    mails = readList('data.xlsx')
    if len(mails) != 0:
        for mail in mails:
            print(mail)
