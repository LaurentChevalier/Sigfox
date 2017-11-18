import imaplib
import email
import time
from datetime import date
from time import gmtime, strftime


while 1:
    time_now=strftime("%d_%m_%Y %HH%M", gmtime())
    time_now_file=strftime("%d/%m/%Y %H:%M", gmtime())
    msrv = imaplib.IMAP4_SSL('imap.gmail.com',993)
    msrv.login('your_email@example.com','yourpassword')
    msrv.list()
    #msrv.select('INBOX')
    msrv.select('INBOX')

    result, data = msrv.uid('search', None, "ALL")

    n = len(data[0].split())

    for i in range(n):
        latest_email_uid = data[0].split()[i] # unique ids wrt to label selected
        result, email_data = msrv.uid('fetch', latest_email_uid, '(RFC822)')
        raw_email = email_data[0][1]
        raw_email_string = raw_email.decode('utf-8')
        email_message = email.message_from_string(raw_email_string)
        for part in email_message.walk():
            if part.get_content_type() == "text/plain":
                body = part.get_payload(decode=True)
                if email_message['subject'] == "F8EAB":
                    print 'Subject:',email_message['subject']
                    try:
                        print body.decode('utf8')
                        chaine=body.decode('utf8')
                        #sousChaine1 = chaine[14:16]
                        sousChaine2 = chaine[16:18]
                        sousChaine3 = chaine[18:20]
                        sousChaine = sousChaine3+sousChaine2
                        valeurdec = int(sousChaine, 16)
                        valeurfloat = float(valeurdec)
                        valeurarrondi = round(valeurfloat * 0.0025,2)
                        index = int(valeurarrondi * 100)
                        print sousChaine
                        print valeurdec
                        print valeurarrondi
                        print index
                        device1 = email_message['subject']
                        path="D:\\Mes documents\DOCUMENTS ECOLE\\csv folder\\csv files to import\\"+time_now+".txt"
                        myfile=open(path, "a")
                        myfile.write(device1+","+time_now_file+"," + str(index)+"\n")
                        print "file write"
                        myfile.close()
                        typ, data = msrv.search(None, 'ALL')
                        for num in data[0].split():
                            msrv.store(num, '+FLAGS', '\\Deleted')
                        msrv.expunge()
                        msrv.close()
                        msrv.logout()
                        time.sleep(600)
                    except UnicodeDecodeError:
                        print body
                        typ, data = msrv.search(None, 'ALL')
                        for num in data[0].split():
                            msrv.store(num, '+FLAGS', '\\Deleted')
                        msrv.expunge()
                        msrv.close()
                        msrv.logout()
                        time.sleep(600)

