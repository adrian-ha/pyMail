
#import
import win32com.client as win32

def create_email(text, subject, mailTo=None, send=False, mailCC=None, mailBCC=None, mailFrom=None, attachements=[]):
    """
    Outlook emails can be send from within python script.

    :param text: email text
    :param subject: email header
    :param mailTo: list of recipients
    :param send: True=Send, False=Draft
    :param mailCC: list of cc recipients
    :param mailBCC: list of bcc recipients
    :param mailFrom: change sender
    :param attachements: list of paths of potential attachements
    :return:
    """
    #init email
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    #amend sender
    if mailFrom is not None:
        sender_account = None
        for account in outlook.Session.Accounts:
            if account.DisplayName == mailFrom:
                sender_account = account
                break
        if sender_account is not None:
            mail._oleobj_.Invoke(*(64209, 0, 8, 0, sender_account))

    #add content
    mail.Subject = subject
    mail.Body = text

    #add recipients
    if mailTo is not None:
        mail.To = ";".join(mailTo)
    if mailCC is not None:
        mail.CC = ";".join(mailCC)
    if mailBCC is not None:
        mail.BCC = ";".join(mailBCC)

    #add attachements
    if len(attachements) > 0:
        for attachement in attachements:
            mail.Attachments.Add(attachement)

    #Define whether email should be send or saved as draft
    if send:
        mail.send()
    else:
        mail.save()


if __name__ == '__main__':
    create_email(text="Hello all,\n\nthis is a simple test.\n\nAll the best,\ntbd", subject="TestMail", mailTo=["dummy1@gmail.com", "dummy2@gmail.com"])







