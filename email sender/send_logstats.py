import win32com.client as win32

def send_email(sender, recipients, subject, body, attachment_path):
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.SentOnBehalfOfName = sender
    mail.Subject = subject
    mail.Body = body
    mail.To = ";".join(recipients)  # Concatenate recipients with semicolon (;) separator
    
    if attachment_path:
        mail.Attachments.Add(attachment_path)
    
    mail.Send()

def main():
    senders = [
        "80_TC_DET_CDR@dsc.army.mil",
        "151_TC_DET_CDR@dsc.army.mil",
        "258_TC_DET_CDR@dsc.army.mil",
        "259_TC_DET_CDR@dsc.army.mil",
        "271_TC_DET_CDR@dsc.army.mil",
        "383_TC_DET_CDR@dsc.army.mil",
        "602_TC_DET_CDR@dsc.army.mil",
        "606_TC_DET_CDR@dsc.army.mil",
        "823_TC_DET_CDR@dsc.army.mil",
        "940_TC_DET_CDR@dsc.army.mil"
    ]
    
    recipients = [
        "zachary.ramirez@dsc.army.mil"
    ]
    
    subjects = [
        "80_TC_DET Logstat",
        "151_TC_DET Logstat",
        "258_TC_DET Logstat",
        "259_TC_DET Logstat",
        "271_TC_DET Logstat",
        "383_TC_DET Logstat",
        "602_TC_DET Logstat",
        "606_TC_DET Logstat",
        "823_TC_DET Logstat",
        "940_TC_DET Logstat"
    ]
    
    body = "This is the body of the email."  # Single body for all emails
    
    attachments = [
        "logstat-80-day3-am.xlsx",
        "logstat-151-day3-am.xlsx",
        "logstat-258-day3-am.xlsx",
        "logstat-259-day3-am.xlsx",
        "logstat-271-day3-am.xlsx",
        "logstat-383-day3-am.xlsx",
        "logstat-602-day3-am.xlsx",
        "logstat-606-day3-am.xlsx",
        "logstat-823-day3-am.xlsx",
        "logstat-940-day3-am.xlsx"
    ]
    
    if len(senders) != len(subjects) != len(attachments):
        print("Error: Number of senders, subjects, and attachments must match!")
        return
    
    for sender, subject, attachment in zip(senders, subjects, attachments):
        send_email(sender, recipients, subject, body, attachment)
        print(f"Email sent from {sender} to {recipients} with subject '{subject}' and attachment '{attachment}'.")

if __name__ == "__main__":
    main()

