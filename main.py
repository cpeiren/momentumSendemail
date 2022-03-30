import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
from datetime import datetime, timedelta

todayDatetime = datetime.today()
#todayDatetime = datetime(2022, 3, 27)
today = todayDatetime.strftime('%Y-%m-%d')

def sendEmail(email_address):
    mail.To = email_address
    mail.Subject = 'Digital Asset Position'
    mail.Body = 'Digital Asset Position'
    #mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

    # To attach a file to the email (optional):
    attachment  = fr"C:\Users\pchen\Desktop\MomentumProduction\Production\trade_txt\{today}trade_doc.txt"
    mail.Attachments.Add(attachment)

    mail.Send()

if __name__ == '__main__':
    sendEmail('pchen@rqsi.com; jkarnawat@rqsi.com; tradedesk@rqsi.com; nramsey@rqsi.com;jwolfe@rqsi.com')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
