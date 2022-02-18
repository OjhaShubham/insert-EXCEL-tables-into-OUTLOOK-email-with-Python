
import os
import win32com.client as client
from PIL import ImageGrab

workbook_path=r'C:\Users\shojha\PycharmProjects\untitled\heatmap.xlsx'
print(type(workbook_path))

excel=client.Dispatch('Excel.Application')


wb = excel.Workbooks.Open(workbook_path)

sheet=wb.Sheets.Item(1)

print(sheet)

sheet=wb.Sheets[0]

#sheet = wb.Sheet['Sheet1']

excel.visible=1

copyrange= sheet.Range('A1:M11')

copyrange.CopyPicture(Appearance=1, Format=2)

ImageGrab.grabclipboard().save('paste.png')

excel.Quit()

image_path=r'C:\Users\shojha\PycharmProjects\untitled\paste.png'

html_body = """
    <div>
          Please review the following report and response with your feedback.
    </div>
    <div>
        <img src={}></img>
    </div>
"""

outlook = client.Dispatch('Outlook.Application')

# create a message
message = outlook.CreateItem(0)

# set the message properties
message.To = 'shubham-geetashankar.ojha@capgemini.com'
message.Subject = 'Please review!'
message.HTMLBody = html_body.format(image_path)

# display the message to review
message.Display()

# save or send the message
message.Send()