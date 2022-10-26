import O365
from O365 import Account
from time import sleep
from O365 import Connection
from O365 import message
from Credentials import *


# This code works even with proxy


def Send_Email(receiverList, filepath,date,now):

    account = Account(credentials)

    # Call this line below  one time before use
    #account.authenticate(scopes=['basic', 'message_all'])
    # notice we insert an image tag with source to: "cid:{content_id}"

    if not account.is_authenticated:  # will check if there is a token and has not expired
        # ask for a login
        # console based authentication See Authentication for other flows
        account.authenticate(scopes=['basic', 'message_all'])



    output= 'ข้อมูล สภาพอากาศ AQI ประจำวันที่ '+str(date)+' ดึง เมื่อเวลา '+str(now)
    body = """
        <html>
            <body>
                <strong> ข้อมูล สภาพอากาศ AQI   </strong>
                <p>
                    <b> สวัสดีครับ </b>
                    <h1 style="color:Tomato;">
                    """+ output+"""</h1>
                </p>
            </body>
        </html>
    
        """
    m = account.new_message()
    for receiver in receiverList:    
        m.to.add(receiver)
    m.subject = 'ข้อมูล สภาพอากาศ AQI ประจำวันที่ '+str(date)+' ดึง เมื่อเวลา '+str(now)
    try:
        m.attachments.add(filepath)
    except:
        print(' no file attached ')
    m.body = body
    m.send()


