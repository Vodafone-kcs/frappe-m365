import frappe
from frappe import _
import requests
from email.parser import Parser
from email.policy import SMTPUTF8, default

def send_email_m365(email_queue, sender, recipient, message):

    email_account_doc = email_queue.get_email_account()
    connected_app_name = email_account_doc.connected_app
    connected_user = email_account_doc.connected_user

    connected_app = frappe.get_doc("Connected App", connected_app_name)
    oauth_token = connected_app.get_active_token(connected_user)
    access_token = oauth_token.get_password("access_token")

    subject = Parser(policy=default).parsestr(email_queue.message)["Subject"]
    # content = Parser(policy=default).parsestr(email_queue.message)
    content = subject
			
    email_data = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "Text",
                "content": content
            },
            "toRecipients": [{"emailAddress": {"address": recipient}}]
        }
    }   

    # Send the email
    response = requests.post(
        'https://graph.microsoft.com/v1.0/me/sendMail',
        headers={
            'Authorization': 'Bearer ' + access_token,
            'Content-Type': 'application/json'
        },
        json=email_data
    )

    frappe.log_error("Send email 365 by API", response.text)
    print(response.text)