import requests
import json


url = "https://outlook.office.com/webhook/c6b2d60e-d541-4de8-9879-4a3f7986e93b@6c637512-c417-4e78-9d62-b61258e4b619/IncomingWebhook/7354b1715c8344aba336521ee11aa849/b2ca721d-b8a7-4c8a-9fbe-40e9182fd8b3"
# url = "https://outlook.office.com/webhook/9fde5266-e7fc-48e4-ad16-6e600aaaf946@6c637512-c417-4e78-9d62-b61258e4b619/IncomingWebhook/5c50f353601948a08990c882776ebe46/b2ca721d-b8a7-4c8a-9fbe-40e9182fd8b3"

message = {
    'title': "Microsoft Teams Automation",
    'text': "This message was written by a Python program.\n\n  https://www.youtube.com/watch?v=MgK-hbCSUeY  \n\n For more info watch the video below!"
    }


def main():
    response_body = requests.post(url=url, data=json.dumps(message))


main()
