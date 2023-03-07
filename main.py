import logging
import os

from dotenv import load_dotenv
from flask import Flask, request

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import google_auth_oauthlib
from googleapiclient.discovery import build

from fast_bitrix24 import Bitrix

load_dotenv()

application = Flask(__name__)

webhook = os.getenv('WEBHOOK')

b = Bitrix(webhook)

logging.getLogger('fast_bitrix24').addHandler(logging.StreamHandler())

client_id = os.getenv('CLIENT_ID')

client_secret = os.getenv('CLIENT_SECRET')


SCOPES = [
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/spreadsheets'
]

creds = None

# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)

# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        creds = google_auth_oauthlib.get_user_credentials(
            SCOPES,
            client_id,
            client_secret
        )
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())


@application.route('/run.py', methods=['POST'])
def run_script():
    print('------------- CREATING ---------------')

    deal_id = int(request.values.get('document_id[2]')[5:])

    title = request.args.get('t', default=' ')
    contract_num = request.args.get('con', default=' ')
    customer = request.args.get('cus', default=' ')
    date = request.args.get('d', default=' ')
    phone = request.args.get('p', default=' ')
    address = request.args.get('a', default=' ')
    subject = request.args.get('s', default=' ')
    comments = request.args.get('com', default=' ')
    editor_email = request.args.get('e')
    viewer_email = request.args.get('v')
    helper_email = request.args.get('h')

    # Define the script ID, function name, and parameters
    script_id = 'YOUR_SCRIPT_ID'
    function_name = 'FillTemplate'
    parameters = [title, contract_num, customer, date, phone, address, subject, comments,
                  editor_email, viewer_email, helper_email]
    result = execute_google_script(script_id, function_name, parameters)

    params = {"ID": deal_id, "fields": {"UF_CRM_1677759756": result}}
    b.call('crm.deal.update', params, raw=True)

    print('------------- CREATE FINISHED ---------------')

    return {'result': 'ok'}, 200


@application.route('/update.py', methods=['POST'])
def update_script():

    title = request.args.get('t', default=' ')
    contract_num = request.args.get('con', default=' ')
    customer = request.args.get('cus', default=' ')
    date = request.args.get('d', default=' ')
    phone = request.args.get('p', default=' ')
    address = request.args.get('a', default=' ')
    subject = request.args.get('s', default=' ')
    comments = request.args.get('com', default=' ')
    editor_email = request.args.get('e')
    viewer_email = request.args.get('v')
    helper_email = request.args.get('h')
    link = request.args.get('link')

    script_id = 'YOUR_SCRIPT_ID'

    if not link:
        print('------------- EMPTY LINK ---------------')
        return {'result': 'link empty'}, 200

    print('------------- UPDATING ---------------')

    spreadsheet_id = link[39:83]

    function_name = 'UpdateSheet'
    parameters = [title, contract_num, customer, date, phone, address, subject, comments, spreadsheet_id, editor_email,
                  viewer_email, helper_email]
    execute_google_script(script_id, function_name, parameters)

    print('------------- UPDATE FINISHED ---------------')

    return {'result': 'ok'}, 200


@application.route('/hello', methods=['GET'])
def check_connection():
    return 'РАСУЛ ШТАНЫ НАДУЛ'


def execute_google_script(script_id, function_name, parameters):
    # Authenticate and build the Google Scripts service
    service = build('script', 'v1', credentials=creds)

    # Execute the Google Script function with the given parameters
    req = {
        'function': function_name,
        'parameters': parameters,
        'devMode': True,
    }

    response = service.scripts().run(scriptId=script_id, body=req).execute()

    # Return the response from the Google Script function
    return response['response'].get('result')


if __name__ == '__main__':
    application.run(debug=True)
