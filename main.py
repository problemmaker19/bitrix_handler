import logging
import os

from dotenv import load_dotenv
from flask import Flask, request

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
import google_auth_oauthlib
from googleapiclient.discovery import build

from fast_bitrix24 import Bitrix
from googleapiclient.errors import HttpError

load_dotenv()

webhook = os.getenv('WEBHOOK')

client_id = os.getenv('CLIENT_ID')

client_secret = os.getenv('CLIENT_SECRET')

script_id = os.getenv('SCRIPT_ID')

application = Flask(__name__)

b = Bitrix(webhook)

logging.getLogger('fast_bitrix24').addHandler(logging.StreamHandler())

SCOPES = [
    'https://www.googleapis.com/auth/script.projects',
    'https://www.googleapis.com/auth/drive',
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


@application.route('/add.py', methods=['POST'])
def add_spreadsheet():
    add_token = 'vfhxhztziufdhgcfwbasxj5fnpvlxipz'

    deal_id = request.form.get('data[FIELDS][ID]')
    webhook_token = request.form.get('auth[application_token]')

    if add_token != webhook_token:
        return 'invalid token', 401

    try:
        params = {"ID": deal_id}
        method = "crm.deal.get"
        deal = b.call(method, params)

        function_name = 'FillTemplate'
        parameters = fetch_data(deal)

        result = execute_google_script(script_id, function_name, parameters)

        params = {"ID": deal_id, "fields": {"UF_CRM_1679160281": result}}
        b.call('crm.deal.update', params)

        return {'result': 'ok'}, 200

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


@application.route('/update.py', methods=['POST'])
def update_spreadsheet():
    update_token = 'e6z78vd5qjqqb1u6ljefyazrblfa7taq'

    deal_id = request.form.get('data[FIELDS][ID]')
    webhook_token = request.form.get('auth[application_token]')

    if update_token != webhook_token:
        return 'invalid token', 401

    try:
        params = {"ID": deal_id}
        method = "crm.deal.get"
        deal = b.call(method, params)

        if not deal['UF_CRM_1679160281']:
            return {'result': 'link empty'}, 200

        link = deal['UF_CRM_1679160281']
        spreadsheet_id = link[39:83]

        function_name = 'UpdateSheet'
        parameters = fetch_data(deal)
        parameters.append(spreadsheet_id)

        result = execute_google_script(script_id, function_name, parameters)

        return {'result': 'ok'}, 200

    except HttpError as error:
        print(f"An error occurred: {error}")
        return error


@application.route('/hello', methods=['GET'])
def check_connection():
    return 'РАСЛ ШТАНЫ НАДУЛ'


def fetch_data(deal):
    params = {}
    method = 'user.get'
    email_list = {'editor': '', 'helper': '', 'viewer': ''}
    contact_data = {'full_name': '', 'phone': ''}

    if deal["UF_CRM_1679160353"]:
        params["ID"] = deal["UF_CRM_1679160353"]
        user = b.call(method, params)
        email_list['editor'] = user['EMAIL']

    if deal["UF_CRM_1679160378"]:
        params["ID"] = deal["UF_CRM_1679160378"]
        user = b.call(method, params)
        email_list['helper'] = user['EMAIL']

    if deal["UF_CRM_1679160402"]:
        params["ID"] = deal["UF_CRM_1679160402"]
        user = b.call(method, params)
        email_list['viewer'] = user['EMAIL']

    if deal["CONTACT_ID"]:
        params["ID"] = deal["CONTACT_ID"]
        method = 'crm.contact.get'
        contact = b.call(method, params)
        contact_data['full_name'] = f'{str(contact["NAME"] or "")} ' \
                                    f'{str(contact["SECOND_NAME"] or "")} ' \
                                    f'{str(contact["LAST_NAME"]) or ""} '
        contact_data['phone'] = ' '.join([x['VALUE'] for x in contact['PHONE']])

    contract_num = deal['UF_CRM_1679160251']
    subject = deal['UF_CRM_1679160454']
    comment = deal['UF_CRM_1679160464']
    address = deal['UF_CRM_1679228043']

    return [contract_num, contact_data['full_name'], contact_data['phone'],
            address, subject, comment,
            email_list['editor'], email_list['viewer'], email_list['helper']]


def execute_google_script(s_id, function_name, parameters):
    # Authenticate and build the Google Scripts service
    service = build('script', 'v1', credentials=creds)

    # Execute the Google Script function with the given parameters
    req = {
        'function': function_name,
        'parameters': parameters,
        'devMode': True,
    }

    response = service.scripts().run(scriptId=s_id, body=req).execute()

    # Return the response from the Google Script function
    return response['response'].get('result')


if __name__ == '__main__':
    application.run(debug=True)
