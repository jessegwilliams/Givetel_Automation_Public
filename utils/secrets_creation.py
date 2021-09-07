import os
import json

SECRETS_FOLDER = os.path.join(os.getcwd(), 'Secrets')
LOGS_FOLDER = os.path.join(os.getcwd(), 'Logs')

if not os.path.exists(SECRETS_FOLDER): os.mkdir(SECRETS_FOLDER)
if not os.path.exists(LOGS_FOLDER): os.mkdir(LOGS_FOLDER)

def get_cs_username_from_user(): return {'CS_USER': input('Please enter your Contactspace username: ')}

def get_cs_password_from_user(): return {'CS_PASS': input('Please enter your Contactspace password: ')}

def get_cs_api_key_from_user(): return {'CS_API': input('Please enter your Contactspace API Key: ')}

def get_secrets():
    '''Retrieve all critical credentials from user. If the credentials don't exist, the user must enter them before running the program.'''
    secrets = {}
    cs_secrets = {}
    print('Hello! You need to enter some credentials for this code to work.')
    cs_secrets.update(get_cs_username_from_user())
    cs_secrets.update(get_cs_password_from_user())
    cs_secrets.update(get_cs_api_key_from_user())
    secrets.update({'CS_SECRETS': cs_secrets})
    return secrets

def create_secrets():
    secrets = get_secrets()
    user_input = input("Write secrets files? (Y/N): ")
    while user_input != 'Y' and user_input != 'N':
        user_input = input("Invalid input. Write secrets files? (Y/N): ")
    if user_input == 'Y':
        with open('Secrets/cs_credentials.json', 'w') as out_file:
            json.dump(secrets['CS_SECRETS'], out_file)

try:
    CS_CREDENTIALS_PATH = "Secrets/cs_credentials.json"
except:
    create_secrets()

if __name__ == '__main__':
    create_secrets()