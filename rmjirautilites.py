import os
from jira import JIRA
from jira.exceptions import JIRAError
import json


def jira_login(username=None, pwd=None):
    credentials = {}
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    if not username or not pwd:
        with open(BASE_DIR+'/config/credentials.json', 'r') as cred_file:
            credentials = json.load(cred_file)
        username = credentials['username']
        pwd = credentials['password']

    jira_server = 'https://ccp.sys.comcast.net/'
    jira_options = {'server': jira_server}

    jira = JIRA(options=jira_options, basic_auth=(username, pwd))
    print('Jira loggin successfull')
    return jira


if __name__ == '__main__':
    jira_login()

