#!/usr/bin/env python

from pygerrit2 import GerritRestAPI, HTTPBasicAuth
import json
import os

def gerrit_login(gerrit_key, creds=None):
    try:
        gerrit_cred = None
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        with open(BASE_DIR+'/config/gerrit.json', 'r') as gerrit_file:
            gerrit_cred = json.load(gerrit_file)
        if not creds:
            auth = HTTPBasicAuth(gerrit_cred[gerrit_key]['username'], gerrit_cred[gerrit_key]['password'])
        else:
            auth = HTTPBasicAuth(creds['username'], creds['password'])
        gerrit = GerritRestAPI(url = gerrit_cred[gerrit_key]['url'], auth = auth)
        print('Successfully logged into Gerrit...!')
        return gerrit
    except Exception as e:
        print(e)
        raise Exception('Gerrit login failed')

if __name__ == '__main__':
    gerrit_login()    
