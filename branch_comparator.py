#!/usr/bin/env python

import os
import re, sys, requests, json
from urllib3.exceptions import HTTPError as BaseHTTPError
from rmGerritUtils import *
from datetime import datetime, timedelta
import pdb
import copy
import xlwt
import xmltodict
import configparser
from pytz import timezone
from rmjirautilites import *

from progress_bar import *

class BranchComparison:
    '''
    Comparator class which compares changes from the given source branch to the target branch
    by getting the changes from the Gerrit using APIs.

    Attributes
    --------------------------------
    branch1: <str>
        source branch name
    branch2: <str>
        target branch name
    offset: <int>
        iterator for the change set
    eob: <bool>
        flag to end of the change set
    commit: <dict>
        buffer for the iteration
    projects_log: <list>
        stores projects logs of the target branch of requires repos
    crossed_start: <bool>
        flag to check current date in range or not
    stage_1_data: <list>
        stage 1 data
    final_data: <list>
        final refined data
    start_time: <datetime.datetime>
        start date time
    end_time: <datetime.datetime>
        end date time
    is_dev_specific: <bool>
        device specific flag
    repos_to_be_checked: <list>
        list of repos to be checked
    devices: <list>
        list of device details to be checked
    manifests: <dict>
        device manifests details

    Methods
    --------------------------------
    update_merge_pending_list()
        updates the merge pending list
    check_in_branch(branch)
        checks the missing changes in the given branch
    all_devices_repos()
        gets all the repos requires for the devices specified
    get_repos(device)
        gets all the repos for the given device
    is_in_range(commit)
        checks the given commit is in the date range or not
    get_change_ids()
        gets the change ids of the target branch
    check_implicit_changes()
        checks the implicit changes in the target branch
    compare_branches(barnch1, branch2)
        checks the changes from branch1 to branch2
    generate_report()
        generates the report from final data
    '''

    branch1 = ''
    branch2 = ''
    offset = 0
    moffset = 0
    doffset = 0
    eob = False
    meob = False
    deob = False
    commit = {}
    crossed_start = False
    projects_log = {}
    stage_1_data = []
    final_data = {}
    start_time = None
    end_time = None
    is_dev_specific = False
    repos_to_be_checked = []
    devices = []
    manifests = None
    exceptional_repos = []
    gerrits = ['primary_gerrit', 'rdk_gerrit']
    gerrit_urls = {
        'primary_gerrit': 'https://gerrit.teamccp.com',
        'rdk_gerrit': 'https://code.rdkcentral.com'
    }
    current_gerrit = ''

    def __init__(self, start_time, end_time, cmd=True, timezone_val='utc'):
        '''
        Initializes the class instance with the required parameters        

        Parameters
        --------------------------------
        start_time: <str>
            start date time
        end_time: <str>
            end date time
        '''
        self.cmd = cmd
        self.timezone_val = timezone_val
        if start_time != 'NO_START' and end_time != 'NO_END':
            if cmd is True:
                self.start_time = datetime.strptime(start_time, '%Y-%m-%d-%H:%M:%S') - timedelta(hours=5, minutes=30)
                self.end_time = datetime.strptime(end_time, '%Y-%m-%d-%H:%M:%S') - timedelta(hours=5, minutes=30)
            else:
                self.start_time = start_time
                self.end_time = end_time

        self.BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        with open(self.BASE_DIR+'/config/manifests.json', 'r') as manifest_file:
            self.manifests = json.load(manifest_file)
        with open(self.BASE_DIR+'/config/devices.json', 'r') as devices_file:
            all_devices = json.load(devices_file)
            self.devices = []
            for device in all_devices.values():
                self.devices += device

        self.is_morty = False
        self.is_dunfell = False

    def update_merge_pending_list(self, branch):
        '''
        Updates the merge pending list  
        '''
        self.stage_1_data.append(copy.deepcopy(self.commit))
        if self.commit['project'] not in self.projects_log[branch].keys():
            self.projects_log[branch][self.commit['project']] = []
        
    def check_in_branch(self, branch):
        '''
        Checks the availability of the change in the given branch        

        Parameters
        --------------------------------
        branch: <str>
            brach name
        '''
        try:
            response = self.gerrit.get('/changes/'+self.commit['urlencoded_project']+'~'+branch+'~'+self.commit['change_id']+'/in')
            if not response['branches']:
                print('Fix is NOT available in %s\n' % branch)
                self.update_merge_pending_list(branch)
            elif branch not in response['branches']:
                print('\nFix is NOT available in ' + branch)
                self.update_merge_pending_list(branch)
            else:
                print('Fix is available in %s\n' % branch)

        except requests.exceptions.HTTPError:
            print('Fix is NOT available in %s\n' % branch)
            self.update_merge_pending_list(branch)

    def all_devices_repos(self):
        '''
        Gets all the repos of the devices specified
        '''
        self.is_dev_specific = True
        for device in self.devices:
            self.get_repos(device)

    def get_repos(self, device):
        '''
        Gets all the repos of the given device      

        Parameters
        --------------------------------
        device: <str>
            device name
        '''
        try:
            self.is_dev_specific = True
            manifest_api = '/projects/%s/branches/%s/files/%s'%(
                self.manifests[device]['project'].replace('/','%2F'),
                self.branch1, self.manifests[device]['manifest_file'] + '/content'
            )
            # print("***********************manifest api*****************************")
            # print(manifest_api)
            manifest_content = self.gerrit.get(manifest_api)
            # print('****************', device, self.manifests[device]['manifest_file'])
            if device in ['XCam', 'XCam2', 'ICam2', 'Doorbell']:
                #manifest = json.loads(manifest_content)
                projects = self.get_deps_content(manifest_content)
                for project in projects:
                    if project not in self.repos_to_be_checked:
                        self.repos_to_be_checked.append(project)
            else:
                manifest = xmltodict.parse(manifest_content)
                if self.is_morty is False or self.is_dunfell is False:
                    try:
                        yv = manifest['manifest'].get('yocto')
                        print(yv,'****************', device, self.manifests[device]['manifest_file'])
                        if yv:
                            if yv['@version'] == 'morty':
                                self.is_morty = True
                                print('Is morty True')
                            elif yv['@version'] == 'dunfell':
                                self.is_dunfell = True
                                print('Is dunfell True')
                            else:
                                print('!!!!!!!!!!!!!!!!!!!!!!!1111111', yv['@version'])
                    except Exception as e:
                        print('Exception while getting yacto version', e)

                for project in manifest['manifest']['project']:
                    if project['@name'] not in self.repos_to_be_checked:
                        self.repos_to_be_checked.append(project['@name'])
            #print('************Repos to be checked ***************', self.repos_to_be_checked)
        except Exception as e:
            print(e)
            raise Exception('Exception occured while getting device repos, '
                            'Please check manifest exist for "%s" in "%s"'
                            %(device, self.branch1))

    def get_deps_content(self, content):
        content = content.replace('\n}', '}').replace(',}', '}')
        config = configparser.RawConfigParser(allow_no_value=True)

        config.read_string('[myconf]' + content)

        y = config.get('myconf', 'deps', raw=True)
        y = y.replace("\'", "\"")

        content = json.loads(y)
        projects = []

        for k, v in content.items():
            if v.startswith('ssh://gerrit.teamccp.com'):
                project = re.search('ssh://gerrit.teamccp.com:29418/(.*)@',v)
                if project:
                    projects.append(project.group(1))

        return projects

    def is_in_range(self, commit):
        '''
        Checks the given commit is in the date range or not       

        Parameters
        --------------------------------
        commit: <dict>
            commit details to be checked
        '''
        if (not self.start_time) and (not self.end_time):
            return True
        updated_time = datetime.strptime(commit['updated'].split('.')[0], '%Y-%m-%d %H:%M:%S')
        merge_time = datetime.strptime(commit['submitted'].split('.')[0], '%Y-%m-%d %H:%M:%S')
        if self.cmd is False:
            updated_time = updated_time.replace(tzinfo=timezone('UTC'))
            merge_time = merge_time.replace(tzinfo=timezone('UTC'))

            if self.timezone_val == 'pst':
                updated_time = updated_time.astimezone(timezone('US/Pacific'))
                merge_time = merge_time.astimezone(timezone('US/Pacific'))
            elif self.timezone_val == 'ist':
                updated_time = updated_time.astimezone(timezone('Asia/Kolkata'))
                merge_time = merge_time.astimezone(timezone('US/Pacific'))
            elif self.timezone_val == 'est':
                updated_time = updated_time.astimezone(timezone('US/Eastern'))
                merge_time = merge_time.astimezone(timezone('US/Pacific'))

        if updated_time < self.start_time:
            self.crossed_start = True
            return False
        if merge_time >= self.start_time and merge_time <= self.end_time:
            return True
        return False

    def get_change_ids(self, branch2):
        '''
        Gets all the change ids of the target branch  
        '''
        total_iteration = 10
        total_length = len(self.projects_log[branch2].keys()) * total_iteration
        current_count = 0
        #printProgressBar (current_count, total_length, prefix = 'Checking implicit changes', suffix = '', decimals = 1, length = 100, fill = '||')
        for project in self.projects_log[branch2].keys():
            iteration_count = 0

            log_api_template = '/plugins/gitiles/' + project + '/+log/' + branch2
            next = ''
            while iteration_count < total_iteration:
                iteration_count += 1
                current_count += 1
                log_api = log_api_template + next
                try:
                    logs = self.gerrit.get(log_api)
                except requests.exceptions.HTTPError:
                    mod = current_count // total_iteration
                    rem = current_count % total_iteration
                    current_count = (mod + 1) * total_iteration
                    #printProgressBar (current_count, total_length, prefix = 'Checking implicit changes', suffix = '', decimals = 1, length = 100, fill = '||')
                    self.exceptional_repos.append(project)
                    break
                for log in logs['log']:
                    start = log['message'].find('Change-Id: ')
                    while start != -1:
                        end = log['message'].find('\n', start + 1)
                        self.projects_log[branch2][project].append(log['message'][start + 11: end])
                        start = log['message'].find('Change-Id: ', start+1)
                #printProgressBar (current_count, total_length, prefix = 'Checking implicit changes', suffix = '', decimals = 1, length = 100, fill = '||')
                if 'next' not in logs.keys():
                    mod = current_count // total_iteration
                    rem = current_count % total_iteration
                    if rem != 0:
                        current_count = (mod + 1) * total_iteration
                        #printProgressBar (current_count, total_length, prefix = 'Checking implicit changes', suffix = '', decimals = 1, length = 100, fill = '||')
                    break
                next = '/?s=' + logs['next']

    def check_implicit_changes(self, branch2):
        '''
        Checks the implicit changes in the target branch with the misiing changes 
        '''
        if len(self.stage_1_data) > 0:
            self.get_change_ids(branch2)
        for commit in self.stage_1_data:
            if commit['branch'] == branch2 and commit['change_id'] not in self.projects_log[branch2][commit['project']]:
                if commit['subject'].startswith('Revert'):
                    commit['is_revert'] = True
                else:
                    commit['is_revert'] = False
                self.final_data[self.current_gerrit]['changes'].append(commit)
        self.final_data[self.current_gerrit]['changes'] = sorted(self.final_data[self.current_gerrit]['changes'], key = lambda x: x['merge_time'])

    def add_to_final_data(self):
        for commit in self.stage_1_data:
            if commit['subject'].startswith('Revert'):
                commit['is_revert'] = True
            else:
                commit['is_revert'] = False
            self.final_data[self.current_gerrit]['changes'].append(commit)
        self.final_data[self.current_gerrit]['changes'] = sorted(
            self.final_data[self.current_gerrit]['changes'], key = lambda x: x['merge_time'])

    def compare_branches(self, branch1, branch2, primary_gerrit=None, rdk_gerrit=None):
        '''
        Compares the given source and target branch and returns the final data      

        Parameters
        --------------------------------
        branch1: <str>
            source branch
        branch2: <str>
            target branch
        '''
        self.branch1 = branch1
        self.branch2 = branch2
        self.jira=jira_login()
        #print(self.gerrits, '*********Gerrits*******')
        for gerrit in self.gerrits:
            self.current_gerrit = gerrit
            #print('***************Current gerrit************%s*************' % self.current_gerrit)
            self.final_data[gerrit] = {
                'changes': []
            }
            if self.cmd is False:
                if self.current_gerrit == 'primary_gerrit':
                    self.gerrit = gerrit_login(self.current_gerrit, primary_gerrit)
                else:
                    self.gerrit = gerrit_login(self.current_gerrit, rdk_gerrit)
            else:
                self.gerrit = gerrit_login(self.current_gerrit)
            if self.current_gerrit == 'primary_gerrit' and self.is_dev_specific:
                self.all_devices_repos()
            if self.branch2:
                self.projects_log = {branch2:{}}
            else:
                self.projects_log = {branch1: {}}
            self.crossed_start = False
            self.stage_1_data = []
            self.offset = 0
            while(1):
                try:
                    commit_details = self.gerrit.get('/changes/?q=branch:'+self.branch1+'+status:merged\
        &o=CURRENT_REVISION&o=CURRENT_COMMIT&o=MESSAGES&n=100&S='+str(self.offset))
                    no_commits = False
                    if not commit_details:
                        no_commits = True
                    commit_details = [commit for commit in commit_details if (not self.crossed_start) and self.is_in_range(commit)]
                except requests.exceptions.SSLError as e:
                    commit_details = []
                    print('------------SSL Error------------------------------')
                    #print(e)
                    no_commits = True
                    break
                for i in range(len(commit_details)):
                    try:                    
                        #('\n\n###', self.offset+i+1, ' ###')
                        if self.is_dev_specific and commit_details[i]['project'] not in self.repos_to_be_checked:
                            print('Skipping commit for the project: ' + commit_details[i]['project'])
                            continue
                        self.commit['change_id'] = commit_details[i]['change_id']
                        self.commit['project'] = commit_details[i]['project']
                        self.commit['subject'] = commit_details[i]['subject']
                        self.commit['urlencoded_project'] = commit_details[i]['project'].replace('/','%2f')
                        self.commit['current_revision'] = commit_details[i]['current_revision']
                        self.commit['commit_msg'] = commit_details[i]['revisions'][self.commit['current_revision']]['commit']['message']
                        #print(self.commit['commit_msg'])
                        issue=re.findall(r'\w+-\d+', self.commit['commit_msg'])
                        
                        for each in issue:
                            try:
                                result=self.jira.issue(each)
                                try:
                                    parent=result.fields.parent.key
                                    index=issue.index(each)
                                    issue[index]=parent
                                except :
                                    pass
                            except:
                                pass

                        
                        self.commit['issues']=list(set(issue))
                        self.commit['merge_time'] = datetime.strptime(commit_details[i]['submitted'].split('.')[0], '%Y-%m-%d %H:%M:%S')
                        if self.branch2:
                            self.commit['branch'] = branch2
                            self.check_in_branch(self.branch2)
                        else:
                            self.commit['branch'] = branch1
                            self.stage_1_data.append(copy.deepcopy(self.commit))
                    except IndexError:
                        self.eob = True
                        break
                if self.eob or no_commits or self.crossed_start:
                    print('Reached End Of Branch')
                    break
                else:
                    self.offset += 100
            if self.is_morty:
                if branch2:
                    branch_str = '%s_morty'%(branch2)
                else:
                    branch_str = '%s_morty' % (branch1)
                self.crossed_start = False
                self.projects_log[branch_str] = {}
                while (1):
                    try:
                        mcommit_details = self.gerrit.get('/changes/?q=branch:' + self.branch1 + '_morty+status:merged\
                        &o=CURRENT_REVISION&o=CURRENT_COMMIT&o=MESSAGES&n=100&S=' + str(self.moffset))
                        mno_commits = False
                        if not mcommit_details:
                            mno_commits = True
                        mcommit_details = [commit for commit in mcommit_details if
                                          (not self.crossed_start) and self.is_in_range(commit)]
                    except requests.exceptions.SSLError as e:
                        mcommit_details = []
                        print('------------SSL Error------------------------------')
                        # print(e)
                        mno_commits = True
                        break
                    for i in range(len(mcommit_details)):
                        try:
                            # if self.is_dev_specific and mcommit_details[i]['project'] not in self.repos_to_be_checked:
                            #     print('Skipping commit for the project: ' + commit_details[i]['project'])
                            #     continue
                            self.commit['change_id'] = mcommit_details[i]['change_id']
                            self.commit['project'] = mcommit_details[i]['project']
                            self.commit['subject'] = mcommit_details[i]['subject']
                            self.commit['urlencoded_project'] = mcommit_details[i]['project'].replace('/', '%2f')
                            self.commit['current_revision'] = mcommit_details[i]['current_revision']
                            self.commit['commit_msg'] = \
                            mcommit_details[i]['revisions'][self.commit['current_revision']]['commit']['message']
                            # print(self.commit['commit_msg'])
                            issue=re.findall(r'\w+-\d+', self.commit['commit_msg'])
                            for each in issue:
                                try:
                                    result=self.jira.issue(each)
                                    try:
                                        parent=result.fields.parent.key
                                        index=issue.index(each)
                                        issue[index]=parent
                                    except :
                                        pass
                                except:
                                    pass
                            self.commit['issues']=list(set(issue))
                            # print(self.commit['issues'])
                            self.commit['merge_time'] = datetime.strptime(mcommit_details[i]['submitted'].split('.')[0],
                                                                          '%Y-%m-%d %H:%M:%S')
                            self.commit['branch'] = branch_str
                            if self.branch2:
                                self.check_in_branch(branch_str)
                            else:
                                self.stage_1_data.append(copy.deepcopy(self.commit))
                        except IndexError:
                            self.meob = True
                            break
                    if self.meob or mno_commits or self.crossed_start:
                        print('Reached End Of Branch')
                        break
                    else:
                        self.moffset += 100
            if self.is_dunfell:
                if branch2:
                    branch_str = '%s_dunfell'%branch2
                else:
                    branch_str = '%s_dunfell' % (branch1)
                self.projects_log[branch_str] = {}
                self.crossed_start = False
                while (1):
                    try:
                        dcommit_details = self.gerrit.get('/changes/?q=branch:' + self.branch1 + '_dunfell+status:merged\
                        &o=CURRENT_REVISION&o=CURRENT_COMMIT&o=MESSAGES&n=100&S=' + str(self.doffset))
                        dno_commits = False
                        #print(dcommit_details)
                        if not dcommit_details:
                            dno_commits = True
                        dcommit_details = [commit for commit in dcommit_details if
                                          (not self.crossed_start) and self.is_in_range(commit)]
                        print(dcommit_details)
                    except requests.exceptions.SSLError as e:
                        mcommit_details = []
                        print('------------SSL Error------------------------------')
                        # print(e)
                        dno_commits = True
                        break
                    for i in range(len(dcommit_details)):
                        try:
                            # if self.is_dev_specific and mcommit_details[i]['project'] not in self.repos_to_be_checked:
                            #     print('Skipping commit for the project: ' + commit_details[i]['project'])
                            #     continue
                            self.commit['change_id'] = dcommit_details[i]['change_id']
                            self.commit['project'] = dcommit_details[i]['project']
                            self.commit['subject'] = dcommit_details[i]['subject']
                            self.commit['urlencoded_project'] = dcommit_details[i]['project'].replace('/', '%2f')
                            self.commit['current_revision'] = dcommit_details[i]['current_revision']
                            self.commit['commit_msg'] = \
                            dcommit_details[i]['revisions'][self.commit['current_revision']]['commit']['message']
                            # print(self.commit['commit_msg'])
                            
                            for each in issue:
                                try:
                                    result=self.jira.issue(each)
                                    try:
                                        parent=result.fields.parent.key
                                        index=issue.index(each)
                                        issue[index]=parent
                                    except :
                                        pass
                                except:
                                    pass
                            self.commit['issues']=list(set(issue))
                            self.commit['merge_time'] = datetime.strptime(dcommit_details[i]['submitted'].split('.')[0],
                                                                          '%Y-%m-%d %H:%M:%S')
                            self.commit['branch'] = branch_str
                            if self.branch2:
                                self.check_in_branch(branch_str)
                            else:
                                self.stage_1_data.append(copy.deepcopy(self.commit))
                        except IndexError:
                            self.deob = True
                            break
                    if self.deob or dno_commits or self.crossed_start:
                        print('Reached End Of Branch')
                        break
                    else:
                        self.doffset += 100
            if self.branch2:
                self.check_implicit_changes(self.branch2)
                if self.is_morty:
                    self.check_implicit_changes('%s_morty'%self.branch2)
                if self.is_dunfell:
                    self.check_implicit_changes('%s_dunfell'%self.branch2)
            else:
                self.add_to_final_data()
        return self.final_data
        # return self.merge_pending

    def write_cell(self, change, sheet, row, change_type, col_width, styles, gerrit):
        if len(str(change['merge_time'])) > col_width[change_type]['merge_time']:
            col_width[change_type]['merge_time'] = len(str(change['merge_time']))
        if len(change['change_id']) > col_width[change_type]['change_id']:
            col_width[change_type]['change_id'] = len(change['change_id'])
        joined_issues = ','.join(change['issues'])
        if len(change['project']) > col_width[change_type]['project']:
            col_width[change_type]['project'] = len(change['project'])
        if len(joined_issues) > col_width[change_type]['issues']:
            col_width[change_type]['issues'] = len(joined_issues)
        if change['is_revert']:
            style = 'revert'
        else:
            style = 'normal'
        # sheet.write(row, 0, change['merge_time'], styles[style]['style_bg'])
        sheet.write(row, 0, str(
            datetime.strptime(change['merge_time'], '%Y-%m-%d %H:%M:%S')+ timedelta(hours=5, minutes=30)),
                    styles[style]['style_bg'])
        sheet.write(row, 1, xlwt.Formula(
            'HYPERLINK("{}","{}")'.format(self.gerrit_urls[gerrit] + '/#/q/' + change['change_id'],
                                          change['change_id'])), styles[style]['style_link'])
        sheet.write(row, 2, change['project'], styles[style]['style_center'])
        sheet.write(row, 3, joined_issues, styles[style]['style_bg'])
        sheet.write(row, 4, 'YES' if change['is_revert'] else 'NO', styles[style]['style_bg'])

    def generate_report(self):
        '''
        Initializes the class instance with the required parameters Generates the final report
        using the final data fetched
        '''
        book = xlwt.Workbook()
        changes_sheet = book.add_sheet('missing_changes')
        style_header = xlwt.easyxf('pattern: pattern solid, fore_colour gray40;' 'font: bold True;')
        styles = {
            'revert': {
                'style_bg': xlwt.easyxf('font: colour red;'),
                'style_center': xlwt.easyxf('alignment: vertical center;' 'font: colour red;'),
                'style_link': xlwt.easyxf('font: colour red, underline single;')
            },
            'normal': {
                'style_bg': xlwt.easyxf(''),
                'style_center': xlwt.easyxf('alignment: vertical center;'),
                'style_link': xlwt.easyxf('font: colour blue, underline single;')
            }
        }
        

        col_width = {
            'changes': {
                'project': 7,
                'merge_time': 10,
                'change_id': 9,
                'issues': 7,
                'revert': 6
            }
        }

        changes_sheet.write(0, 0, 'Merge Time (IST)', style_header)        
        changes_sheet.write(0, 1, 'Change Id', style_header)
        changes_sheet.write(0, 2, 'Project', style_header)
        changes_sheet.write(0, 3, 'Issues', style_header)
        changes_sheet.write(0, 4, 'Revert', style_header)

        gerrit_switch = False
        row = 1
        sheet = changes_sheet
        change_type = 'changes'

        for gerrit in self.final_data.keys():
            morty_changes = []
            dunfell_changes = []
            if gerrit_switch:
                sheet.write(row, 0, '')
                sheet.write(row, 1, '')
                sheet.write(row, 2, '')
                sheet.write(row, 3, '')
                row += 1
            gerrit_switch = True
            for change in self.final_data[gerrit][change_type]:
                if change['branch'].find('_morty') != -1:
                    morty_changes.append(change)
                elif change['branch'].find('_dunfell') != -1:
                    dunfell_changes.append(change)
                else:
                    self.write_cell(change, sheet, row, change_type, col_width, styles, gerrit)
                    row += 1

            if morty_changes:
                row += 3
                sheet.write(row, 0, morty_changes[0]['branch'])
                row += 2
                for change in morty_changes:
                    self.write_cell(change, sheet, row, change_type, col_width, styles, gerrit)
                    row += 1

            if dunfell_changes:
                row += 3
                sheet.write(row, 0, dunfell_changes[0]['branch'])
                row += 2
                for change in dunfell_changes:
                    self.write_cell(change, sheet, row, change_type, col_width, styles, gerrit)
                    row += 1

        for sheet in col_width.keys():
            for col in col_width[sheet].keys(): col_width[sheet][col] = 254 if col_width[sheet][col] > 254 else col_width[sheet][col]

        changes_sheet.col(0).width = int(col_width['changes']['merge_time'] * 256)
        changes_sheet.col(1).width = int(col_width['changes']['change_id'] * 256)
        changes_sheet.col(2).width = int(col_width['changes']['project'] * 256)
        changes_sheet.col(3).width = int((col_width['changes']['issues'] + 1) * 256)
        changes_sheet.col(4).width = int(col_width['changes']['revert'] * 256)

        changes_sheet.set_horz_split_pos(1)
        changes_sheet.set_vert_split_pos(1)
        changes_sheet.panes_frozen = True
        changes_sheet.remove_splits = True
        if self.branch2:
            report_file_name = 'reports/' + self.branch1 + '_' + self.branch2 + '_' + 'missing_report_' \
                               + datetime.now().strftime('%d_%m_%Y_%H_%M_%S') + '.xls'
        else:
            report_file_name = 'reports/' + self.branch1 + '_diffreport_' \
                               + datetime.now().strftime('%d_%m_%Y_%H_%M_%S') + '.xls'
        book.save(self.BASE_DIR+'/'+report_file_name)
        return report_file_name

if __name__ == '__main__':
    args_length = len(sys.argv)
    if args_length != 6 and args_length != 4:
        print('Usage: python branch_comparator.py <branch1> <branch2> <device_specific_flag> [<start_time> <end_time>]')
        sys.exit()
    if args_length == 4:
        start, end = 'NO_START', 'NO_END'
    else:
        start, end = sys.argv[4], sys.argv[5]
    branch_comparison = BranchComparison(start, end)
    if sys.argv[3].lower() == 'true':
        branch_comparison.is_dev_specific = True
    results = branch_comparison.compare_branches(sys.argv[1], sys.argv[2])
    if results:
        issues = []
        for gerrit in results.values():
            for change_list in gerrit.values(): 
                for change in change_list: issues += change['issues']
        for gerrit in results.keys():
            for change_type_key in results[gerrit].keys():
                for i in range(len(results[gerrit][change_type_key])): results[gerrit][change_type_key][i]['merge_time'] = str(results[gerrit][change_type_key][i]['merge_time'])
        jql = 'issue in ('+', '.join(issues)+')'
        print('\nCommit-IDs which are available in ' + sys.argv[1] + ' but NOT available in ' + sys.argv[2])
        print('\n')
        print(jql)
        print('\n')
        print(json.dumps(results, indent=2))
        print('\n')
        if len(branch_comparison.exceptional_repos) > 0:
            print('Below project(s) do not have the target branch ({}):'.format(branch_comparison.branch2))
            print(' ,'.join(branch_comparison.exceptional_repos))
        branch_comparison.generate_report()
    else:
        print('\n*** No merge pending tickets ***\n')
 
