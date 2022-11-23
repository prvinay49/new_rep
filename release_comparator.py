
import re, sys, requests, json
from urllib3.exceptions import HTTPError as BaseHTTPError
from rmGerritUtils import *
from datetime import datetime,timedelta
import pdb
import copy
import xlwt
import xmltodict
import multiprocessing
from functools import reduce
from progress_bar import *
import time

from distutils.version import LooseVersion



class ReleaseComparison:
    '''
    Comparator class which compares changes from the given source release tag to the target release tag
    by getting the changes from the Gerrit using APIs.
    Attributes
    --------------------------------
    source release tag: <str>
        source release tag
    final_data: <list>
    repos_to_be_checked: <list>
        list of repos to be checked
    manifests: <dict>
        device manifests details
    Methods
    --------------------------------
    get_repos(device)
        gets all the repos for the given device
    get_change_ids()
        gets the change ids of the release tags
    compare_changes(source release tag change ids, target release tag change ids)
        checks the changes from target to source
    generate_report()
        generates the report from final data
    '''

    commit = {}

    projects_log = {}
    stage_1_data = []
    final_data = {}
    repos_to_be_checked = []
    devices = []
    manifests = None
    exceptional_repos = []
    #gerrits = ['primary_gerrit', 'rdk_gerrit']
    gerrits = ['primary_gerrit']
    gerrit_urls = {
        'primary_gerrit': 'https://gerrit.teamccp.com',
        'rdk_gerrit': 'https://code.rdkcentral.com'
    }

    def __init__(self, source_release_no, target_release_no, source_release_tag,
                 target_release_tag, selected_device_release, project_name=None, manifest_file=None,
                 diff_report=False):
        '''
        Initializes the class instance with the required parameters
        Parameters
        --------------------------------
        start_time: <str>
            start date time
        end_time: <str>
            end date time
        '''

        self.source_release_tag = source_release_tag
        self.target_release_tag = target_release_tag

        self.source_release_no = source_release_no
        self.target_release_no = target_release_no
        self.selected_device_release = selected_device_release

        self.model_full_name = "_".join(self.source_release_tag.split('_')[0:-1])
        self.model_name = "".join(self.source_release_tag.split('_')[0].split('-')[0])

        self.project_input = project_name
        self.manifest_input = manifest_file
        self.diff_report = diff_report

        print(source_release_no, target_release_no, source_release_tag,
               target_release_tag, self.model_full_name, self.model_name)

        self.source_commit_list = {}
        self.target_commit_list = {}
        self.log_error_repos = []
        self.tag_error_repos = []
        self.source_change_id=[]
        self.BASE_DIR = os.path.dirname(os.path.abspath(__file__))

        with open(self.BASE_DIR+'/config/manifests.json', 'r') as manifest_file:
            self.manifests = json.load(manifest_file)

        self.manifest_project = self.manifests[self.selected_device_release]['project']

    def get_min_version(self):
        src_version = LooseVersion(self.source_release_no)
        tgt_version = LooseVersion(self.target_release_no)

        if src_version < tgt_version:
            result = src_version.version
            smallest_version = self.source_release_no
            max_version = self.target_release_no
        else:
            result = tgt_version.version
            smallest_version = self.target_release_no
            max_version = self.source_release_no
        return "".join([str(result[0]), '.', str(result[1]), '.0.0']), smallest_version, max_version

    @staticmethod
    def find_min_version(version1, version2):
        if LooseVersion(version1) < LooseVersion(version2):
            return version1
        return version2

    @staticmethod
    def find_max_version(version1, version2):
        if LooseVersion(version1) > LooseVersion(version2):
            return version1
        return version2

    def find_max_min_version_tag(self):
        if LooseVersion(self.source_release_no) > LooseVersion(self.target_release_no):
            return self.source_release_tag, self.target_release_tag
        return self.target_release_tag, self.source_release_tag

    def get_source_commit_id(self, project):
        stag_api = '/projects/%s/tags?m=%s' % (project, self.source_release_tag)
        stags = self.gerrit.get(stag_api.replace("/","",1))
        stag_versions = [tag['ref'].replace('refs/tags/', '') for tag in stags]
        # print(stag_versions, self.source_release_no)
        return stags[stag_versions.index(self.source_release_tag)]['object']

    def get_repos(self):
        '''
        Gets all the repos of the given device
        Parameters
        --------------------------------
        device: <str>
            device name
        '''
        try:

            source_commit_id = self.get_source_commit_id(self.manifest_project.replace('/', '%2F'))

            if not self.manifest_input:
                manifest_input_file = self.manifests[self.selected_device_release]['manifest_file']
            else:
                manifest_input_file = self.manifest_input

            manifest_api = "/projects/%s/commits/%s/files/%s/content" \
                           % (self.manifest_project.replace('/', '%2F'), source_commit_id,
                              manifest_input_file)

            if self.manifests[self.selected_device_release]['manifest_file'].split('.')[-1].lower() == 'xml':
                manifest = xmltodict.parse(self.gerrit.get(manifest_api))
                for project in manifest['manifest']['project']:
                    if project['@name'] not in self.repos_to_be_checked:
                        self.repos_to_be_checked.append(project['@name'])
            elif self.manifests[self.selected_device_release]['manifest_file'].split('.')[-1].lower() == 'git':
                manifest_data = self.gerrit.get(manifest_api)
                start = manifest_data.find('{', manifest_data.find('\ndeps') + 1)
                end = manifest_data.find('}', manifest_data.find('\ndeps') + 1) + 1
                manifest = json.loads(manifest_data[start:end].replace("'", '"').replace(',\n}', '\n}'))
                for project_url in manifest.values():
                    if 'gerrit.teamccp.com' in project_url:
                        project = '/'.join(project_url.split('gerrit.teamccp.com')[-1].split('@')[0].split('/')[1:])
                        if project not in self.repos_to_be_checked:
                            self.repos_to_be_checked.append(project.replace("/","",1))
        except Exception as e:
            print(e)
            if str(e).find('401 Client Error')!=-1:
                raise e
            raise Exception('Exception occured while getting device repos')

    def get_change_ids(self, project, tag_version, cpoint_commit_id_1, cpoint_commit_id_2, first_iter, to_append):
        '''
        Gets all the change ids of the release tag
        '''

        # print('           Checking commit id match for tag "%s" and append list is "%s"'%(tag_version, to_append))

        iteration_count = 0
        current_count = 0
        total_length = 100
        log_api_template = '/plugins/gitiles/%s/+log/%s' % (project, tag_version)
        
        next = ''
        cbreak = False
        while True:
            # printProgressBar(current_count, total_length, prefix='Checking commit id match', suffix='', decimals=1,
            #                  length=100, fill='█')
            iteration_count += 1
            current_count += 1
            log_api = log_api_template + next
            print(log_api)
            try:
                logs = self.gerrit.get(log_api)

            except requests.exceptions.HTTPError as e:
                print('----unable to get logs----: %s' % e)
                self.log_error_repos.append(project)
                break
            for log in logs['log']:
                
                start = log['message'].find('Change-Id: ')
                end = log['message'].find('\n', start + 1)
                
                if log['message'].startswith('Revert'):
                    if log['message'].find('Revert Revert') != -1:
                        is_revert = False
                    else:
                        is_revert = True
                else:
                    is_revert = False

                more_change_ids = log['message'].split('Change-Id')
                

                # if log['message'].count('Change-Id:') > 1:
                #     print('****************Log message mutiple change ids********************')
                #     print(log['message'])
                #     #pass
                issue_start = 0
                while start != -1:
                    end = log['message'].find('\n', start + 1)
                    change_dict = {}
                    if len(log['message'])<1100:
                        change_dict['change_id'] = log['message'][start + 11: end]
                        change_dict['merge_time'] = log['committer']['time']
                        change_dict['is_revert'] = is_revert
                        change_dict['project'] = project
                        change_dict['author']=log['author']['name']
                        change_dict['commit']=log['commit']
                        issues = re.findall(r'\w+-\d+', log['message'])
                        print("issues"+str(log['message'][start + 11: end]))
                        print(issues)
                        if len(more_change_ids) > 1:
                            print("mords_change_id"+str(log['message'][start + 11: end]))
                            str_q=re.findall(r'\w+-\d+', log['message'][issue_start:end])

                            if str_q:
                                print("uuuuuuuuu")
                                change_dict['issues'] = re.findall(r'\w+-\d+', log['message'][issue_start:end])
                            else:
                                change_dict['issues']=issues
                                print("issues_444")
                                print(change_dict['issues'])
                        else:
                            change_dict['issues'] = issues
                    else:
                        if start<(len(log['message'])//2):
                            pass
                        elif start<((len(log['message'])//2)+(len(log['message'])//3)):
                            pass
                        elif start<((len(log['message'])//2)+(len(log['message'])//3)+50):
                            pass
                        else:
                            change_dict['change_id'] = log['message'][start + 11: end]
                            change_dict['merge_time'] = log['committer']['time']
                            change_dict['is_revert'] = is_revert
                            change_dict['project'] = project
                            change_dict['author']=log['author']['name']
                            change_dict['commit']=log['commit']
                            issues = re.findall(r'\w+-\d+', log['message'])
                            print("isues")
                            print(issues)
                            if len(more_change_ids) > 1:
                                print("mords_change_id"+str(log['message'][start + 11: end]))
                                str_q=re.findall(r'\w+-\d+', log['message'][issue_start:end])

                                if str_q:
                                    print("uuuuuuuuu")
                                    change_dict['issues'] = re.findall(r'\w+-\d+', log['message'][issue_start:end])
                                else:
                                    change_dict['issues']=issues
                                    print("issues_444")
                                    print(change_dict['issues'])
                            else:
                                change_dict['issues'] = issues
                            

                    if project == self.manifest_project:
                        non_stpt_issues = [issue for issue in issues if not issue.startswith('STBT-')]
                        if len(non_stpt_issues) == 0:
                            print('Ignoring manifest project only STBT tickets')
                            break
                    if to_append == 'source':
                        if change_dict:
                            self.source_commit_list[project].append(change_dict)
                    else:
                        self.target_commit_list[project].append(change_dict)

                    start = log['message'].find('Change-Id: ', start + 1)
                    issue_start = end + 1

                if first_iter is True:
                    if log['commit'] == cpoint_commit_id_1 or log['commit'] == cpoint_commit_id_2:
                        cbreak = True
                else:
                    if log['commit'] == cpoint_commit_id_1:
                        cbreak = True
                if cbreak is True:
                    # print(
                    #     '           *********************Check Point Commit id match found "%s"*********************'
                    #     '**************'%log['commit'])
                    # print('           Iteration count "%s"' %iteration_count)
                    check_point = log['commit']
                    return check_point

            # if cbreak is True:
            #     break
            # print('**************************************************************')
            if 'next' not in logs.keys():
                break
            next = '/?s=' + logs['next']

    def get_tags(self, tag_api):
        try:
            if tag_api.startswith("a")or tag_api.startswith("/a"):
                c=tag_api.lstrip("a")
                tags = self.gerrit.get(c)
            else:
                tags = self.gerrit.get(tag_api)
        except requests.exceptions.HTTPError as e:
            # print('Skipping project *****')
            if str(e).find('401 Client Error')!=-1:
                raise e
            # self.exceptional_repos.append(project)
            # print('Error while calling tag api')
            # break
            tags = []
        except requests.exceptions.RetryError as e:
            print('**************Retry Error ***********************')
            tags = []
        except requests.exceptions.ReadTimeout as e:
            print('**************Retry Error ***********************')
            tags = []

        return tags

    def compare_changes(self, project):
        '''
        Checks thechanges in the target change ids with the source changes
        '''

        change_data = []
        target_change_ids = [change['change_id'] for change in self.target_commit_list[project]if change]
        #print( "target_change_ids"+str(target_change_ids))
        source_change_ids = [change['change_id'] for change in self.source_commit_list[project]if change]
        source_commit_id=[change['commit']for change in self.source_commit_list[project]if change]
        #print("source_change_ids"+str(source_change_ids))
        print("commit_id"+str(set(source_commit_id)))
        change_data_1=[]
        for each in set(source_commit_id):
            changes_template='/changes/?q=%s+status:Merged'%(each)
            change=self.gerrit.get(changes_template)
            change_data_1.append(change)

        change_data_by_change_id=[]
        for each in change_data_1:
            if len(each)>1:
                a_single=each[0]
                change_data_by_change_id.append(a_single)
            else:
                change_data_by_change_id.append(each)

        # print('           ***************************************************')
        # print('           Target change ids: %s' %target_change_ids)
        # print('           ***************************************************')
        # # print(           '***************************************************')
        # print('           Source change ids:%s'%source_change_ids)
        # print('           ***************************************************')
        for change_item in self.source_commit_list[project]:
            if change_item['change_id'] not in target_change_ids:
                for each in change_data_by_change_id:
                    if isinstance(each,list):
                        for i in each:
                            if i['change_id']==change_item['change_id']:
                                print('change_id missing %s' % change_item['change_id'])
                                a=i['submitted'].split('.')
                                change_item['merge_time']=datetime.strptime(a[0],"%Y-%m-%d %H:%M:%S")+timedelta(hours=5, minutes=30)
                                change_data.append(change_item)
                            
                    else:
                        if each['change_id']==change_item['change_id']:
                            print('change_id missing %s' % change_item['change_id'])
                            a=each['submitted'].split('.')
                            change_item['merge_time']=datetime.strptime(a[0],"%Y-%m-%d %H:%M:%S")+timedelta(hours=5, minutes=30)
                            change_data.append(change_item)
                             
        
        # self.final_data[self.current_gerrit]['changes'] = sorted(
        #     self.final_data[self.current_gerrit]['changes'],
        #     key = lambda x: datetime.strptime(x['merge_time'], '%a %b %d %H:%M:%S %Y %z'))

        return change_data

    def compare_relase_tags(self, primary_gerrit=None, rdk_gerrit=None):
        '''
        Compares the given source and target release tags and returns the final data
        '''

        manifest_content = self.manifests.get(self.selected_device_release)

        if not manifest_content:
            print('Mapping "%s" not found in manifests json' % self.source_release_tag)
            return

        for gerrit in self.gerrits:
            self.current_gerrit = gerrit
            self.final_data[gerrit] = {
                'changes': []
            }
            if self.current_gerrit == 'primary_gerrit':
                self.gerrit = gerrit_login(self.current_gerrit, primary_gerrit)
            else:
                self.gerrit = gerrit_login(self.current_gerrit, rdk_gerrit)
            self.repos_to_be_checked = []
            if not self.project_input:
                self.get_repos()
            else:
                self.repos_to_be_checked = set(self.project_input.split(","))
            print('*************************Projects***********************************')
            print(len(self.repos_to_be_checked))
            print('********************************************************************')

            min_version_no, smallest_version, max_version = self.get_min_version()

            print('smallest version formatted is "%s" and "%s" is smallest one in "%s",  "%s"'
                  % (min_version_no, smallest_version, self.source_release_no, self.target_release_no))

            # self.repos_to_be_checked = ['rdk/yocto_oe/layers/meta-rdk-oem-tch-brcm-xb6']
            # self.repos_to_be_checked = ['rdk/components/sdk/soc/intel/5.0-gw/watchdog']
            # self.repos_to_be_checked = ['rdk/components/generic/libunpriv/generic']

            self.projects_log = {}
            pool = multiprocessing.Pool(processes=10)

            print(min_version_no, smallest_version, max_version)
            zip_list = [(project, project_count + 1, min_version_no, smallest_version, max_version)
                        for project_count, project in enumerate(self.repos_to_be_checked)]

            results = pool.starmap(self.get_changes_for_project, zip_list)
            # except:
            #     print("exception occured")
            #     time.sleep(5)
            #     self.tag_error_repos.append()
            #print("results:"+ str(results))

            #print(len(results))

            self.final_data[self.current_gerrit]['changes'] = [item for sublist in results for item in sublist]
            self.final_data[self.current_gerrit]['changes'] = sorted(
                self.final_data[self.current_gerrit]['changes'],
                key=lambda x: x['merge_time'])

        return self.final_data

    def get_changes_for_project(self, project, project_count, min_version_no, smallest_version, max_version):
        # print('************', project, min_version_no, smallest_version, max_version)
        # print(multiprocessing.current_process())
        print('--------------Checking for project "%s", "%s" out of "%s" ------------------' % (
            project, project_count, len(self.repos_to_be_checked)))
        change_data = []
        while True:
            # Getting the tag list for min version(stable release) ex:4.2.0.0
            tag_api = '/projects/%s/tags?m=%s_%s' % (project.replace('/', '%2F'),
                                                     self.model_full_name, min_version_no)
            try:                                         
                tags = self.get_tags(tag_api.replace("/","",1))
            except:
                self.tag_error_repos.append(tag_api)
                break
          

            tag_versions = [tag['ref'].split('_')[-1] for tag in tags]
            if tag_versions:
                latest_version = reduce(self.find_max_version, tag_versions)
                c1_version = self.find_min_version(smallest_version, latest_version)
                #print("version_c1:"+str(c1_version))

                # print('           Tag vesions: %s' % tag_versions)
                # print('           Latest version is "%s"' % latest_version)
                # print('           smallest version is "%s"' % smallest_version)
                # print('           Min version of latest and source/target(C1 version):  %s' % c1_version)
                # print('           Source release no:%s, Target release no:%s'
                #       %(self.source_release_no, self.target_release_no))

                if smallest_version in tag_versions:
                    cpoint_commit_id_1 = tags[tag_versions.index(c1_version)]['object']
                    cpoint_commit_id_2 = tags[tag_versions.index(smallest_version)]['object']
                elif max_version in tag_versions:
                    cpoint_commit_id_1 = None
                    cpoint_commit_id_2 = tags[tag_versions.index(max_version)]['object']
                else:
                    stag_api = '/projects/%s/tags?m=%s_%s' % (project.replace('/', '%2F')
                                                              , self.model_full_name, smallest_version)
                    stags = self.get_tags(stag_api.replace("/","",1))
                    stag_versions = [tag['ref'].split('_')[-1] for tag in stags]
                    if stag_versions and smallest_version in stag_versions:
                        cpoint_commit_id_2 = stags[stag_versions.index(smallest_version)]['object']
                        if latest_version == c1_version:
                            cpoint_commit_id_1 = tags[tag_versions.index(c1_version)]['object']
                        else:
                            cpoint_commit_id_1 = stags[stag_versions.index(c1_version)]['object']
                    else:
                        target_tag_api = '/projects/%s/tags?m=%s' % (
                        project.replace('/', '%2F'), self.target_release_tag)
                        target_tags = self.get_tags(target_tag_api.replace("/","",1))
                        target_tag_versions = [tag['ref'].split('_')[-1] for tag in target_tags]
                        cpoint_commit_id_1 = None
                        if target_tag_versions and smallest_version in target_tag_versions:
                            cpoint_commit_id_2 = target_tags[target_tag_versions.index(smallest_version)]['object']
                        else:
                            cpoint_commit_id_2 = None
            else:
                # print('No tags found for stable release: %s' % min_version_no)
                # print("Getting change ids with out check point")
                # print('checking target tag available "%s"' %self.target_release_tag)
                target_tag_api = '/project"s/%s/tags?m=%s' % (project.replace('/', '%2F'), self.target_release_tag)
                tags = self.get_tags(target_tag_api.replace("/","",1))
                tag_versions = [tag['ref'].split('_')[-1] for tag in tags]
                cpoint_commit_id_1 = None
                if tag_versions and smallest_version in tag_versions:
                    cpoint_commit_id_2 = tags[tag_versions.index(smallest_version)]['object']
                else:
                    cpoint_commit_id_2 = None

                # self.exceptional_repos.append(project)
                # print('Skipping project ...')
                # break

                # print('           ------------ cpoint_commit_id 1:', cpoint_commit_id_1)
                # print('           ------------ cpoint_commit_id 2:', cpoint_commit_id_2)

            print('Check point commit id1:%s, check point commit id2:%s ' % (
            cpoint_commit_id_1, cpoint_commit_id_2))

            self.source_commit_list[project] = []
            self.target_commit_list[project] = []

            fmax_verison_tag, fmin_version_tag = self.find_max_min_version_tag()

            to_append1 = 'target'
            to_append2 = 'source'
            if fmax_verison_tag == self.source_release_tag:
                to_append1 = 'source'
                to_append2 = 'target'

            # print('           Max verison is "%s", first getting its change ids' % fmax_verison_tag)

            check_point_commit_id = self.get_change_ids(project, fmax_verison_tag,
                                                        cpoint_commit_id_1, cpoint_commit_id_2,
                                                        True, to_append1)
            print("check point commit id:"+str(check_point_commit_id))
            self.get_change_ids(project, fmin_version_tag, check_point_commit_id,
                                cpoint_commit_id_2, False, to_append2)


            change_data = self.compare_changes(project)
            print('           ---------------Completed the project "%s"-------------' % (project_count))
            break
        # break
        # print('Exiting')
        print(change_data)

        return change_data

    def generate_report(self):
        '''
        Initializes the class instance with the required parameters Generates the final report
        using the final data fetched
        '''
        book = xlwt.Workbook()
        if self.diff_report is True:
            changes_sheet = book.add_sheet('diff_changes')
            report_file_name = 'reports/%s_%s_%s_%s.xls' \
                               % (self.target_release_no,
                                  self.source_release_no,
                                  'diff_report',
                                  datetime.now().strftime('%d_%m_%Y_%H_%M_%S'))
        else:
            changes_sheet = book.add_sheet('missing_changes')
            report_file_name = 'reports/%s_%s_%s_%s.xls' \
                               % (self.source_release_no,
                                  self.target_release_no,
                                  'missing_report',
                                  datetime.now().strftime('%d_%m_%Y_%H_%M_%S'))

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

        changes_sheet.write(0, 0, 'Merge Time', style_header)
        changes_sheet.write(0, 1, 'Change Id', style_header)
        changes_sheet.write(0, 2, 'Project', style_header)
        changes_sheet.write(0, 3, 'Issues', style_header)
        changes_sheet.write(0, 4, 'Revert', style_header)

        gerrit_switch = False
        row = 1
        sheet = changes_sheet
        change_type = 'changes'

        for gerrit in self.final_data.keys():
            if gerrit_switch:
                sheet.write(row, 0, '')
                sheet.write(row, 1, '')
                sheet.write(row, 2, '')
                sheet.write(row, 3, '')
                row += 1
            gerrit_switch = True
            for change in self.final_data[gerrit][change_type]:
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


                sheet.write(row, 0,change['merge_time'], styles[style]['style_bg'])
                sheet.write(row, 1, xlwt.Formula(
                    'HYPERLINK("{}","{}")'.format(self.gerrit_urls[gerrit] + '/#/q/' + change['change_id'],
                                                  change['change_id'])), styles[style]['style_link'])
                sheet.write(row, 2, change['project'], styles[style]['style_center'])
                sheet.write(row, 3, joined_issues, styles[style]['style_bg'])
                sheet.write(row, 4, 'YES' if change['is_revert'] else 'NO', styles[style]['style_bg'])
                row += 1

        for sheet in col_width.keys():
            for col in col_width[sheet].keys(): col_width[sheet][col] = 254 if col_width[sheet][col] > 254 else \
            col_width[sheet][col]

        changes_sheet.col(0).width = int(col_width['changes']['merge_time'] * 256)
        changes_sheet.col(1).width = int(col_width['changes']['change_id'] * 256)
        changes_sheet.col(2).width = int(col_width['changes']['project'] * 256)
        changes_sheet.col(3).width = int((col_width['changes']['issues'] + 1) * 256)
        changes_sheet.col(4).width = int(col_width['changes']['revert'] * 256)

        changes_sheet.set_horz_split_pos(1)
        changes_sheet.set_vert_split_pos(1)
        changes_sheet.panes_frozen = True
        changes_sheet.remove_splits = True



        book.save(self.BASE_DIR+'/'+report_file_name)

        return report_file_name


if __name__ == '__main__':
    args_length = len(sys.argv)
    if args_length not in [3, 4, 5]:
        print('Usage: release_comparator.exe <source_release_tag> <target_release_tag>')
        print('or')
        print('release_comparator.exe <source_release_tag> <target_release_tag> <project_name> <manifest-file>')
        sys.exit()

    # log_file_name = 'logs/%s_%s_logs_%s.txt'%(sys.argv[1], sys.argv[2], datetime.now().strftime('%d_%m_%Y_%H_%M_%S'))
    # sys.stdout = open(log_file_name, 'w')

    project_name = None
    manifest_file = None

    if args_length > 3:
        project_name = sys.argv[3]

    if args_length > 4:
        manifest_file = sys.argv[4]

    try:
        source_release_no = sys.argv[1].split('_')[-1]
        target_release_no = sys.argv[2].split('_')[-1]

        source_model = "".join(sys.argv[1].split('_')[0:-1])
        target_model = "".join(sys.argv[2].split('_')[0:-1])
    except ValueError as e:
        print('Invalid source/target release tag, Please check your input! ')
        sys.exit()

    # if LooseVersion(source_release_no) > LooseVersion(target_release_no):
    #     print('Please check your Input!!! Source release number "%s" is greater than target release number "%s" '
    #     %(source_release_no, target_release_no))
    #     sys.exit()

    # if LooseVersion(source_release_no) > LooseVersion(target_release_no):
    #     diff_report = True

    if source_model != target_model:
        print('Please check your Input!!! Source tag "%s" not same as target tag "%s"' % (source_model, target_model))
        sys.exit()

    start = datetime.now()

    print("Start Time:%s" % start)

    release_comparison = ReleaseComparison(source_release_no, target_release_no,
                                           sys.argv[1], sys.argv[2], project_name, manifest_file)

    results = release_comparison.compare_relase_tags()
    if results:
        issues = []
        for gerrit in results.values():
            for change_list in gerrit.values():
                for change in change_list: issues += change['issues']
        for gerrit in results.keys():
            for change_type_key in results[gerrit].keys():
                for i in range(len(results[gerrit][change_type_key])):
                    results[gerrit][change_type_key][i]['merge_time'] = str(
                        results[gerrit][change_type_key][i]['merge_time'])
        jql = 'issue in (' + ', '.join(issues) + ')'
        print('\nCommit-IDs which are available in ' + sys.argv[1] + ' but NOT available in ' + sys.argv[2])
        print('\n')
        print(jql)
        print('\n')
        print(json.dumps(results, indent=2))
        print('\n')
        if len(release_comparison.exceptional_repos) > 0:
            print('Below project(s) do not have the check point '
                  'release number ({0}), ({1}):'.format(source_release_no, target_release_no))
            print(' ,'.join(release_comparison.exceptional_repos))
        release_comparison.generate_report()
    else:
        print('\n*** No changes ***\n')

    end = datetime.now()

    print("Total Time:%s" % (end - start))
