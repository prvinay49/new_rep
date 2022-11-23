import json
import os
import copy
import requests
from datetime import datetime,timedelta
from pytz import timezone


from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS,cross_origin
from werkzeug.http import parse_authorization_header

from distutils.version import LooseVersion

from branch_comparator import BranchComparison
from release_comparator import ReleaseComparison
from rmGerritUtils import gerrit_login

app = Flask(__name__)
cors = CORS(app, support_credentials=True)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOWNLOAD_DIRECTORY = BASE_DIR + '/'

@app.route("/files/<path:path>")
def get_file(path):
    """Download a file."""
    return send_from_directory(DOWNLOAD_DIRECTORY, path, as_attachment=True)

@app.route("/getdevicedetails", methods=['GET'])
def get_device_details():
    device_details = {}
    with open(BASE_DIR + '/config/devices.json', 'r') as device_file:
        device_details = json.load(device_file)
    return device_details

@app.route("/get_project_map", methods=['GET'])
def get_project_map():
    map = {}
    with open(BASE_DIR + '/config/device_project_map.json', 'r') as map_file:
        map = json.load(map_file)
    return jsonify(map)

@app.route("/update_device_details", methods=['POST'])
def update_device_details():
    if request.method == 'POST':
        data = request.get_json()
        device_list = [v for k, v in data.items()]

        with open(BASE_DIR + '/config/manifests.json', 'r') as manifest_file:
            manifest_json = json.load(manifest_file)
        man_bak = copy.deepcopy(manifest_json)
        manifest_keys = man_bak.keys()

        for v in device_list:
            for d in v:
                if d not in manifest_keys:
                    print('Removed device:', d)
                    del man_bak[d]

        with open(BASE_DIR + '/config/manifests.json', 'w') as manifest_file:
            json.dump(man_bak, manifest_file, indent=4)

        with open(BASE_DIR + '/config/devices.json', 'w') as device_file:
            json.dump(data, device_file, indent=4)

        return jsonify({
            'status': 'success'
        })

@app.route("/update_project_map", methods=['POST'])
def update_project_map():
    if request.method == 'POST':
        data = request.get_json()
        with open(BASE_DIR + '/config/device_project_map.json', 'w') as map_file:
            json.dump(data, map_file, indent=4)
        return jsonify({
            'status': 'success'
        })


@app.route("/compare_branch", methods=['POST'])
def home():
    req_data = request.get_json()

    start = req_data['start']
    end = req_data['end']
    source = req_data['source']
    target = req_data['target']

    device_type = req_data['device_type']
    devices = req_data['devices']
    device_specific_type = req_data['device_specific_type']
    timezone_val = req_data['timezone_val']
    gerrit_option = req_data['gerrit_option']

    auth_header = request.headers.get("Authorization")
    rdk_central = request.headers.get("RDK-CENTRAL")
    primary_gerrit = parse_authorization_header(auth_header)
    rdk_gerrit = parse_authorization_header(rdk_central)

    branch_report_type = req_data['branch_report_type']

    if not start or not end:
        start, end = 'NO_START', 'NO_END'
    else:
        start = datetime.strptime(start, '%Y-%m-%dT%H:%M:%S.%fZ')
        end = datetime.strptime(end, '%Y-%m-%dT%H:%M:%S.%fZ')

        start = start.replace(tzinfo=timezone('UTC'))
        end = end.replace(tzinfo=timezone('UTC'))

        if timezone_val == 'pst':
            start = start.astimezone(timezone('US/Pacific'))
            end = end.astimezone(timezone('US/Pacific'))
        elif timezone == 'ist':
            start = start.astimezone(timezone('Asia/Kolkata'))
            end = end.astimezone(timezone('Asia/Kolkata'))
        elif timezone_val == 'est':
            start = start.astimezone(timezone('US/Eastern'))
            end = end.astimezone(timezone('US/Eastern'))


    print(req_data)
    print('---------------------------')

    branch_comparison = BranchComparison(start, end, False, timezone_val)

    branch_comparison.is_dev_specific = True
    branch_comparison.devices = []
    for k,v in devices.items():
        if v:
            branch_comparison.devices += [u for u,w in v.items() if w is True]

    # with open(BASE_DIR + '/config/devices.json', 'r') as devices_file:
    #     all_devices = json.load(devices_file)
    #
    # for device in all_devices.values():
    #     branch_comparison.devices += device
    #
    print('------------------------------------------')
    print(branch_comparison.devices)

    if gerrit_option == 'ccp':
        branch_comparison.gerrits = ['primary_gerrit']
    elif gerrit_option == 'rdk':
        branch_comparison.gerrits = ['rdk_gerrit']

    try:
        results = {}
        for key in results.keys():
            del results[key]
        results = branch_comparison.compare_branches(source, target, primary_gerrit, rdk_gerrit)
    except requests.exceptions.HTTPError as e:
        print("-------------------Exception------------------------------")
        print(e)
        print('--------------------*************--------------------------')
        return {'error': "Invalid Gerrit username or password!"}
    except Exception as  e:
        print ("-------------------Exception------------------------------")
        print (e)
        print('--------------------*************--------------------------')
        return {'error': str(e)}
    if results:
        print('--------------------Results-------------------------')
        print(results, type(results))
        if results.get('report_file_name'):
            del results['report_file_name']
        issues = []
        for gerrit in results.values():
            for change_list in gerrit.values():
                for change in change_list: issues += change['issues']
        for gerrit in results.keys():
            for change_type_key in results[gerrit].keys():
                for i in range(len(results[gerrit][change_type_key])): results[gerrit][change_type_key][i][
                    'merge_time'] = str(results[gerrit][change_type_key][i]['merge_time'])
        jql = 'issue in (' + ', '.join(issues) + ')'
        # print('\nCommit-IDs which are available in ' + source + ' but NOT available in ' + target)
        # print('\n')
        # print(jql)
        # print('\n')
        # print(json.dumps(results, indent=2))
        # print('\n')
        if len(branch_comparison.exceptional_repos) > 0:
            print('Below project(s) do not have the target branch ({}):'.format(branch_comparison.branch2))
            #print(' ,'.join(branch_comparison.exceptional_repos))
        report_file_name = branch_comparison.generate_report()
        results['report_file_name'] = report_file_name
    else:
        print('\n*** No merge pending tickets ***\n')

    if  gerrit_option == 'ccp':
        results['rdk_gerrit'] = {'changes': []}
    elif gerrit_option == 'rdk':
        results['primary_gerrit'] = {'changes': []}

    return jsonify(results)

@app.route("/add_device", methods=['POST'])
def add_deivce():
    if request.method == 'POST':
        input_data = request.get_json()
        with open(BASE_DIR + '/config/devices.json', 'r') as device_file:
            device_json = json.load(device_file)
            dev_bak = copy.deepcopy(device_json)
        # print(input_data)
        if input_data['displayName'].strip() in dev_bak[input_data['type']]:
            return jsonify({
                'status': False,
                'message': 'Device already exist!'})
        with open(BASE_DIR + '/config/manifests.json', 'r') as manifest_file:
            manifest_json = json.load(manifest_file)
            manifest_bak = copy.deepcopy(manifest_json)
        if input_data['displayName'].strip() in manifest_bak.keys():
            return jsonify({
                'status': False,
                'message': 'Device already exist!'})

        dev_bak[input_data['type']].append(input_data['displayName'].strip())
        manifest_bak[input_data['displayName'].strip()] = {
            "project":input_data['project'],
            "manifest_file": input_data['manifest_file'],
            "model_name": input_data['model_name']
        }

        with open(BASE_DIR + '/config/devices.json', 'w') as device_file:
            json.dump(dev_bak, device_file, indent=4)

        with open(BASE_DIR + '/config/manifests.json', 'w') as manifest_file:
            json.dump(manifest_bak, manifest_file, indent=4)

        return jsonify({
            'status': True,
            'message': 'Created Successfully!'})

@app.route("/get_device_manifest_detail", methods=['GET'])
def get_device_manifest_detail():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    with open(BASE_DIR + '/config/manifests.json', 'r') as manifest_file:
        manifests = json.load(manifest_file)

    device_name = request.args.get('device_name')

    try:
       return jsonify({
           'status': 'success',
           'data': manifests[device_name]})
    except KeyError:
        return {"error" : "Invalid device name"}



@app.route("/get_tags", methods=['GET'])
def get_release_tags():
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    with open(BASE_DIR + '/config/manifests.json', 'r') as manifest_file:
        manifests = json.load(manifest_file)

    device_name = request.args.get('device_name')

    try:
        manifest_project = manifests[device_name]['project']
        model_name = manifests[device_name]['model_name']
    except KeyError:
        return {"error" : "Invalid device name"}

    url = 'projects/%s/tags?m=%s' % (manifest_project.replace('/', '%2F'), model_name)
    #print(url)
    gerrit = gerrit_login('primary_gerrit')
    res = gerrit.get(url)
    tags = []
    for r in res:
        tag = r['ref'].replace('refs/tags/', '')
        tags.append(tag)

    return jsonify(tags)



@app.route("/compare_release", methods=['POST'])
def compare_release_tags():
    req_data = request.get_json()
    print(req_data)
    project_name = req_data.get('project_name')
    manifest_file = req_data.get('manifest_file')
    release_report_type = req_data.get('release_report_type')
    selected_device_release = req_data.get('selected_device_release')
    print(selected_device_release)

    source_release_tag = req_data['source_release_tag']
    target_release_tag = req_data['target_release_tag']

    gerrit_option = req_data['gerrit_option']

    auth_header = request.headers.get("Authorization")
    rdk_central = request.headers.get("RDK-CENTRAL")
    primary_gerrit = parse_authorization_header(auth_header)
    rdk_gerrit = parse_authorization_header(rdk_central)

    try:
        source_release_no = source_release_tag.split('_')[-1]
        target_release_no = target_release_tag.split('_')[-1]

        source_model = "".join(source_release_tag.split('_')[0:-1])
        target_model = "".join(target_release_tag.split('_')[0:-1])
        if source_model == '':
            raise ValueError

        if LooseVersion(source_release_no) > LooseVersion(target_release_no):
            return {'error': 'Please check your Input!!! '
                             'Source release number "%s" is greater than target release number "%s"'
                             %(source_release_no, target_release_no)}

    except ValueError as e:
        return {'error': 'Invalid source/target release tag, Please check your input! '}

    if source_model != target_model:
        return {'error': 'Please check your Input!!! Source tag "%s" not same as ' \
               'target tag "%s"' % (source_model, target_model)}

    start = datetime.now()

    print("Start Time:%s" % start)

    if release_report_type == 'missing':
        diff_report = False
    else:
        source_release_tag, target_release_tag = target_release_tag, source_release_tag
        source_release_no, target_release_no = target_release_no, source_release_no
        diff_report = True
        release_comparison = ReleaseComparison(source_release_no, target_release_no,
                                           source_release_tag, target_release_tag,
                                           selected_device_release,
                                           project_name, manifest_file, diff_report)

    if gerrit_option == 'ccp':
        release_comparison.gerrits = ['primary_gerrit']
    elif gerrit_option == 'rdk':
        release_comparison.gerrits = ['rdk_gerrit']

    try:
        results = {}
        for key in results.keys():
            del results[key]
        results = release_comparison.compare_relase_tags(primary_gerrit, rdk_gerrit)
    except requests.exceptions.HTTPError as e:
        print("-------------------Exception------------------------------")
        print(e)
        print('--------------------*************--------------------------')
        return {'error': "Invalid Gerrit Username or Password!"}
    except Exception as  e:
        print ("-------------------Exception------------------------------")
        print (e)
        raise e
        print('--------------------*************--------------------------')
        return {'error': 'Unable to get report, Please contact support'}

    if results:
        if results.get('report_file_name'):
            del results['report_file_name']
        issues = []
        for res in results.values():
            for change_list in res.values():
                for change in change_list: issues += change['issues']
        for gerrit in results.keys():
            for change_type_key in results[gerrit].keys():
                for i in range(len(results[gerrit][change_type_key])):
                    results[gerrit][change_type_key][i]['merge_time'] = str(
                        results[gerrit][change_type_key][i]['merge_time'])
        jql = 'issue in (' + ', '.join(issues) + ')'
        print('\nCommit-IDs which are available in %s  but NOT available in  %s'
              %(source_release_tag, target_release_tag))
        print('\n')
        print(jql)
        print('\n')
        #print(json.dumps(results, indent=2))
        print('\n')
        if len(release_comparison.exceptional_repos) > 0:
            print('Below project(s) do not have the check point '
                  'release number ({0}), ({1}):'.format(source_release_no, target_release_no))
            print(' ,'.join(release_comparison.exceptional_repos))
        report_file_name = release_comparison.generate_report()
        results['report_file_name'] = report_file_name
    else:
        print('\n*** No changes ***\n')

    end = datetime.now()

    print("Total Time:%s" % (end - start))

    if  gerrit_option == 'ccp':
        results['rdk_gerrit'] = {'changes': []}
    elif gerrit_option == 'rdk':
        results['primary_gerrit'] = {'changes': []}

    return jsonify(results)


if __name__ == "__main__":
    app.run(host='0.0.0.0', debug=True)
