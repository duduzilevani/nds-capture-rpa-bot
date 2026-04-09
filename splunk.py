import datetime
import requests
from datetime import datetime
import json
import socket
import getpass
import time
import urllib3
from configurations import Configurations
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def add_event_to_splunk(passed, start, finish, heartbeat, unique_id, index):
    json_event = splunk_object(passed, start, finish, heartbeat, unique_id, index)
    auth = str(Configurations('NDS', 'auth').read_config())
    url = str(Configurations('NDS', 'splunk-url').read_config())
    authHeader = {'Authorization': 'Splunk ' + auth}
    r = requests.post(url, headers=authHeader, data=json.dumps(json_event, sort_keys=True), verify=False)



def add_event_to_splunk(passed, start, finish, heartbeat, unique_id, index):
    json_event = splunk_object(passed, start, finish, heartbeat, unique_id, index)
    aauth = str(Configurations('NDS', 'auth').read_config())
    url = str(Configurations('NDS', 'splunk-url').read_config())
    authHeader = {'Authorization': 'Splunk ' + auth}
    r = requests.post(url, headers=authHeader, data=json.dumps(json_event, sort_keys=True), verify=False)


def splunk_object(passed, start, finish, heartbeat, unique_id, index):
    index = index  # (Always the same non prod, it will change once in prod)
    global uniqueID
    uniqueID = unique_id  # (Obtain your unique id from the RDA Tracker)
    robot_user = getpass.getuser()
    bot_machine = socket.gethostname()
    host = bot_machine
    json_object = {"host": host,
    "event": {"bot_user": robot_user, "unique_id": uniqueID,
    "bot_machine": bot_machine,
    "heartbeat": heartbeat,
    "start": start,
    "end": finish,
    "passed": passed},
    "source": "http:" + index, "sourcetype": "_json", "index": index,
    "time": int(time.time()),
    "fields": {}}
    return json_object

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def splunk_object(passed, start, finish, heartbeat, unique_id, index):
    index = index  # (Always the same non prod, it will change once in prod)
    global uniqueID
    uniqueID = unique_id  # (Obtain your unique id from the RDA Tracker)
    robot_user = getpass.getuser()
    bot_machine = socket.gethostname()
    host = bot_machine
    json_object = {"host": host,
                   "event": {"bot_user": robot_user, "unique_id": uniqueID,
                             "bot_machine": bot_machine,
                             "heartbeat": heartbeat,
                             "start": start,
                             "end": finish,
                             "passed": passed},
                   "source": "http:" + index, "sourcetype": "_json", "index": index,
                   "time": int(time.time()),
                   "fields": {}}
    return json_object
