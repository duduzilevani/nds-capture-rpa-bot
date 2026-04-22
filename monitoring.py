import json
import socket
import time
import getpass
from typing import Any

import requests


def build_monitoring_event(
    success: int,
    start_time: str,
    end_time: str,
    heartbeat: int,
    unique_id: str,
    index_name: str,
) -> dict[str, Any]:
    user = getpass.getuser()
    machine = socket.gethostname()

    return {
        "host": machine,
        "event": {
            "bot_user": user,
            "unique_id": unique_id,
            "bot_machine": machine,
            "heartbeat": heartbeat,
            "start": start_time,
            "end": end_time,
            "success": success,
        },
        "source": f"http:{index_name}",
        "sourcetype": "_json",
        "index": index_name,
        "time": int(time.time()),
        "fields": {},
    }


def send_monitoring_event(
    monitoring_url: str,
    auth_token: str,
    success: int,
    start_time: str,
    end_time: str,
    heartbeat: int,
    unique_id: str,
    index_name: str,
    verify_ssl: bool = True,
    timeout: int = 30,
) -> requests.Response:
    payload = build_monitoring_event(
        success=success,
        start_time=start_time,
        end_time=end_time,
        heartbeat=heartbeat,
        unique_id=unique_id,
        index_name=index_name,
    )

    headers = {
        "Authorization": f"Bearer {auth_token}",
        "Content-Type": "application/json",
    }

    response = requests.post(
        monitoring_url,
        headers=headers,
        data=json.dumps(payload, sort_keys=True),
        verify=verify_ssl,
        timeout=timeout,
    )
    response.raise_for_status()
    return response