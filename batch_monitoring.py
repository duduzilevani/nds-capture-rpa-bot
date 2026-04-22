#splunk-batch.py
import json
import requests
from datetime import datetime

class SplunkBatchLogger:
    def __init__(self, splunk, batch_size=10, index="acoe_bot_events"):
        self.splunk = splunk
        self.batch_size = batch_size
        self.index = index
        self.buffer = []

    def add_event(self, status, start, finish, unique_id, heartbeat=1):
        """Add event to buffer and flush when batch limit is reached"""
        event = {
            "status": status,
            "start": start,
            "finish": finish,
            "unique_id": unique_id,
            "index": self.index,
            "heartbeat": heartbeat,
        }
        self.buffer.append(event)
        if len(self.buffer) >= self.batch_size:
            self.flush()

    def flush(self):
        """Push all buffered events to Splunk"""
        if not self.buffer:
            return
        try:
            for e in self.buffer:
                self.splunk.add_event_to_splunk(
                    e["status"],
                    e["start"],
                    finish=e["finish"],
                    unique_id=e["unique_id"],
                    index=e["index"],
                    heartbeat=e["heartbeat"]
                )
            self.buffer.clear()
        except requests.exceptions.ReadTimeout:
            print("Splunk timeout, will retry later...")
        except Exception as e:
            print(f"Failed to push events to Splunk: {e}")
            self.save_to_file()

    def save_to_file(self, filename="splunk_buffer.json"):
        """Save buffer to disk in case Splunk is unavailable"""
        with open(filename, "w") as f:
            json.dump(self.buffer, f, indent=4)
        print(f"📝 Buffered {len(self.buffer)} events saved locally.")
