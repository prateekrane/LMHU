import json
import os
from datetime import datetime

SESSION_PATH = os.path.join("sessions")
os.makedirs(SESSION_PATH, exist_ok=True)

class AppState:
    def __init__(self):
        self.state = {}
        self.session_file = os.path.join(SESSION_PATH, f"session_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json")

    def update(self, key, value):
        self.state[key] = value

    def save(self):
        with open(self.session_file, "w") as f:
            json.dump(self.state, f, indent=2)

    def load(self, filepath):
        with open(filepath, "r") as f:
            self.state = json.load(f)

