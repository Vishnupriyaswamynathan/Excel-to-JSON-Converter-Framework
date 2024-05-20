import json
import logging

class JSONConverter:
    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)

    def convert(self, data):
        try:
            json_data = json.dumps(data, indent=4)
            return json_data
        except Exception as e:
            self.logger.error(f"Error converting data to JSON: {e}")
            return None
