import yaml
import logging

class ConfigManager:
    def __init__(self, config_path):
        self.config_path = config_path
        self.config = None
        self.logger = logging.getLogger(__name__)

    def load_config(self):
        try:
            with open(self.config_path, 'r') as file:
                self.config = yaml.safe_load(file)
        except Exception as e:
            self.logger.error(f"Error loading config file: {e}")
        return self.config
