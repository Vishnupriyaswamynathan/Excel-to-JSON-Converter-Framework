import logging

class DataValidator:
    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)

    def validate(self, data):
        errors = []
        # check for empty rows
        for row in data:
            if not any(row.values()):
                errors.append(f"Empty row detected: {row}")
        return errors
