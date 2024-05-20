class DataTracker:
    def __init__(self):
        self.errors = []
        self.warnings = []

    def log_error(self, message):
        self.errors.append(message)

    def log_warning(self, message):
        self.warnings.append(message)

    def get_errors(self):
        return self.errors

    def get_warnings(self):
        return self.warnings
