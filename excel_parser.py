import xlrd
import json
import logging
from config_manager import ConfigManager
from data_tracker import DataTracker

class ExcelParser:
    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger(__name__)
        self.tracker = DataTracker()

    def parse(self, file_path):
        try:
            workbook = xlrd.open_workbook(file_path)
            sheet = workbook.sheet_by_index(0)
            data = []
            for row_idx in range(sheet.nrows):
                row_data = {}
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell_value(row_idx, col_idx)
                    row_data[sheet.cell_value(0, col_idx)] = cell_value
                data.append(row_data)
            return data
        except Exception as e:
            self.logger.error(f"Error parsing Excel file: {e}")
            self.tracker.log_error(f"Error parsing Excel file: {e}")
            return None
