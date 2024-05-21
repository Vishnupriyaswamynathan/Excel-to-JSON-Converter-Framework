import os
import pandas as pd
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
            includes = self.config['input'].get('includes', [])
            excludes = self.config['input'].get('excludes', [])
            output_dir = self.config['output']['directory']
            
            # Read all sheet names from the Excel file
            xls = pd.ExcelFile(file_path)
            all_sheets = xls.sheet_names
            
            # Determine sheets to parse based on includes and excludes
            if includes:
                sheets_to_parse = [sheet for sheet in includes if sheet not in excludes]
            else:
                sheets_to_parse = [sheet for sheet in all_sheets if sheet not in excludes]

            base_file_name = os.path.splitext(os.path.basename(file_path))[0]
            
            for sheet_name in sheets_to_parse:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                data = df.to_dict(orient='records')
                
                output_file_name = f"{base_file_name}_{sheet_name}.json"
                output_file_path = os.path.join(output_dir, output_file_name)
                
                with open(output_file_path, 'w') as json_file:
                    json.dump(data, json_file, indent=4)
            
            self.logger.info(f"Successfully parsed and saved data for sheets: {sheets_to_parse}")
        except Exception as e:
            self.logger.error(f"Error parsing Excel file: {e}")
            self.tracker.log_error(f"Error parsing Excel file: {e}")
            return None
