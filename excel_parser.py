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

    def parse(self):
        # Get input directory from configuration
        excel_directory = self.config['input']['excel_directory']
        # List all files in the input directory
        xlsx_files = [file for file in os.listdir(excel_directory) if file.endswith('.xlsx')]
        for file_name in xlsx_files:
            file_path = os.path.join(excel_directory, file_name)
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
                    try:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)
                        data = df.to_dict(orient='records')
                        data = self._convert_timestamps(data)
                    
                        output_file_name = f"{base_file_name}_{sheet_name}.json"
                        output_file_path = os.path.join(output_dir, output_file_name)
                        
                        with open(output_file_path, 'w') as json_file:
                            json.dump(data, json_file, indent=4)                        
                    except Exception as sheet_error:
                        error_message = f"Error parsing sheet '{sheet_name}' in file '{file_name}': {sheet_error}"
                        self.logger.error(error_message)
                        self.tracker.log_error(error_message)
                        continue    # Continue with the next sheet
                self.logger.info(f"Successfully parsed and saved data for sheets: {sheets_to_parse}")
            except Exception as e:
                error_message = f"Error reading Excel file '{file_name}': {file_error}"
                self.logger.error(error_message)
                self.tracker.log_error(error_message)
                continue            # Continue with the next file

    def _convert_timestamps(self, data):
        for record in data:
            for key, value in record.items():
                if isinstance(value, pd.Timestamp):
                    record[key] = value.isoformat()
        return data