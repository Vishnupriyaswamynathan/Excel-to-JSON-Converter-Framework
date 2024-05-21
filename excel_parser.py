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
    
    def get_sheets_to_parse(self, all_sheets, includes, excludes):
        if includes:
            sheets_to_parse = [sheet for sheet in includes if sheet not in excludes]
        else:
            sheets_to_parse = [sheet for sheet in all_sheets if sheet not in excludes]
        return sheets_to_parse
    
    def parse_sheet(self, file_path, sheet_name):
        chunk_size = self.config['chunk'].get('size')
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        any_lines_executed = False
        
        df.dropna(axis=0, how='all', inplace=True)  # Drop rows where all values are missing
        df.dropna(axis=1, how='all', inplace=True)

        non_empty_df = self.drop_empty_rows(df,file_path,sheet_name)
        if non_empty_df.empty:
            self.logger.info(f"No data found in sheet '{sheet_name}'")
            return None
        data = self.convert_df_to_dict(non_empty_df)
        return data

    def drop_empty_rows(self, df,file_path,sheet_name):
        non_empty_df = df.dropna(how='all')        
        header_row_index = non_empty_df.first_valid_index()
        # setting the first non-empty row as the header
        df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=range(header_row_index), header=0)    
        # Drop any completely empty columns
        df = df.dropna(axis=1, how='all')    
        return df 
    
    def convert_df_to_dict(self, df):
        df.fillna("", inplace=True)
        data = df.to_dict(orient='records')
        data = self._convert_timestamps(data)
        return data
    
    def save_data_to_json(self, data, output_file_path):
        with open(output_file_path, 'w') as json_file:
            json.dump(data, json_file, indent=4)
               
    def handle_error(self, error_message):
        self.logger.error(error_message)
        self.tracker.log_error(error_message)
        
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
                sheets_to_parse = self.get_sheets_to_parse(all_sheets, includes, excludes)
                base_file_name = os.path.splitext(os.path.basename(file_path))[0]
                
                for sheet_name in sheets_to_parse:
                    try:
                        data = self.parse_sheet(file_path, sheet_name)
                        if data:                        
                            output_file_name = f"{base_file_name}_{sheet_name}.json"
                            output_file_path = os.path.join(output_dir, output_file_name)                        
                            self.save_data_to_json(data, output_file_path)
                            self.logger.info(f"Successfully parsed and saved data for file {file_name}: {sheet_name}")                  
                    except Exception as sheet_error:
                        error_message = f"Error parsing sheet '{sheet_name}' in file '{file_name}': {sheet_error}"
                        self.handle_error(error_message)
                        continue    # Continue with the next sheet
            except Exception as e:
                error_message = f"Error reading Excel file '{file_name}': {file_error}"
                self.handle_error(error_message)
                continue            # Continue with the next file

    def _convert_timestamps(self, data):
        for record in data:
            for key, value in record.items():
                if isinstance(value, pd.Timestamp):
                    record[key] = value.isoformat()
        return data