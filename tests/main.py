import logging
from config_manager import ConfigManager
from logger import setup_logging
from excel_parser import ExcelParser
from json_converter import JSONConverter
from data_validator import DataValidator

def main():
    # Load configuration
    config_manager = ConfigManager('config.yaml')
    config = config_manager.load_config()
    
    # Setup logging
    setup_logging(config['logging'])
    
    # Parse Excel file
    parser = ExcelParser(config)
    data = parser.parse(config['input']['excel_file_path'])
    
    if data:
        # Validate data
        validator = DataValidator(config)
        errors = validator.validate(data)
        
        if errors:
            for error in errors:
                logging.error(error)
            return
        
        # Convert to JSON
        converter = JSONConverter(config)
        json_data = converter.convert(data)
        
        if json_data:
            with open(config['output']['json_file_path'], 'w') as json_file:
                json_file.write(json_data)
                logging.info(f"Successfully converted to {config['output']['json_file_path']}")

if __name__ == '__main__':
    main()
