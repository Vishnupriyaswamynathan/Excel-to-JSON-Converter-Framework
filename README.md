# Excel-to-JSON-Converter-Framework

The Excel to JSON Converter is a Python application designed to parse Excel files with dynamic schemas and convert the data into JSON format. It provides a flexible solution for handling various Excel formats and structures, allowing users to specify input/output configurations and customize data processing options.

## Features

- Dynamic Schema Parsing: Capable of efficiently parsing Excel files of varying sizes, handling diverse data types, special characters, and effectively managing empty rows and columns.
- JSON Output: Converts parsed Excel data into JSON format, preserving the original data structure.
- Data Validation: Implements data validation checks to ensure the integrity and consistency of the converted JSON data.
- Error Tracking: Logs any data validation errors or warnings encountered during the conversion process for analysis.
- Configurability: Configurable through a YAML configuration file, enabling users to customize input/output paths, chunk size and logging options.
- Error Handling: Implemented error handling to handle errors during file parsing and JSON conversion, logging the encountered error within a specific sheet, and continues processing the remaining sheets in a file.
- Unit Testing: Unit tests to ensure the functionality of the converter across various scenarios.

## Installation

- Clone the repository:
```bash
git clone https://github.com/your_username/excel-to-json-converter.git
cd excel-to-json-converter
```

- Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

- Customize the configuration settings in `config.yaml` according to your requirements.
- Place your Excel files in the input directory specified in the configuration.
- Run the converter using the below command in terminal.
```bash
python main.py 
```
- The converter will parse the Excel files, perform data validation, and generate corresponding JSON files in the output directory specified in the configuration(config.yaml).
- In the current configuration, input files are expected to be located in the 'data' folder, while output files are generated in the same directory. However, this setup can be customized within the 'config.yaml' file, allowing users to specify different input and output directories. Additionally, the output file name will follow the format of 'inputfilename_sheetname.json'.

## Configuration

The converter's behavior can be customized through the configuration file `config.yaml`. Below are the configurable options:

- Logging: Configure logging settings such as log level, format, and output destination.
- Input: Specify the directory containing input Excel files (`excel_directory`) and include/exclude specific sheets for parsing.
- Output: Define the directory for saving JSON files (`directory`).
- Chunk Size: Set the chunk size for processing large datasets.



