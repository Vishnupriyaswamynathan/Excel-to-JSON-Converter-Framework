.. Excel-to-JSON-Converter-Framework documentation master file, created by
   sphinx-quickstart on Tue May 21 21:31:21 2024.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome to Excel-to-JSON-Converter-Framework's documentation!
=============================================================

.. toctree::
   :maxdepth: 2
   :caption: Contents:



Indices and tables
==================


* :ref:`search`

Installation
------------

To install the Excel to JSON Converter Framework, follow these steps:

1. Clone the repository:

   .. code-block:: bash

      git clone https://github.com/your_username/excel-to-json-converter.git
      cd excel-to-json-converter

2. Install dependencies:

   .. code-block:: bash

      pip install -r requirements.txt


.. _usage:

Usage
-----

To use the Excel to JSON Converter Framework, follow these steps:

1. Customize the configuration settings in `config.yaml` according to your requirements.
2. Place your Excel files in the input directory specified in the configuration.
3. Run the converter using the following command in the terminal:

   .. code-block:: bash

      python main.py

4. The converter will parse the Excel files, perform data validation, and generate corresponding JSON files in the output directory specified in the configuration (`config.yaml`).


.. _configuration:

Configuration
--------------

The Excel to JSON Converter Framework's behavior can be customized through the configuration file `config.yaml`. Below are the configurable options:

- **Logging**: Configure logging settings such as log level, format, and output destination.
- **Input**: Specify the directory containing input Excel files (`excel_directory`) and include/exclude specific sheets for parsing.
- **Output**: Define the directory for saving JSON files (`directory`).
- **Chunk Size**: Set the chunk size for processing large datasets.


.. _file_descriptions:

File Descriptions
-----------------

The Excel to JSON Converter Framework consists of the following files:

- **main.py**: This module contains the main script for running the converter. This file serves as the entry point for the application. It orchestrates the parsing of an Excel file, validation of the data, conversion to JSON format, and logging of errors and informational messages.
- **excel_parser.py**: This module provides functionality to parse Excel files, process sheets, and convert data to JSON format.
- **json_converter.py**: This module converts parsed Excel data into JSON format.
- **data_validator.py**: This module implements data validation checks on input and output data.
- **data_tracker.py**: This module provides a data tracking mechanism to log errors and warnings encountered during data validation.
- **config_manager.py**: This module handles configuration file parsing and settings management.
- **logger.py**: This module handles the configuration and setup of the logging system for the application.
- **config.yaml**: This YAML configuration file specifies various settings for the application, including logging configuration, input and output file paths, and chunk size for processing large datasets.



