import logging
import logging.config

# def setup_logging(config_path):
#     with open(config_path, 'r') as file:
#         config = yaml.safe_load(file)
#         logging.config.dictConfig(config)

def setup_logging(logging_config):
    logging.config.dictConfig(logging_config)