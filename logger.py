import logging
import logging.config

def setup_logging(logging_config):
    logging.config.dictConfig(logging_config)