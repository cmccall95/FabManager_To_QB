# C:\Users\cmccall\source\repos\cmccall95\Excelsior-Quality-Manager-Dashboard\__init__.py

print("__init__.py accessed")

import logging
from logging.handlers import RotatingFileHandler
import os, sys

def setup_logger():
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    # Create logs directory if it doesn't exist
    log_dir = os.path.join(os.path.dirname(__file__), 'log')
    os.makedirs(log_dir, exist_ok=True)

    # Set up file handler with rotation
    log_file = os.path.join(log_dir, 'app_log.log')
    file_handler = RotatingFileHandler(log_file, maxBytes=1024*1024, backupCount=5)
    file_handler.setLevel(logging.INFO)

    # Set up console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)

    # Create a formatter and add it to the handlers
    formatter = logging.Formatter('%(asctime)s | %(name)s | %(levelname)s | %(message)s | line: %(lineno)d')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    # Add the handlers to the logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    logger.info("Logger configured (__init__)...")
    print("---> Logger configured...")

    return logger

# Set up the logger
logger = setup_logger()
# # Import main after setting up logger
# from .main import main

if __name__ == "__main__":
    logger.info("Application starting")
#     #sys.exit(main())

