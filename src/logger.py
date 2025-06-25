import logging
import os
from datetime import datetime

def setup_logger(log_level=logging.INFO, log_to_file=True):
    """Configure and return a logger instance.
    
    Args:
        log_level: Logging level (default: logging.INFO)
        log_to_file: Whether to log to a file (default: True)
        
    Returns:
        Configured logger instance
    """
    # Create logs directory if it doesn't exist
    if log_to_file:
        os.makedirs('logs', exist_ok=True)
    
    # Set up logging
    logger = logging.getLogger('part_checker')
    logger.setLevel(log_level)
    
    # Remove any existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # File handler for detailed logs
    if log_to_file:
        log_filename = f"logs\\part_check_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        file_handler = logging.FileHandler(log_filename)
        file_handler.setLevel(log_level)
        file_format = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_format)
        logger.addHandler(file_handler)
    
    # Console handler for important messages
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_format = logging.Formatter('%(levelname)s: %(message)s')
    console_handler.setFormatter(console_format)
    logger.addHandler(console_handler)
    
    return logger