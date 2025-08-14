import logging
import os
import sys
from logging.handlers import RotatingFileHandler

def setup_logger(name, log_level=logging.INFO):
    """
    Set up a logger with different behavior for development vs production
    
    Args:
        name (str): Logger name
        log_level: Logging level (default: INFO)
        
    Returns:
        logging.Logger: Configured logger
    """
    # Create logger
    logger = logging.getLogger(name)
    
    # Check if we're running as a frozen executable (production)
    is_frozen = getattr(sys, 'frozen', False)
    
    if is_frozen:
        # Production mode - minimal logging, no files
        logger.setLevel(logging.ERROR)  # Only log errors
        
        # Only add console handler for critical errors
        console_handler = logging.StreamHandler(sys.stderr)
        console_handler.setLevel(logging.ERROR)
        
        # Simple formatter for production
        formatter = logging.Formatter("Error: %(message)s")
        console_handler.setFormatter(formatter)
        
        logger.addHandler(console_handler)
        
    else:
        # Development mode - full logging with files
        logger.setLevel(log_level)
        
        # Create logs directory if needed
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        
        # Create file handler
        log_file = os.path.join(log_dir, f"{name}.log")
        file_handler = RotatingFileHandler(
            log_file, maxBytes=5*1024*1024, backupCount=3
        )
        file_handler.setLevel(log_level)
        
        # Create console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(log_level)
        
        # Create formatter
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
        
        # Add formatter to handlers
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # Add handlers to logger
        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
    
    return logger


def get_simple_logger(name):
    """
    Get a simple logger that does nothing in production mode
    For cases where you want to completely disable logging
    """
    logger = logging.getLogger(name)
    
    # Check if we're running as a frozen executable (production)
    is_frozen = getattr(sys, 'frozen', False)
    
    if is_frozen:
        # Production mode - create a null logger
        logger.setLevel(logging.CRITICAL + 1)  # Higher than any real log level
        logger.addHandler(logging.NullHandler())
    else:
        # Development mode - normal logging
        return setup_logger(name)
    
    return logger