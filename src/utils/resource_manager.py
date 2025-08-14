import os
import sys
import json
from PyQt5.QtGui import QPixmap

class ResourceManager:
    def __init__(self):
        # Determine resource base path
        self.base_path = self._get_resource_base_path()
        self.logger = self._get_temp_logger()
        
    def _get_resource_base_path(self):
        """Determine the base path for resources"""
        # Try relative path first
        relative_path = os.path.join(os.path.dirname(__file__), "..", "..", "resources")
        if os.path.exists(relative_path):
            return relative_path
        
        # Try absolute path for packaged app
        app_path = os.path.dirname(os.path.abspath(sys.argv[0]))
        packaged_path = os.path.join(app_path, "resources")
        if os.path.exists(packaged_path):
            return packaged_path
            
        # Fallback to current directory
        return os.path.join(os.getcwd(), "resources")
    
    def _get_temp_logger(self):
        """Get a simple logger - silent in production mode"""
        import logging
        
        # Check if we're running as a frozen executable (production)
        is_frozen = getattr(sys, 'frozen', False)
        
        logger = logging.getLogger("ResourceManager")
        
        if is_frozen:
            # Production mode - minimal logging
            logger.setLevel(logging.ERROR)
            if not logger.handlers:
                logger.addHandler(logging.NullHandler())
        else:
            # Development mode - normal logging
            logger.setLevel(logging.INFO)
            if not logger.handlers:
                console_handler = logging.StreamHandler()
                console_handler.setLevel(logging.INFO)
                formatter = logging.Formatter("%(name)s - %(levelname)s - %(message)s")
                console_handler.setFormatter(formatter)
                logger.addHandler(console_handler)
        
        return logger
    
    def get_config(self, key_path, default=None):
        """
        Get configuration value from app_config.json
        
        Args:
            key_path (str): Dot-separated path to configuration value
            default: Default value if not found
            
        Returns:
            Config value or default
        """
        config_path = os.path.join(self.base_path, "config", "app_config.json")
        
        if not os.path.exists(config_path):
            self.logger.warning(f"Config file not found: {config_path}")
            return default
            
        try:
            with open(config_path, "r") as f:
                config = json.load(f)
            
            # Traverse dot-separated path
            keys = key_path.split(".")
            value = config
            for key in keys:
                if key in value:
                    value = value[key]
                else:
                    return default
            return value
        except Exception as e:
            self.logger.error(f"Error loading config: {str(e)}")
            return default
    
    def get_icon(self, icon_name):
        """
        Get QPixmap for an icon
        
        Args:
            icon_name (str): Icon file name
            
        Returns:
            QPixmap or None if not found
        """
        icon_path = os.path.join(self.base_path, "icons", icon_name)
        
        if os.path.exists(icon_path):
            try:
                return QPixmap(icon_path)
            except Exception as e:
                self.logger.error(f"Error loading icon {icon_name}: {str(e)}")
                return None
        
        # Only log missing icons in development mode
        if not getattr(sys, 'frozen', False):
            self.logger.warning(f"Icon not found: {icon_path}")
        return None
    
    def get_icon_path(self, icon_name):
        """
        Get absolute path to an icon
        
        Args:
            icon_name (str): Icon file name
            
        Returns:
            str or None if not found
        """
        icon_path = os.path.join(self.base_path, "icons", icon_name)
        return icon_path if os.path.exists(icon_path) else None
    
    def get_stylesheet(self, stylesheet_name):
        """
        Load stylesheet as string
        
        Args:
            stylesheet_name (str): Stylesheet file name
            
        Returns:
            str or None if not found
        """
        stylesheet_path = os.path.join(self.base_path, "styles", stylesheet_name)
        
        if os.path.exists(stylesheet_path):
            try:
                with open(stylesheet_path, "r") as f:
                    return f.read()
            except Exception as e:
                self.logger.error(f"Error loading stylesheet {stylesheet_name}: {str(e)}")
                return None
        
        # Only log missing stylesheets in development mode
        if not getattr(sys, 'frozen', False):
            self.logger.warning(f"Stylesheet not found: {stylesheet_path}")
        return None