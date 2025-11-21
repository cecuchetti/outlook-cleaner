"""Configuration management module"""
import json
from pathlib import Path


def load_config(config_path="config.json"):
    """
    Loads configuration from a JSON file
    
    Args:
        config_path: Path to the configuration file
        
    Returns:
        dict: Dictionary with the configuration
        
    Raises:
        FileNotFoundError: If the configuration file does not exist
        json.JSONDecodeError: If the JSON file is invalid
    """
    config_file = Path(config_path)
    
    if not config_file.exists():
        example_file = Path("config.json.example")
        if example_file.exists():
            raise FileNotFoundError(
                f"[ERROR] Configuration file '{config_path}' not found.\n"
                f"        Copy 'config.json.example' to '{config_path}' and fill in your data."
            )
        else:
            raise FileNotFoundError(
                f"[ERROR] Configuration file '{config_path}' not found.\n"
                f"        Create a '{config_path}' file with your configuration."
            )
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            config = json.load(f)
        
        # Validate basic structure
        required_keys = ['email', 'oauth2', 'cleaning']
        for key in required_keys:
            if key not in config:
                raise ValueError(f"[ERROR] Required key '{key}' is missing in configuration")
        
        return config
    except json.JSONDecodeError as e:
        raise ValueError(f"[ERROR] Error parsing JSON file: {e}")


def get_config_value(config, *keys, default=None):
    """
    Safely gets a configuration value using a key path
    
    Args:
        config: Configuration dictionary
        *keys: Key path (e.g., 'oauth2', 'client_id')
        default: Default value if not found
        
    Returns:
        Configuration value or default
    """
    value = config
    for key in keys:
        if isinstance(value, dict) and key in value:
            value = value[key]
        else:
            return default
    return value

