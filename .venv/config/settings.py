import json
import os

DEFAULT_CONFIG_PATH = 'default_config.json'

class Settings:
    _cache = None  # Class variable to cache the configuration

    def __init__(self, config_dict):
        self.__dict__.update(config_dict)

    @classmethod
    def validate_config(cls, config):
        required_keys = ['USE_GUI', 'LOGGING_ENABLED', 'TEMPLATES_DIR']  # Add all required keys
        if not all(key in config for key in required_keys):
            missing_keys = [key for key in required_keys if key not in config]
            raise ValueError(f"Missing configuration keys: {missing_keys}")

def load_defaults():
    if Settings._cache:
        return Settings._cache

    if os.path.exists(DEFAULT_CONFIG_PATH):
        with open(DEFAULT_CONFIG_PATH, 'r') as f:
            config = json.load(f)
            Settings.validate_config(config)  # Validate configuration
            Settings._cache = Settings(config)  # Cache the loaded configuration
            return Settings._cache
    else:
        raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")
