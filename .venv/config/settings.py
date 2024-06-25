import json
import os

DEFAULT_CONFIG_PATH = 'default_config.json'

class Settings:
    _cache = None  # Class variable to cache the configuration

    def __init__(self, config_dict):
        self.__dict__.update(config_dict)

    @classmethod
    def validate_config(cls, config):
        required_keys = [
            'USE_GUI', 'LOGGING_ENABLED', 'TEMPLATES_DIR', 'PRINT_SERVER_DIR',
            'HEADER_ROW', 'LOCAL_EXCEL_FILE', 'EXCEL_SHEET_NAME',
            'ADDRESS_COLUMN', 'NAME_COLUMN', 'WORK_ORDER_COLUMN',
            'LETTER_1_COLUMN', 'LETTER_2_COLUMN', 'LETTER_3_COLUMN',
            'REVIEW_COLUMN', 'REVIEW_POSITIVE_VALUE', 'PLACEHOLDERS',
            'TEMPLATE_GROUP1', 'TEMPLATE_GROUP2'
        ]
        missing_keys = [key for key in required_keys if key not in config]
        if missing_keys:
            raise ValueError(f"Missing configuration keys: {missing_keys}")

    @classmethod
    def load_defaults(cls):
        if cls._cache:
            return cls._cache

        if os.path.exists(DEFAULT_CONFIG_PATH):
            with open(DEFAULT_CONFIG_PATH, 'r') as f:
                config = json.load(f)
                cls.validate_config(config)  # Validate configuration
                cls._cache = Settings(config)  # Cache the loaded configuration
                return cls._cache
        else:
            raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")

def load_defaults():
    return Settings.load_defaults()
