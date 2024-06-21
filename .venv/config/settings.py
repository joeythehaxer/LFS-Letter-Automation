import json
import os

DEFAULT_CONFIG_PATH = 'default_config.json'

class Settings:
    def __init__(self, config_dict):
        self.__dict__.update(config_dict)

def load_defaults():
    if os.path.exists(DEFAULT_CONFIG_PATH):
        with open(DEFAULT_CONFIG_PATH, 'r') as f:
            config = json.load(f)
            return Settings(config)
    else:
        raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")
