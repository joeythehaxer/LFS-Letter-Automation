import json
import os

DEFAULT_CONFIG_PATH = 'default_config.json'

def load_defaults(config_path=DEFAULT_CONFIG_PATH):
    if os.path.exists(config_path):
        with open(config_path, 'r') as f:
            return json.load(f)
    else:
        raise FileNotFoundError(f"{config_path} not found. Please create it with the necessary configurations.")

config = load_defaults()