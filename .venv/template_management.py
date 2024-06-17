import os
from docx import Document
import json

DEFAULT_CONFIG_PATH = 'default_config.json'


class TemplateManager:
    def __init__(self, templates_dir, logger):
        self.templates_dir = templates_dir
        self.logger = logger
        self.load_defaults()

    def load_defaults(self):
        if os.path.exists(DEFAULT_CONFIG_PATH):
            with open(DEFAULT_CONFIG_PATH, 'r') as f:
                self.default_config = json.load(f)
        else:
            raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")

    def load_template(self, template_name):
        template_path = os.path.join(self.templates_dir, f"{template_name}.docx")
        document = Document(template_path)
        return document

    def pick_template(self, review_value):
        if review_value == self.default_config['REVIEW_POSITIVE_VALUE']:
            return self.default_config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']
        else:
            return self.default_config['TEMPLATE_GROUP2']['LETTER_1_TEMPLATE']
