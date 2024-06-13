import os
from docx import Document
import config

class TemplateManager:
    def __init__(self, templates_dir, logger):
        self.templates_dir = templates_dir
        self.logger = logger

    def load_template(self, template_name):
        template_path = os.path.join(self.templates_dir, f"{template_name}.docx")
        document = Document(template_path)
        return document

    def pick_template(self, review_value):
        if review_value == config.REVIEW_POSITIVE_VALUE:
            return config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE']
        else:
            return config.TEMPLATE_GROUP2['LETTER_1_TEMPLATE']
