import os
import config
from custom_logging import Logger

class TemplateManager:
    def __init__(self, templates_dir, logger):
        self.templates_dir = templates_dir
        self.logger = logger

    def get_template_path(self, template_name):
        return os.path.join(self.templates_dir, f"{template_name}.docx")

    def load_template(self, template_name):
        self.logger.log('info', f'Loading template: {template_name}')
        template_path = self.get_template_path(template_name)
        with open(template_path, 'rb') as file:
            template = file.read()
        return template

    def pick_template(self, resident):
        review_value = resident.get(config.REVIEW_COLUMN, None)
        if review_value == config.REVIEW_POSITIVE_VALUE:
            if not resident.get(config.LETTER_1_COLUMN):
                self.logger.log('info', 'Selected template1_group1 for the first letter')
                return config.TEMPLATE_GROUP1['LETTER_1_TEMPLATE']
            elif not resident.get(config.LETTER_2_COLUMN):
                self.logger.log('info', 'Selected template2_group1 for the second letter')
                return config.TEMPLATE_GROUP1['LETTER_2_TEMPLATE']
            elif not resident.get(config.LETTER_3_COLUMN):
                self.logger.log('info', 'Selected template3_group1 for the third letter')
                return config.TEMPLATE_GROUP1['LETTER_3_TEMPLATE']
        else:
            if not resident.get(config.LETTER_1_COLUMN):
                self.logger.log('info', 'Selected template1_group2 for the first letter')
                return config.TEMPLATE_GROUP2['LETTER_1_TEMPLATE']
            elif not resident.get(config.LETTER_2_COLUMN):
                self.logger.log('info', 'Selected template2_group2 for the second letter')
                return config.TEMPLATE_GROUP2['LETTER_2_TEMPLATE']
            elif not resident.get(config.LETTER_3_COLUMN):
                self.logger.log('info', 'Selected template3_group2 for the third letter')
                return config.TEMPLATE_GROUP2['LETTER_3_TEMPLATE']

        self.logger.log('info', 'All letters have been sent')
        return None  # All letters have been sent
