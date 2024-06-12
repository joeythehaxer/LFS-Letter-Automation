from custom_logging import Logger
import os


class TemplateManager:
    def __init__(self, templates_dir, logger):
        self.templates_dir = templates_dir
        self.logger = logger

    def get_template_path(self, template_name):
        """
        Returns the file path of the specified template.

        :param template_name: Name of the template.
        :return: The file path of the template.
        """
        return os.path.join(self.templates_dir, f"{template_name}.docx")

    def load_template(self, template_name):
        self.logger.log('info', f'Loading template: {template_name}')
        # Logic to load a specific template
        template_path = self.get_template_path(template_name)
        with open(template_path, 'rb') as file:
            template = file.read()
        return template

    def pick_template(self, resident):
        """
        Picks the appropriate template based on the value in the review 1 column
        and the letter status columns.

        :param resident: Dictionary containing resident data.
        :return: The name of the template to use.
        """
        review_value = resident.get('review 1', None)
        if review_value == 'CertainValue':  # Replace with the actual value you're checking for
            if not resident.get(config.LETTER_1_COLUMN):
                self.logger.log('info', 'Selected template1_group1 for the first letter')
                return "template1_group1"
            elif not resident.get(config.LETTER_2_COLUMN):
                self.logger.log('info', 'Selected template2_group1 for the second letter')
                return "template2_group1"
            elif not resident.get(config.LETTER_3_COLUMN):
                self.logger.log('info', 'Selected template3_group1 for the third letter')
                return "template3_group1"
        else:
            if not resident.get(config.LETTER_1_COLUMN):
                self.logger.log('info', 'Selected template1_group2 for the first letter')
                return "template1_group2"
            elif not resident.get(config.LETTER_2_COLUMN):
                self.logger.log('info', 'Selected template2_group2 for the second letter')
                return "template2_group2"
            elif not resident.get(config.LETTER_3_COLUMN):
                self.logger.log('info', 'Selected template3_group2 for the third letter')
                return "template3_group2"

        self.logger.log('info', 'All letters have been sent')
        return None  # All letters have been sent
