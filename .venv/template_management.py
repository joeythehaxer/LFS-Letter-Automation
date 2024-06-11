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
        # Logic to pick the appropriate template based on letter status
        if not resident.get('1st letter'):
            self.logger.log('info', 'Selected template1 for the first letter')
            return "template1"  # First letter template
        elif not resident.get('2nd letter'):
            self.logger.log('info', 'Selected template2 for the second letter')
            return "template2"  # Second letter template
        elif not resident.get('3rd letter'):
            self.logger.log('info', 'Selected template3 for the third letter')
            return "template3"  # Third letter template
        self.logger.log('info', 'All letters have been sent')
        return None  # All letters have been sent
