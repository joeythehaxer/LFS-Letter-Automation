from string import Template
from logging import Logger

class LetterGenerator:
    def __init__(self, template_manager, logger):
        self.template_manager = template_manager
        self.logger = logger

    def generate_letter(self, data, template_name):
        self.logger.log('info', f'Generating letter using template: {template_name}')
        template_content = self.template_manager.load_template(template_name)
        template = Template(template_content)
        # Substitute the placeholders with actual data
        letter = template.safe_substitute(data)
        return letter

    def input_data_into_letter(self, letter, data):
        # Additional logic if needed to insert data into letter
        return letter

    def generate_customized_letters(self, data):
        self.logger.log('info', 'Generating customized letters')
        letters = []
        for resident in data:
            template_name = self.template_manager.pick_template(resident)
            if template_name:
                letter = self.generate_letter(resident, template_name)
                letters.append(letter)
        return letters
