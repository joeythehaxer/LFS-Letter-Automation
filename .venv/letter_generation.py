import os
import json
from docx import Document
from openai import OpenAI

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
from custom_logging import Logger

DEFAULT_CONFIG_PATH = 'default_config.json'

# Ensure your OpenAI API key is set in the environment variables

class LetterGenerator:
    def __init__(self, config, logger, printer):
        self.config = config
        self.logger = logger
        self.printer = printer

    def clean_name(self, text):
        if not text:
            return "Resident"
        try:
            response = client.chat.completions.create(model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"Extract the name including the person's title from the following text. Only provide the name, no additional text. If there is no obvious name, return 'Resident': '{text}'"}
            ],
            max_tokens=50)
            name = response.choices[0].message.content.strip()
            return name if name else "Resident"
        except Exception as e:
            self.logger.log('error', f"Error extracting name: {e}")
            return "Resident"

    def format_address(self, text):
        if not text:
            return "Address not available"
        try:
            response = client.chat.completions.create(model="gpt-4o",
            messages=[
                {"role": "system", "content": "You are a helpful assistant."},
                {"role": "user", "content": f"Format the following address for a letter with proper line breaks. Only provide the formatted address, no additional text: '{text}'"}
            ],
            max_tokens=150)
            formatted_address = response.choices[0].message.content.strip()
            return formatted_address if formatted_address else text
        except Exception as e:
            self.logger.log('error', f"Error formatting address: {e}")
            return text

    def sanitize_filename(self, filename):
        # Remove all non-alphanumeric characters except underscore
        valid_filename = "".join(c for c in filename if c.isalnum() or c == "_")
        valid_filename = valid_filename[:30]  # Further limit filename length to avoid path length issues
        return f"{valid_filename}.docx"  # Ensure the filename has the .docx extension

    def replace_placeholders(self, document, data):
        for placeholder, column in self.config['PLACEHOLDERS'].items():
            if column == self.config['NAME_COLUMN']:
                value = self.clean_name(data[column])
            elif column == self.config['ADDRESS_COLUMN']:
                value = self.format_address(data[column])
            else:
                value = str(data[column])

            self.logger.log('info', f'Replacing placeholder {placeholder} with {value}')

            # Replace placeholders in paragraphs
            for paragraph in document.paragraphs:
                if f'{{{{{placeholder}}}}}' in paragraph.text:
                    inline_replace(paragraph, f'{{{{{placeholder}}}}}', value)

            # Replace placeholders in tables
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if f'{{{{{placeholder}}}}}' in cell.text:
                            inline_replace(cell, f'{{{{{placeholder}}}}}', value)

        return document

    def generate_and_print_letters(self, data_list):
        for data in data_list:
            template_path = os.path.join(self.config['TEMPLATES_DIR'], f"{self.config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']}.docx")
            document = Document(template_path)
            personalized_document = self.replace_placeholders(document, data)
            sanitized_name = self.sanitize_filename(f"{data[self.config['NAME_COLUMN']]}")
            file_path = os.path.join(self.config['PRINT_SERVER_DIR'], sanitized_name)
            file_path = os.path.normpath(file_path)  # Normalize the path to make it Windows-friendly
            personalized_document.save(file_path)
            # self.printer.print_letter(file_path)
            self.logger.log('info', f'Generated letter for {data[self.config['NAME_COLUMN']]}')

def inline_replace(element, old_text, new_text):
    if old_text in element.text:
        element.text = element.text.replace(old_text, new_text)
    for run in element.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)

def load_defaults():
    if os.path.exists(DEFAULT_CONFIG_PATH):
        with open(DEFAULT_CONFIG_PATH, 'r') as f:
            return json.load(f)
    else:
        raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")

# Assuming the rest of your script uses this load_defaults function to load the configuration
if __name__ == "__main__":
    config = load_defaults()
    logger = Logger()
    # Assuming you have a Printer class already defined
    printer = Printer(config['PRINT_SERVER_DIR'], logger)
    letter_generator = LetterGenerator(config, logger, printer)
    # Assuming you have data_list available
    letter_generator.generate_and_print_letters(data_list)
