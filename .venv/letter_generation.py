import json
from docx import Document
import os
import re
from email_validator import validate_email, EmailNotValidError
from openai import OpenAI

DEFAULT_CONFIG_PATH = 'default_config.json'

# Load your OpenAI API key from an environment variable or a secure location


client = OpenAI(
    api_key=os.environ['sk-proj-0jmj1NH7ohBRKNRw7mP0T3BlbkFJY0ELu4NE0cEQbnqFbuLU'],
    # this is also the default, it can be omitted
)


class LetterGenerator:
    def __init__(self, template_manager, logger, printer):
        self.template_manager = template_manager
        self.logger = logger
        self.printer = printer
        self.load_defaults()

    def load_defaults(self):
        if os.path.exists(DEFAULT_CONFIG_PATH):
            with open(DEFAULT_CONFIG_PATH, 'r') as f:
                self.default_config = json.load(f)
        else:
            raise FileNotFoundError(
                f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")

    def extract_name(self, text):
        titles = ["Mr", "Mrs", "Ms", "Dr", "Miss", "Mister"]

        # Split the text by line breaks
        parts = text.split("\n")
        for part in parts:
            part = part.strip()

            # Remove numbers
            part = re.sub(r'\d+', '', part)

            # Remove emails
            try:
                validate_email(part)
                continue
            except EmailNotValidError:
                pass

            # Remove addresses (simplified, you can use more robust solutions)
            addresses = re.findall(r'\b\d{1,5}\s\w+\s\w+', part)
            if addresses:
                continue

                # Use OpenAI API to identify names and titles
                try:
                    client.completions.create(
                        engine="text-davinci-003",
                        prompt=f"Extract the name and title from the following text: '{part}'",
                        max_tokens=50
                    )
                    name = response.choices[0].text.strip()

                    return name
                except Exception as e:
                    print(f"Error during API processing: {e}")

            return None

    def clean_name(self, text):
        if not text:
            return "Resident"
        name = self.extract_name(text)
        if not name:
            return "Resident"
        return name

    def sanitize_filename(self, filename):
        return "".join(c for c in filename if c.isalnum() or c in (" ", ".", "_")).rstrip()

    def replace_placeholders(self, document, data):
        for placeholder, column in self.default_config['PLACEHOLDERS'].items():
            if column in data:
                value = str(data[column])  # Ensure the value is a string
                if placeholder == 'NAME_PLACEHOLDER':
                    value = self.clean_name(value)
                for paragraph in document.paragraphs:
                    if f'{{{{{placeholder}}}}}' in paragraph.text:
                        self.logger.log('info', f'Replacing {placeholder} with {value} in paragraph: {paragraph.text}')
                        paragraph.text = paragraph.text.replace(f'{{{{{placeholder}}}}}', value)
                for table in document.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if f'{{{{{placeholder}}}}}' in cell.text:
                                self.logger.log('info', f'Replacing {placeholder} with {value} in cell: {cell.text}')
                                cell.text = cell.text.replace(f'{{{{{placeholder}}}}}', value)
        return document

    def generate_and_print_letters(self, data_list):
        for data in data_list:
            template_name = self.template_manager.pick_template(data[self.default_config['REVIEW_COLUMN']])
            document = self.template_manager.load_template(template_name)
            personalized_document = self.replace_placeholders(document, data)
            sanitized_name = self.sanitize_filename(f"{data[self.default_config['NAME_COLUMN']]}.docx")
            file_path = os.path.join(self.default_config['PRINT_SERVER_DIR'], sanitized_name)
            file_path = os.path.normpath(file_path)  # Normalize the path to make it Windows-friendly
            personalized_document.save(file_path)
            self.printer.print_letter(file_path)
            self.logger.log('info', f'Generated and sent letter for {data[self.default_config['NAME_COLUMN']]}')
