import os
from datetime import datetime
from docx import Document
from openai import OpenAI
from template_management.template_manager import TemplateManager
from custom_logging.logger import Logger

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))


class LetterGenerator:
    def __init__(self, config, logger, printer, template_manager):
        self.config = config
        self.logger = logger
        self.printer = printer
        self.template_manager = template_manager

    def clean_name(self, text):
        if not text:
            return "Resident"
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user",
                     "content": f"Extract the name including the person's title from the following text. Only provide the name, no additional text. If there is no obvious name, return 'Resident': '{text}'"}
                ],
                max_tokens=50
            )
            name = response.choices[0].message.content.strip()
            return name if name else "Resident"
        except Exception as e:
            self.logger.log('error', f"Error extracting name: {e}")
            return "Resident"

    def format_address(self, text):
        if not text:
            return "Address not available"
        try:
            response = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user",
                     "content": f"Format the following address for a letter with proper line breaks, fill in missing parts if you know the address. Only provide the formatted address, no additional text: '{text}'"}
                ],
                max_tokens=150
            )
            formatted_address = response.choices[0].message.content.strip()
            return formatted_address if formatted_address else text
        except Exception as e:
            self.logger.log('error', f"Error formatting address: {e}")
            return text

    def sanitize_filename(self, wo, address):
        """Generate a valid filename using the work order number and a shortened address."""
        short_address = address[:20].replace('/', '').replace('\\', '').replace(':', '').replace('*', '').replace('?',
                                                                                                                  '').replace(
            '"', '').replace('<', '').replace('>', '').replace('|', '')
        filename = f"{wo}_{short_address}".replace(' ', '_')
        valid_filename = "".join(c for c in filename if c.isalnum() or c in "_-")
        valid_filename = valid_filename[:255]  # Limit the filename length if necessary
        return f"{valid_filename}.docx"

    def replace_placeholders(self, document, data):
        for placeholder, column in self.config.PLACEHOLDERS.items():
            value = self.get_value_for_placeholder(column, data)

            self.logger.log('info', f'Replacing placeholder {placeholder} with {value}')
            self.replace_in_document(document, placeholder, value)

        return document

    def get_value_for_placeholder(self, column, data):
        if column == self.config.NAME_COLUMN:
            return self.clean_name(data[column])
        elif column == self.config.ADDRESS_COLUMN:
            return self.format_address(data[column])
        elif column == 'Date':
            return datetime.now().strftime("%d %B %Y")
        elif column == self.config.WORK_ORDER_COLUMN:
            return data[column]
        else:
            return str(data[column])

    def replace_in_document(self, document, placeholder, value):
        tag = f'{{{{{placeholder}}}}}'
        for paragraph in document.paragraphs:
            paragraph.text = paragraph.text.replace(tag, value)
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell.text = cell.text.replace(tag, value)
        for shape in document.inline_shapes:
            if hasattr(shape, 'text_frame'):
                for paragraph in shape.text_frame.paragraphs:
                    paragraph.text = paragraph.text.replace(tag, value)

    def generate_and_print_letters(self, data_list):
        for data in data_list:
            try:
                wo = data[self.config.PLACEHOLDERS['WO']]
                address = data[self.config.PLACEHOLDERS['ADDRESS_PLACEHOLDER']]
                sanitized_name = self.sanitize_filename(wo, address)
                file_path = os.path.join(self.config.PRINT_SERVER_DIR, sanitized_name)

                # Debugging logs to check the constructed paths
                self.logger.log('debug', f'Work Order: {wo}')
                self.logger.log('debug', f'Address: {address}')
                self.logger.log('debug', f'Sanitized Name: {sanitized_name}')
                self.logger.log('debug', f'Print Server Directory: {self.config.PRINT_SERVER_DIR}')
                self.logger.log('debug', f'Constructed File Path: {file_path}')

                self.logger.log('info', f'Saving document to: {file_path}')
                self.logger.log('info', f"Processing data for: {data}")
                template_name = self.template_manager.determine_next_letter(data)
                if template_name:
                    self.logger.log('info', f'Using template: {template_name}')
                    document = self.template_manager.load_template(template_name)
                    self.logger.log('info', f'Template loaded: {template_name}')
                    personalized_document = self.replace_placeholders(document, data)
                    personalized_document.save(file_path)
                    if os.path.exists(file_path):
                        self.logger.log('info', f'Document saved successfully: {file_path}')
                        try:
                            pass
                            # self.printer.print_letter(sanitized_name)  # Pass only the file name to the printer
                            # self.logger.log('info', f'Printed letter for {data[self.config.NAME_COLUMN]}')
                        except Exception as e:
                            self.logger.log('error', f'Error printing document {file_path}: {e}')
                    else:
                        self.logger.log('error', f'Failed to save document: {file_path}')
                    self.template_manager.update_excel(data, template_name)
                else:
                    self.logger.log('info', f'Skipping {data[self.config.NAME_COLUMN]}, all letters have been sent.')
            except Exception as e:
                self.logger.log('error',
                                f"Error generating and printing letters for {data.get(self.config.NAME_COLUMN, 'Unknown')}: {e}")
