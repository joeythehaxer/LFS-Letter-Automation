import os
from datetime import datetime
from docx import Document
from openai import OpenAI
from template_management import TemplateManager
import pandas as pd
from openpyxl import load_workbook

client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
from custom_logging import Logger


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
                     "content": f"Format the following address for a letter with proper line breaks. Only provide the formatted address, no additional text: '{text}'"}
                ],
                max_tokens=150
            )
            formatted_address = response.choices[0].message.content.strip()
            return formatted_address if formatted_address else text
        except Exception as e:
            self.logger.log('error', f"Error formatting address: {e}")
            return text

    def sanitize_filename(self, filename):
        valid_filename = "".join(c for c in filename if c.isalnum() or c == "_")
        valid_filename = valid_filename[:30]  # Further limit filename length to avoid path length issues
        return f"{valid_filename}.docx"  # Ensure the filename has the .docx extension

    def replace_placeholders(self, document, data):
        for placeholder, column in self.config['PLACEHOLDERS'].items():
            if column == self.config['NAME_COLUMN']:
                value = self.clean_name(data[column])
            elif column == self.config['ADDRESS_COLUMN']:
                value = self.format_address(data[column])
            elif column == 'Date':
                value = datetime.now().strftime("%d %B %Y")
            elif column == self.config['WORK_ORDER_COLUMN']:
                value = data[column]
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
                        inline_replace(cell, f'{{{{{placeholder}}}}}', value)

            # Replace placeholders in shapes (text boxes)
            for shape in document.inline_shapes:
                if shape._inline.graphic.graphicData.uri.endswith('/textFrame'):
                    for paragraph in shape.text_frame.paragraphs:
                        inline_replace(paragraph, f'{{{{{placeholder}}}}}', value)

        return document

    def generate_and_print_letters(self, data_list):
        for data in data_list:
            try:
                self.logger.log('info', f"Processing data for: {data}")
                template_name = self.template_manager.determine_next_letter(data)
                if template_name:
                    self.logger.log('info', f'Using template: {template_name}')
                    document = self.template_manager.load_template(template_name)
                    self.logger.log('info', f'Template loaded: {template_name}')
                    personalized_document = self.replace_placeholders(document, data)
                    sanitized_name = self.sanitize_filename(f"{data[self.config['NAME_COLUMN']]}")
                    file_path = os.path.join(self.config['PRINT_SERVER_DIR'], sanitized_name)
                    file_path = os.path.normpath(file_path)  # Normalize the path to make it Windows-friendly
                    self.logger.log('info', f'Saving document to: {file_path}')
                    personalized_document.save(file_path)
                    if os.path.exists(file_path):
                        self.logger.log('info', f'Document saved successfully: {file_path}')
                        try:
                            pass  # self.printer.print_letter(file_path)
                            # self.logger.log('info', f'Printed letter for {data[self.config["NAME_COLUMN"]]}')
                            # Update the respective letter column with "sent letter" + current date
                            self.update_letter_column(data, template_name)
                        except Exception as e:
                            self.logger.log('error', f'Error printing document {file_path}: {e}')
                    else:
                        self.logger.log('error', f'Failed to save document: {file_path}')
                else:
                    self.logger.log('info', f'Skipping {data[self.config["NAME_COLUMN"]]}, all letters have been sent.')
            except Exception as e:
                self.logger.log('error',
                                f"Error generating and printing letters for {data.get(self.config['NAME_COLUMN'], 'Unknown')}: {e}")

    def update_letter_column(self, data, template_name):
        self.logger.log('info', f'Updating letter column for template: {template_name}')
        if template_name == self.config['TEMPLATE_GROUP1']['LETTER_1_TEMPLATE']:
            column = self.config['LETTER_1_COLUMN']
        elif template_name == self.config['TEMPLATE_GROUP1']['LETTER_2_TEMPLATE']:
            column = self.config['LETTER_2_COLUMN']
        elif template_name == self.config['TEMPLATE_GROUP1']['LETTER_3_TEMPLATE']:
            column = self.config['LETTER_3_COLUMN']
        else:
            self.logger.log('warning', f'No matching column found for template: {template_name}')
            return

        current_date = datetime.now().strftime("%d %B %Y")
        data[column] = f"sent letter {current_date}"
        self.logger.log('info',
                        f'Updated {column} with "sent letter {current_date}" for {data[self.config["NAME_COLUMN"]]}')

        # Save the updated cell to Excel
        self.save_updated_data(data, column)

    def save_updated_data(self, data, column):
        try:
            # Load the workbook and select the sheet
            workbook = load_workbook(self.config['LOCAL_EXCEL_FILE'])
            sheet = workbook[self.config['EXCEL_SHEET_NAME']]
            self.logger.log('info', 'Loaded Excel file for updating')

            # Find the address column index
            headers = [cell.value for cell in sheet[self.config['HEADER_ROW']]]
            self.logger.log('info', f'Headers in sheet: {headers}')
            if self.config['ADDRESS_COLUMN'] not in headers:
                self.logger.log('error', 'Address column not found in headers')
                raise ValueError('Address column not found in headers')

            address_idx = headers.index(self.config['ADDRESS_COLUMN']) + 1  # Convert to 1-based index

            # Find the row that matches the current data
            for row in sheet.iter_rows(min_row=self.config['HEADER_ROW'] + 1, max_row=sheet.max_row):
                if str(row[address_idx - 1].value) == str(data[self.config['ADDRESS_COLUMN']]):
                    row_idx = row[0].row
                    self.logger.log('info', f'Matching row found: {row_idx}')
                    # Update the specific cell
                    col_idx = headers.index(column) + 1  # Convert to 1-based index
                    sheet.cell(row=row_idx, column=col_idx, value=data[column])
                    break
            else:
                self.logger.log('warning', f'No matching row found for address: {data[self.config["ADDRESS_COLUMN"]]}')

            # Save the workbook
            workbook.save(self.config['LOCAL_EXCEL_FILE'])
            self.logger.log('info', 'Excel file updated successfully')
        except Exception as e:
            self.logger.log('error', f'Error updating Excel file: {e}')
            raise


def inline_replace(element, old_text, new_text):
    if old_text in element.text:
        element.text = element.text.replace(old_text, new_text)
    for run in element.runs:
        if old_text in run.text:
            run.text = run.text.replace(old_text, new_text)


def load_defaults():
    with open('default_config.json', 'r') as f:
        return json.load(f)


if __name__ == "__main__":
    config = load_defaults()
    logger = Logger()
    printer = Printer(config['PRINT_SERVER_DIR'], logger)
    template_manager = TemplateManager(config['TEMPLATES_DIR'], logger)
    letter_generator = LetterGenerator(config, logger, printer, template_manager)
    data_list = []  # Assume this is populated from somewhere
    letter_generator.generate_and_print_letters(data_list)
