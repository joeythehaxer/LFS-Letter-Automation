import os
from datetime import datetime
from docx import Document
from openai import OpenAI
from template_management import TemplateManager
import pandas as pd

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

        # Save the updated DataFrame to Excel
        self.save_updated_data(data)

    def save_updated_data(self, data):
        try:
            df = pd.read_excel(self.config['LOCAL_EXCEL_FILE'], sheet_name=self.config['EXCEL_SHEET_NAME'],
                               header=self.config['HEADER_ROW'] - 1)
            self.logger.log('info', 'Loaded Excel file for updating')

            # Ensure the headers are correct
            expected_headers = set(
                [self.config['NAME_COLUMN'], self.config['LETTER_1_COLUMN'], self.config['LETTER_2_COLUMN'],
                 self.config['LETTER_3_COLUMN']])
            actual_headers = set(df.columns)
            self.logger.log('info', f'Expected headers: {expected_headers}')
            self.logger.log('info', f'Actual headers: {actual_headers}')
            if not expected_headers.issubset(actual_headers):
                self.logger.log('error', 'Excel file headers do not match the expected headers')
                raise ValueError('Excel file headers do not match the expected headers')

            # Ensure the data types are consistent for matching
            data_name = str(data[self.config['NAME_COLUMN']])
            df[self.config['NAME_COLUMN']] = df[self.config['NAME_COLUMN']].astype(str)
            self.logger.log('info', f'Converted {self.config["NAME_COLUMN"]} column to string for matching')

            # Find the row that matches the current data
            matching_row = df[self.config['NAME_COLUMN']] == data_name
            self.logger.log('info', f'Matching row: {matching_row}')

            # Check if any row matches
            if matching_row.any():
                self.logger.log('info', f'Match found, updating the row for {data_name}')
                # Create a Series with the same columns as df
                updated_series = pd.Series(data, index=df.columns)
                self.logger.log('info', f'Updated series to be set: {updated_series}')
                df.loc[matching_row, :] = updated_series
                self.logger.log('info', 'Updated DataFrame with new data')
            else:
                self.logger.log('warning', f'No matching row found for {data_name}')

            df.to_excel(self.config['LOCAL_EXCEL_FILE'], index=False)
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
    return config.load_defaults()


if __name__ == "__main__":
    config = load_defaults()
    logger = Logger()
    printer = Printer(config['PRINT_SERVER_DIR'], logger)
    template_manager = TemplateManager(config['TEMPLATES_DIR'], logger)
    letter_generator = LetterGenerator(config, logger, printer, template_manager)
    data_list = []  # Assume this is populated from somewhere
    letter_generator.generate_and_print_letters(data_list)
