import config
from custom_logging import Logger
from docx import Document

class LetterGenerator:
    def __init__(self, template_manager, logger, printer):
        self.template_manager = template_manager
        self.logger = logger
        self.printer = printer

    def replace_placeholders(self, document, data):
        for placeholder, column in config.PLACEHOLDERS.items():
            if column in data:
                for paragraph in document.paragraphs:
                    if f'{{{{{placeholder}}}}}' in paragraph.text:
                        paragraph.text = paragraph.text.replace(f'{{{{{placeholder}}}}}', data[column])
                for table in document.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            if f'{{{{{placeholder}}}}}' in cell.text:
                                cell.text = cell.text.replace(f'{{{{{placeholder}}}}}', data[column])
        return document

    def generate_and_print_letters(self, data_list):
        for data in data_list:
            template_name = self.template_manager.pick_template(data[config.REVIEW_COLUMN])
            document = self.template_manager.load_template(template_name)
            personalized_document = self.replace_placeholders(document, data)
            file_path = os.path.join(config.PRINT_SERVER_DIR, f"{data[config.NAME_COLUMN]}.docx")
            personalized_document.save(file_path)
            self.printer.print_letter(file_path)
            self.logger.log('info', f'Generated and sent letter for {data[config.NAME_COLUMN]}')
