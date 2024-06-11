from docx import Document
from custom_logging import Logger


class LetterGenerator:
    def __init__(self, template_manager, logger, printer):
        self.template_manager = template_manager
        self.logger = logger
        self.printer = printer

    def generate_letter(self, data, template_name):
        """
        Generates a letter by filling in the template with provided data.

        :param data: Dictionary containing the data to fill into the template.
        :param template_name: Name of the template to be used.
        :return: Generated letter as a Document object.
        """
        self.logger.log('info', f'Generating letter using template: {template_name}')
        # Load the template document
        template_path = self.template_manager.get_template_path(template_name)
        doc = Document(template_path)

        # Replace placeholders with actual data
        self.replace_placeholders(doc, data)

        return doc

    def replace_placeholders(self, doc, data):
        """
        Replaces placeholders in the document with actual data.

        :param doc: The Word document.
        :param data: Dictionary containing the data to fill into the template.
        """
        for paragraph in doc.paragraphs:
            self.replace_text_in_paragraph(paragraph, data)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self.replace_text_in_paragraph(paragraph, data)

    def replace_text_in_paragraph(self, paragraph, data):
        """
        Replaces placeholders in a paragraph with actual data.

        :param paragraph: The paragraph in the Word document.
        :param data: Dictionary containing the data to fill into the template.
        """
        for key, value in data.items():
            if f'{{{{ {key} }}}}' in paragraph.text:
                paragraph.text = paragraph.text.replace(f'{{{{ {key} }}}}', str(value))

    def generate_and_print_letters(self, data):
        """
        Generates and prints letters for all residents.

        :param data: List of dictionaries, each containing data for one resident.
        """
        self.logger.log('info', 'Generating and printing customized letters')
        for resident in data:
            # Determine which template to use based on the resident's letter status
            template_name = self.template_manager.pick_template(resident)
            if template_name:
                # Generate the letter using the determined template
                letter = self.generate_letter(resident, template_name)
                # Print the letter
                filename = f"{resident['address']}_{resident['work_order_number']}.docx"
                self.printer.print_letter(letter, filename)
