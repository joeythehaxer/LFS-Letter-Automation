import os
from custom_logging import Logger

class Printer:
    def __init__(self, print_server_dir, logger):
        self.print_server_dir = print_server_dir
        self.logger = logger

    def print_letter(self, letter, filename):
        self.logger.log('info', f'Printing letter: {filename}')
        # Save the document to the print server directory
        file_path = os.path.join(self.print_server_dir, filename)
        letter.save(file_path)

    def send_to_print_server(self, letters, data):
        self.logger.log('info', 'Sending letters to print server')
        for (letter, resident) in letters:
            filename = f"{resident['address']}_{resident['work_order_number']}.docx"
            self.print_letter(letter, filename)
