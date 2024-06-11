import os
from logging import Logger

class Printer:
    def __init__(self, print_server_dir, logger):
        self.print_server_dir = print_server_dir
        self.logger = logger

    def print_letter(self, letter, filename):
        self.logger.log('info', f'Printing letter: {filename}')
        # Logic to print/send letter to print server
        file_path = os.path.join(self.print_server_dir, filename)
        with open(file_path, 'w') as file:
            file.write(letter)

    def send_to_print_server(self, letters, data):
        self.logger.log('info', 'Sending letters to print server')
        for letter, resident in zip(letters, data):
            filename = f"{resident['address']}_{resident['work_order_number']}.txt"
            self.print_letter(letter, filename)
