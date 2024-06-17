import os
from custom_logging import Logger

class Printer:
    def __init__(self, print_server_dir, logger):
        self.print_server_dir = print_server_dir
        self.logger = logger

    def print_letter(self, file_path):
        try:
            self.logger.log('info', f'Printing document: {file_path}')
            # Placeholder for actual print logic
            if os.path.exists(file_path):
                # Example command to print a document on Windows (this command may vary)
                os.system(f'start /min notepad /p {file_path}')
                self.logger.log('info', f'Document sent to printer: {file_path}')
            else:
                self.logger.log('error', f'File not found: {file_path}')
        except Exception as e:
            self.logger.log('error', f'Error printing document {file_path}: {e}')
