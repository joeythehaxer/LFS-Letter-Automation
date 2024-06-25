import os
import win32com.client as win32
from custom_logging.logger import Logger


class Printer:
    def __init__(self, print_server_dir, logger):
        self.print_server_dir = print_server_dir
        self.logger = logger

    def print_letter(self, file_name):
        full_path = os.path.join(self.print_server_dir, file_name)

        # Debugging logs to check paths
        self.logger.log('debug', f'Print server directory: {self.print_server_dir}')
        self.logger.log('debug', f'File name: {file_name}')
        self.logger.log('debug', f'Full file path for printing: {full_path}')

        if not os.path.exists(full_path):
            self.logger.log('error', f'File not found: {full_path}')
            raise FileNotFoundError(f'File not found: {full_path}')

        try:
            self.logger.log('info', f'Opening document: {full_path}')
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(full_path)
            self.logger.log('info', f'Printing document: {full_path}')
            doc.PrintOut()
            doc.Close(False)
            word.Quit()
            self.logger.log('info', f'Successfully printed: {full_path}')
        except Exception as e:
            self.logger.log('error', f'Error printing document {full_path}: {e}')
            raise
