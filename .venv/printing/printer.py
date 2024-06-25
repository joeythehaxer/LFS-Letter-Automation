import os
import win32com.client as win32
from custom_logging.logger import Logger

class Printer:
    def __init__(self, print_server_dir, logger):
        self.print_server_dir = print_server_dir
        self.logger = logger

    def print_letter(self, file_path):
        if not os.path.exists(file_path):
            self.logger.log('error', f'File not found: {file_path}')
            raise FileNotFoundError(f'File not found: {file_path}')

        word = None
        doc = None
        try:
            self.logger.log('info', f'Opening document: {file_path}')
            word = win32.gencache.EnsureDispatch('Word.Application')
            word.Visible = False
            doc = word.Documents.Open(file_path)
            self.logger.log('info', f'Printing document: {file_path}')
            doc.PrintOut()
            self.logger.log('info', f'Successfully printed: {file_path}')
        except Exception as e:
            self.logger.log('error', f'Error printing document {file_path}: {e}')
            raise
        finally:
            if doc:
                doc.Close(False)
            if word:
                word.Quit()
