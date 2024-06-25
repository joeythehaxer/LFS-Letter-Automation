import logging
from logging.handlers import RotatingFileHandler

class Logger:
    def __init__(self, config, log_file='app.log', log_level=logging.INFO):
        self.config = config
        self.logger = logging.getLogger('LetterAutomationLogger')
        self.logger.setLevel(log_level)
        if not self.logger.hasHandlers():  # Check if handlers are already set
            fh = RotatingFileHandler(log_file, maxBytes=1048576, backupCount=5)  # 1MB per file, max 5 files
            fh.setLevel(log_level)
            ch = logging.StreamHandler()
            ch.setLevel(log_level)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            fh.setFormatter(formatter)
            ch.setFormatter(formatter)
            self.logger.addHandler(fh)
            self.logger.addHandler(ch)

    def log(self, level, message):
        if not self.config.LOGGING_ENABLED:
            return
        getattr(self.logger, level)(message)
