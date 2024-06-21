import logging

class Logger:
    def __init__(self, config, log_file='app.log', log_level=logging.INFO):
        self.config = config
        self.logger = logging.getLogger('LetterAutomationLogger')
        self.logger.setLevel(log_level)
        fh = logging.FileHandler(log_file)
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
        if level == 'debug':
            self.logger.debug(message)
        elif level == 'info':
            self.logger.info(message)
        elif level == 'warning':
            self.logger.warning(message)
        elif level == 'error':
            self.logger.error(message)
        elif level == 'critical':
            self.logger.critical(message)
