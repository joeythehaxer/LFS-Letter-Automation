from data_collection import DataCollector
from template_management import TemplateManager
from letter_generation import LetterGenerator
from printing import Printer
from watcher import TeamsExcelWatcher
from custom_logging import Logger
import json
import os

DEFAULT_CONFIG_PATH = 'default_config.json'


def load_defaults():
    if os.path.exists(DEFAULT_CONFIG_PATH):
        with open(DEFAULT_CONFIG_PATH, 'r') as f:
            return json.load(f)
    else:
        raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")


def run_cli():
    # Load default configuration
    config = load_defaults()

    # Initialize logger
    logger = Logger()

    # Initialize components
    data_collector = DataCollector(logger)
    template_manager = TemplateManager(config['TEMPLATES_DIR'], logger)
    printer = Printer(config['PRINT_SERVER_DIR'], logger)
    letter_generator = LetterGenerator(template_manager, logger, printer)

    if config['USE_TEAMS_EXCEL']:
        watcher = TeamsExcelWatcher(
            data_collector,
            logger,
            config['TENANT_ID'],
            config['CLIENT_ID'],
            config['CLIENT_SECRET'],
            config['EXCEL_FILE_ID'],
            config['EXCEL_FILE_DRIVE'],
            config['WATCHER_INTERVAL']
        )
        excel_data = watcher.get_excel_data()
        df = data_collector.parse_excel_data(excel_data)
    else:
        df = data_collector.collect_data()

    # Collect and filter data
    filtered_data = data_collector.filter_data(df)
    data = filtered_data.to_dict(orient='records')

    # Generate and print letters
    letter_generator.generate_and_print_letters(data)


if __name__ == "__main__":
    config = load_defaults()
    if config['USE_GUI']:
        from gui import LetterAutomationGUI
        import tkinter as tk

        root = tk.Tk()
        app = LetterAutomationGUI(root)
        root.mainloop()
    else:
        run_cli()
