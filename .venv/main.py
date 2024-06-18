import json
import os
from data_collection import DataCollector
from letter_generation import LetterGenerator
from printing import Printer
from custom_logging import Logger

DEFAULT_CONFIG_PATH = 'default_config.json'


def load_defaults():
    if os.path.exists(DEFAULT_CONFIG_PATH):
        with open(DEFAULT_CONFIG_PATH, 'r') as f:
            return json.load(f)
    else:
        raise FileNotFoundError(f"{DEFAULT_CONFIG_PATH} not found. Please create it with the necessary configurations.")


def run_cli():
    config = load_defaults()
    logger = Logger()
    data_collector = DataCollector(logger, config)
    printer = Printer(config['PRINT_SERVER_DIR'], logger)
    letter_generator = LetterGenerator(config, logger, printer)

    if config['USE_TEAMS_EXCEL']:
        from watcher import TeamsExcelWatcher
        watcher = TeamsExcelWatcher(data_collector, logger, config)
        excel_data = watcher.get_excel_data()
        df = data_collector.parse_excel_data(excel_data)
    else:
        df = data_collector.collect_data()

    filtered_data = data_collector.filter_data(df)
    letter_generator.generate_and_print_letters(filtered_data)


if __name__ == "__main__":
    config = load_defaults()
    if config['USE_GUI']:
        from gui import run_gui

        run_gui(config)
    else:
        run_cli()
