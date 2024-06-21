import argparse
from gui.gui import run_gui
from config.settings import load_defaults
from custom_logging.logger import Logger
from data_collection.data_collector import DataCollector
from printing.printer import Printer
from template_management.template_manager import TemplateManager
from letter_generation.letter_generator import LetterGenerator
from watcher.teams_excel_watcher import TeamsExcelWatcher
import logging

def main():
    parser = argparse.ArgumentParser(description="Letter Automation System")
    parser.add_argument('--verbose', action='store_true', help="Enable verbose logging")
    args = parser.parse_args()

    config = load_defaults()
    logger = Logger(config, log_level=logging.DEBUG if args.verbose else logging.INFO)

    if config.USE_GUI:
        run_gui(config)
    else:
        data_collector = DataCollector(logger, config)
        printer = Printer(config.PRINT_SERVER_DIR, logger)
        template_manager = TemplateManager(config, logger)
        letter_generator = LetterGenerator(config, logger, printer, template_manager)

        if config.USE_TEAMS_EXCEL:
            watcher = TeamsExcelWatcher(data_collector, logger, config.TENANT_ID, config.CLIENT_ID, config.CLIENT_SECRET, config.EXCEL_FILE_ID, config.EXCEL_FILE_DRIVE)
            excel_data = watcher.get_excel_data()
            df = data_collector.parse_excel_data(excel_data)
        else:
            df = data_collector.collect_data()

        filtered_data = data_collector.filter_data(df)
        letter_generator.generate_and_print_letters(filtered_data.to_dict(orient='records'))

if __name__ == "__main__":
    main()
