import argparse
import logging
from gui.gui import run_gui
from config.settings import load_defaults
from custom_logging.logger import Logger
from data_collection.data_collector import DataCollector
from printing.printer import Printer
from template_management.template_manager import TemplateManager
from letter_generation.letter_generator import LetterGenerator
from watcher.teams_excel_watcher import TeamsExcelWatcher

def main():
    parser = argparse.ArgumentParser(description="Letter Automation System")
    parser.add_argument('--verbose', action='store_true', help="Enable verbose logging")
    args = parser.parse_args()

    # Load configurations
    config = load_defaults()

    # Initialize Logger
    logger = Logger(config, log_level=logging.DEBUG if args.verbose else logging.INFO)

    # Disable logging if configured
    if not config.LOGGING_ENABLED:
        logger.logger.disabled = True

    # Check if the application should run with GUI
    if config.USE_GUI:
        run_gui(config)
    else:
        # Initialize necessary components for processing
        data_collector = DataCollector(logger, config)
        printer = Printer(config.PRINT_SERVER_DIR, logger)
        template_manager = TemplateManager(config, logger)
        letter_generator = LetterGenerator(config, logger, printer, template_manager)

        # Handling for Microsoft Teams Excel integration or local Excel files
        if config.USE_TEAMS_EXCEL:
            watcher = TeamsExcelWatcher(data_collector, logger, config.TENANT_ID, config.CLIENT_ID,
                                        config.CLIENT_SECRET, config.EXCEL_FILE_ID, config.EXCEL_FILE_DRIVE)
            excel_data = watcher.get_excel_data()
            df = data_collector.parse_excel_data(excel_data)
        else:
            df = data_collector.collect_data()

        # Generate and print letters without filtering
        try:
            letter_generator.generate_and_print_letters(df.to_dict(orient='records'))
        except Exception as e:
            logger.log('error', f"Error during data processing: {e}")

if __name__ == "__main__":
    main()
