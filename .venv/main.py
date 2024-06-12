from data_collection import DataCollector
from template_management import TemplateManager
from letter_generation import LetterGenerator
from printing import Printer
from watcher import TeamsExcelWatcher
from custom_logging import Logger
import config


def main():
    # Configuration
    templates_dir = config.TEMPLATES_DIR
    print_server_dir = config.PRINT_SERVER_DIR

    # Initialize logger
    logger = Logger()

    # Initialize components
    data_collector = DataCollector(logger)
    template_manager = TemplateManager(templates_dir, logger)
    printer = Printer(print_server_dir, logger)
    letter_generator = LetterGenerator(template_manager, logger, printer)

    if config.USE_TEAMS_EXCEL:
        watcher = TeamsExcelWatcher(
            data_collector,
            logger,
            config.TENANT_ID,
            config.CLIENT_ID,
            config.CLIENT_SECRET,
            config.EXCEL_FILE_ID,
            config.EXCEL_FILE_DRIVE,
            config.WATCHER_INTERVAL
        )
        excel_data = watcher.get_excel_data()
        df = data_collector.parse_excel_data(excel_data)
    else:
        df = data_collector.collect_data()

    # Define filters if needed
    filters = {
        config.LETTER_1_COLUMN: '',  # Example filter: only residents who haven't received the first letter
        config.LETTER_2_COLUMN: '',  # Example filter: only residents who haven't received the second letter
        config.LETTER_3_COLUMN: ''  # Example filter: only residents who haven't received the third letter
    }

    # Collect and filter data
    filtered_data = data_collector.filter_data(df, filters)
    data = filtered_data.to_dict(orient='records')

    # Generate and print letters
    letter_generator.generate_and_print_letters(data)

    # Start watcher (in a separate thread or process in a real implementation)
    # watcher.watch_for_changes()


if __name__ == '__main__':
    main()
