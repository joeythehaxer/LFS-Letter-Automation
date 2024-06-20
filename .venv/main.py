import config
import webbrowser
from data_collection import DataCollector
from letter_generation import LetterGenerator
from printing import Printer
from custom_logging import Logger
from template_management import TemplateManager
import auth_helper

def run_cli():
    logger = Logger()
    data_collector = DataCollector(logger, config.config)
    printer = Printer(config.config['PRINT_SERVER_DIR'], logger)
    template_manager = TemplateManager(config.config['TEMPLATES_DIR'], logger)
    letter_generator = LetterGenerator(config.config, logger, printer, template_manager)

    if config.config['USE_TEAMS_EXCEL']:
        from watcher import TeamsExcelWatcher
        watcher = TeamsExcelWatcher(data_collector, logger, config.config)
        excel_data = watcher.get_excel_data()
        if not excel_data:
            print("Failed to fetch Excel data. Starting authentication flow...")
            webbrowser.open("http://127.0.0.1:5000/login")
            auth_helper.app.run(debug=True, use_reloader=False)  # Start Flask app for authentication
            return
        df = data_collector.parse_excel_data(excel_data)
    else:
        df = data_collector.collect_data()

    filtered_data = data_collector.filter_data(df)
    letter_generator.generate_and_print_letters(filtered_data.to_dict(orient='records'))

if __name__ == "__main__":
    if config.config['USE_GUI']:
        from gui import run_gui
        run_gui(config.config)
    else:
        run_cli()
