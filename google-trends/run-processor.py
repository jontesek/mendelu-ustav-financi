from src.GoogleDataProcessor import GoogleDataProcessor

file_paths = {
    'input_file': 'input/google_data_auto.xlsx',
    'output_dir': 'output'
}
gdp = GoogleDataProcessor(file_paths, 'CZ', 2014)
#gdp.process_country('LIT')
#gdp.process_all_countries_to_one_sheet()
gdp.process_all_countries()
