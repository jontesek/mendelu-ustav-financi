from src.GoogleDataProcessor import GoogleDataProcessor

file_paths = {
    'input_file': 'input/google_data.xlsx',
    'output_dir': 'output'
}
gdp = GoogleDataProcessor(file_paths)
gdp.process_country('CZ')
