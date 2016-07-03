from src.GoogleDataJoiner import GoogleDataJoiner

####
# PARAMETER DEFINITIONS
####

# Define file paths.
file_paths = {
    'input_file': 'input/google_data_manual.xlsx',
    'country_codes_file': 'input/country_codes.csv',
    'output_dir': 'input',
    'gdata_dir': 'g_data',
}

# Create the main object.
gdg = GoogleDataJoiner(file_paths)

####
# EXECUTION PART
####
#gdg.process_country('SWE')
gdg.process_all_countries()
