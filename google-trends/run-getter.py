from src.GoogleDataGetter import GoogleDataGetter

####
# PARAMETER DEFINITIONS
####

# Define file paths.
file_paths = {
    'input_file': 'input/google_data_manual_edited.xlsx',
    'country_codes_file': 'input/country_codes.csv',
    'gdata_dir': 'g_data',
}

# Get google credentials.
with open('config.txt') as cfile:
    clines = cfile.read().split('\n')
    g_credentials = {
        'google_username': clines[0].split('=')[1],
        'google_password': clines[1].split('=')[1],
    }

# Create the main object.
gdg = GoogleDataGetter(file_paths, g_credentials)

####
# EXECUTION PART
####
#gdg.process_country('PL')
gdg.process_all_countries()
