import os
import time
from random import randint
from collections import OrderedDict

import openpyxl
from pytrends.pyGTrends import pyGTrends


class GoogleDataGetter(object):

    def __init__(self, file_paths, credentials):
        # Set file paths
        self.file_paths = file_paths
        wb_filepath = os.path.abspath(file_paths['input_file'])
        # Read an input workbook
        self.in_workbook = openpyxl.load_workbook(filename=wb_filepath)
        # Read country codes.
        self.country_codes = self.get_country_codes(file_paths['country_codes_file'])
        # Connect to Google
        self.g_connector = pyGTrends(credentials['google_username'], credentials['google_password'])

    def process_all_countries(self):
        for country_code in self.in_workbook.get_sheet_names():
            print('====PROCESSING %s====' % country_code)
            self.process_country(country_code)

    def process_country(self, country_code):
        in_sheet = self.in_workbook[country_code]   # Read the country sheet.
        s_terms = self._get_search_terms_from_sheet(in_sheet)
        correct_c_code = self.country_codes[country_code]   # Replace Kapounek country code with correct 2-letter code.
        self.save_data_for_country(s_terms, correct_c_code, country_code)

    def save_data_for_country(self, terms_list, country_code, kapounek_code):
        """
        Download data from Google Trends and save it to CSV file.

        Args:
            terms_list (OrderedDict): Dict {ID, search term} of search terms to search on Google Trends.
            country_code (string): Two letter country abbreviation.
            kapounek_code (string): Kapounek country code.
        """
        # Prepare directory for files.
        csv_dir_path = self.file_paths['gdata_dir']+'/'+kapounek_code
        if not os.path.exists(csv_dir_path):
            os.makedirs(csv_dir_path)
        # Get all terms.
        for (term_id, term_text) in terms_list.items():
            # Prepare search term text.
            search_text = term_text.encode('utf8').strip()
            # Make request. URL example: https://www.google.com/trends/explore#q=koruna&geo=CZ&cmpt=q&tz=Etc%2FGMT-2
            self.g_connector.request_report(search_text, geo=country_code, tz="Etc/GMT-2")
            # Wait a random amount of seconds between requests to avoid bot detection.
            time.sleep(randint(5, 10))
            # Download file.
            self.g_connector.save_csv(csv_dir_path+'/', kapounek_code+'_'+str(term_id))
        # OK
        return True

    def _get_search_terms_from_sheet(self, in_sheet):
        terms = OrderedDict()
        for col in in_sheet.columns:
            # Check if an empty column was reached.
            if not col[0].value:
                break    # stop the loop
            terms[col[0].value] = col[1].value
        # result
        return terms

    @staticmethod
    def get_country_codes(codes_filepath):
        c_codes = OrderedDict()
        with open(codes_filepath) as code_file:
            code_file.readline()    # skip first line
            for line in code_file:
                l_items = line.split(';')
                c_codes[l_items[1].strip()] = l_items[2].strip()
            # result
            return c_codes
