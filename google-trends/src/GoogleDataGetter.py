from os import path
import time
from random import randint
from collections import OrderedDict
import urllib

import openpyxl
from pytrends.pyGTrends import pyGTrends


class GoogleDataGetter(object):

    def __init__(self, file_paths, credentials):
        # Set file paths
        self.file_paths = file_paths
        wb_filepath = path.abspath(file_paths['input_file'])
        # Read an input workbook
        self.in_workbook = openpyxl.load_workbook(filename=wb_filepath)
        # Create an output workbook
        self.out_workbook = openpyxl.Workbook()
        self.out_workbook.remove_sheet(self.out_workbook.active)    # remove defaultly created sheet
        # Connect to Google
        self.g_connector = pyGTrends(credentials['google_username'], credentials['google_password'])


    def process_all_countries(self):
        for country_code in self.in_workbook.get_sheet_names():
            print('====PROCESSING %s====' % country_code)
            self.process_country(country_code)
        # Save workbook to file
        out_filename = path.abspath(self.file_paths['output_dir']+'/google_data_auto.xlsx')
        self.out_workbook.save(out_filename)

    def process_country(self, country_code):
        # Read the country sheet.
        in_sheet = self.in_workbook[country_code]
        # Create a new empty sheet.
        out_sheet = self.out_workbook.create_sheet(country_code)
        # Get all words from sheet.
        s_terms = self.get_search_terms_from_sheet(in_sheet)
        c_data = self.get_data_for_country(s_terms, country_code)

    def get_data_for_country(self, terms_list, country_code):
        """
        Args:
            terms_list (OrderedDict): Dict {ID, search term} of search terms to search on Google Trends.
            country_code (string): Two letter country abbreviation

        Returns:
            data_list (list): list of {'term', 'values'} dicts.
        """
        data_list = []
        for (term_id, term_text) in terms_list.items():
            # Make request.
            # https://www.google.com/trends/explore#q=koruna&geo=CZ&cmpt=q&tz=Etc%2FGMT-2
            self.g_connector.request_report(term_text.encode('utf8'), geo=country_code, tz="Etc/GMT-2")
            # Wait a random amount of seconds between requests to avoid bot detection.
            time.sleep(randint(5, 10))
            # Download file.
            self.g_connector.save_csv(self.file_paths['gdata_dir']+'/', country_code+'_'+str(term_id))

    def write_data_to_sheet(self, data_list):
        pass

    def get_search_terms_from_sheet(self, in_sheet):
        terms = OrderedDict()
        for col in in_sheet.columns:
            # Check if an empty column was reached.
            if not col[0].value:
                break    # stop the loop
            terms[col[0].value] = col[1].value
        # result
        return terms
