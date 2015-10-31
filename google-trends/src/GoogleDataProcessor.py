from os import path
import re

import openpyxl


class GoogleDataProcessor(object):

    def __init__(self, file_paths):
        # Set file paths
        self.file_paths = file_paths
        wb_filepath = path.abspath(file_paths['input_file'])
        # Read an input workbook
        self.in_workbook = openpyxl.load_workbook(filename=wb_filepath)
        self.countries = self.in_workbook.get_sheet_names()
        # Create an output workbook
        self.out_workbook = openpyxl.Workbook()
        self.out_workbook.remove_sheet(self.out_workbook.active)    # remove defaultly created sheet
        # Regexp
        self.reg_week = re.compile('^(\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2}),(\d+)')
        self.reg_month = re.compile('^(\d{4}-\d{2}),(\d+)')

    def process_country(self, country_code):
        # Read the country sheet
        in_sheet = self.in_workbook[country_code]
        # Create a new empty sheet
        self.out_workbook.create_sheet(country_code)
        # Browse all columns (search terms)
        for col in in_sheet.columns:
            # Check if an empty column was reached.
            if not col[0].value:
                break    # stop the loop
            self._process_search_word(col)
            if col[0].value > 1:
                break

    def _process_search_word(self, column):
        word_id = column[0].value
        word_string = column[1].value
        # Check if the values have monthly or weekly frequency.
        if self.reg_week.match(column[2].value):
            print('week')
            word_data = self._process_weekly_values(column[2:])
        else:
            print('month')
            word_data = self._process_monthly_values(3, column[2:])
            return 0
        # Write the counted data to the output workbook.



    def _process_weekly_values(self, items):
        pass

    def _process_monthly_values(self, months_period, items):
        word_data = []
        period = {'year': None, 'sum': 0, 'number': 1}
        for index, item in enumerate(items):
            # Check if an empty row was reached.
            if not item.value:
                break    # stop the loop
            # Get values
            #print item.value
            result = self.reg_month.match(item.value.strip())
            result_count = long(result.group(2))
            result_year = result.group(1)[0:4]
            period['sum'] = result_count + period['sum']
            period['year'] = result_year
            # Check if the current period should end (there are 3 months in a quarter of a year).
            if long(index+1) % months_period == 0:
                print period
                # If yes, count the average score of the word and save it with the period ID.
                period_avg = period['sum'] / months_period
                period_name = '%s-%s' % (period['year'], period['number'])
                word_data.append((period_name, period_avg))
                # Reset the period
                period = {'year': None, 'sum': 0, 'number': period['number']+1}
                # If the new year is on the plan, reset period number.
                if period['number'] == 5:
                    period['number'] = 1
        # return data
        print word_data[0]
        return word_data
