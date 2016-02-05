from os import path

import openpyxl


class IndicatorsProcessor(object):

    def __init__(self, file_paths):
        # Set file paths
        self.file_paths = file_paths
        out_path = path.abspath(file_paths['output_file'])
        # Set input filenames
        self.input_files = {
            'business_regulations': path.abspath(file_paths['input']+'/'+'economic-freedom-of-the-world-2015-dataset.xlsx'),
            'doing_business': path.abspath(file_paths['input']+'/'+'Data_Extract_From_Doing_Business.xlsx'),
        }
        # Open output file
        self.out_workbook = openpyxl.load_workbook(filename=out_path)

    def write_years_and_countries(self):
        # Get country codes from out workbook.
        countries_ws = self.out_workbook.get_sheet_by_name('countries')
        country_codes = []
        for row in countries_ws.iter_rows('B2:B237'):
            code = row[0].value
            country_codes.append(code)
        # Write data
        data_ws = self.out_workbook.get_sheet_by_name('data')
        write_row_n = 3
        for country in country_codes:
            for year in range(1990, 2016):
                data_ws.cell(row=write_row_n, column=1).value = year
                data_ws.cell(row=write_row_n, column=2).value = country
                write_row_n += 1
        # Save file
        self.out_workbook.save(self.file_paths['output_file'])

    def write_business_regulations(self):
        # Open input
        br_wb = openpyxl.load_workbook(self.input_files['business_regulations'])
        # Read input and write to output
        
