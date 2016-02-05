from os import path

import openpyxl


class IndicatorsProcessor(object):

    def __init__(self, file_paths):
        # Set file paths
        self.file_paths = file_paths
        self.input_files = {
            'business_regulations': path.abspath(file_paths['input_dir']+'/'+'economic-freedom-of-the-world-2015-dataset.xlsx'),
            'doing_business': path.abspath(file_paths['input_dir']+'/'+'Data_Extract_From_Doing_Business.xlsx'),
        }
        # Open output file
        out_path = path.abspath(file_paths['output_file'])
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
        br_sheet = br_wb.get_sheet_by_name('Unadjusted Data')
        out_sheet = self.out_workbook.get_sheet_by_name('data')
        # Prepare input and output rows
        in_rows = br_sheet.rows[5:]
        out_rows = out_sheet.rows[2:]
        last_out_row_n = 0
        # Read all input rows.
        for row_in in in_rows:
            # Check if the column is not empty - the end.
            if not row_in[1].value:
                break
            # Read values.
            year_in = row_in[1].value
            code_in = row_in[2].value.strip()
            print('IN %d: %s') % (year_in, code_in)
            # Get "Business regulations" score.
            br_col_idx = openpyxl.utils.column_index_from_string('BQ')
            br_value = row_in[br_col_idx - 1].value
            if not br_value:
                continue    # Skip empty values
            # Write score to the correct output row.
            for i, row_out in enumerate(out_rows[last_out_row_n:], start=last_out_row_n):
                year_out = row_out[0].value
                code_out = row_out[1].value.strip()
                print('%d OUT %d: %s') % (i, year_out, code_out)
                if year_out == year_in and code_out == code_in:
                    br_out_cell = row_out[2]
                    br_out_cell.value = br_value
                    last_out_row_n = i
                    break
                else:
                    continue
        # Save file
        self.out_workbook.save('output/all_data_edited.xlsx')

    def write_cpi(self):
        """
        http://data.okfn.org/data/core/corruption-perceptions-index
        format: CSV
        years: 1998-2014
        :return:
        """
    def write_finopen(self):
        """
        http://web.pdx.edu/~ito/Chinn-Ito_website.htm
        format: excel
        years: 1970-2013
        :return:
        """

    def write_econfree(self):
        """
        http://www.heritage.org/index/explore?view=by-region-country-year
        format: csv
        years: 1995-2016
        :return:
        """

    def write_oecd(self):
        """
        http://stats.oecd.org/Index.aspx?QueryName=426
        format: excel
        years: 1995-2012
        :return:
        """

    def write_global(self):
        """
        http://globalization.kof.ethz.ch/
        format: excel
        years: 1970-2012
        :return:
        """

    def write_polcon(self):
        """
        https://mgmt.wharton.upenn.edu/faculty/heniszpolcon/polcondataset/
        format: excel
        years: 1960-2012
        :return:
        """

    def write_shadow(self):
        """
        just PDFs not interested in this stuff
        :return:
        """
