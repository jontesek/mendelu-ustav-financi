from os import path
import csv

import openpyxl


class IndicatorsProcessor(object):

    def __init__(self, file_paths):
        # Set file paths
        self.file_paths = file_paths
        self.input_files = {
            'econ_freedom': path.abspath(file_paths['input_dir']+'/'+'economic-freedom-of-the-world-2015-dataset.xlsx'),
            'doing_business': path.abspath(file_paths['input_dir']+'/'+'Data_Extract_From_Doing_Business.xlsx'),
            'econ_heritage': path.abspath(file_paths['input_dir']+'/'+'economic_freedom_heritage-sorted.xlsx'),
            'cpi_db': path.abspath(file_paths['input_dir']+'/'+'cpi_okfn.csv'),
            'fin_open': path.abspath(file_paths['input_dir']+'/'+'kaopen_2013.xls'),
        }
        # Open output file
        out_path = path.abspath(file_paths['output_file'])
        self.out_workbook = openpyxl.load_workbook(filename=out_path)
        # Set source start positions (columns)
        self.source_pos = {
            'cpi': 'D', 'fin_open': 'E', 'global': 'F', 'polcon': 'J', 'shadow': 'L', 'ulc': 'M',
            'econ_freedom': 'N', 'econ_heritage': 'Q', 'do_biz': 'AB'
        }


    def write_years_and_countries(self, from_year, to_year):
        """
        Write years and countries to out worksheet data.
        """
        # Get country codes (and names) from out workbook.
        countries_ws = self.out_workbook.get_sheet_by_name('countries')
        countries_data = []
        for row in countries_ws.iter_rows('A2:B237'):
            name = row[0].value.strip()
            code = row[1].value.strip()
            countries_data.append([code, name])
        # Write data
        data_ws = self.out_workbook.get_sheet_by_name('data')
        write_row_n = 3
        for country in countries_data:
            for year in range(from_year, to_year+1):
                data_ws.cell(row=write_row_n, column=1).value = year
                data_ws.cell(row=write_row_n, column=2).value = country[0]
                data_ws.cell(row=write_row_n, column=3).value = country[1]
                write_row_n += 1
        # Save file
        self.out_workbook.save(self.file_paths['output_file'])

    def write_econ_freedom(self):
        """
        Economic freedom of the world:
        -> Business regulations, Freedom to trade internationaly, Top Marginal Tax Rate
        """
        # Open input
        in_wb = openpyxl.load_workbook(self.input_files['econ_freedom'])
        in_sheet = in_wb.get_sheet_by_name('Unadjusted Data')
        out_sheet = self.out_workbook.get_sheet_by_name('data')
        # Prepare input and output rows
        in_rows = in_sheet.rows[5:]
        out_rows = out_sheet.rows[2:]
        last_out_row_n = 0
        # Get numeric column index for the first output item.
        out_first_col_idx = openpyxl.utils.column_index_from_string(self.source_pos['econ_freedom'])
        # Read all input rows.
        for row_in in in_rows:
            # Check if the column is not empty - the end.
            if not row_in[1].value:
                break
            # Read values.
            year_in = row_in[1].value
            code_in = row_in[2].value.strip()
            print('IN %d: %s') % (year_in, code_in)
            # Get scores
            br_col_idx = openpyxl.utils.column_index_from_string('BQ')
            br_value = row_in[br_col_idx - 1].value
            freetr_idx = openpyxl.utils.column_index_from_string('AY')
            freetr_value = row_in[freetr_idx - 1].value
            topmar_idx = openpyxl.utils.column_index_from_string('P')
            topmar_value = row_in[topmar_idx - 1].value
            # Write scores to the correct output row.
            for i, row_out in enumerate(out_rows[last_out_row_n:], start=last_out_row_n):
                year_out = row_out[0].value
                code_out = row_out[1].value.strip()
                print('%d OUT %d: %s') % (i, year_out, code_out)
                if year_out == year_in and code_out == code_in:
                    row_out[out_first_col_idx - 1].value = br_value
                    row_out[out_first_col_idx].value = freetr_value
                    row_out[out_first_col_idx + 1].value = topmar_value
                    last_out_row_n = i
                    break
                else:
                    continue
        # Save file
        self.out_workbook.save('output/all_data_edited.xlsx')

    def write_econ_heritage(self):
        """
        Index of economic freedom (Heritage.org)
        Write all fields in order of appereance.
        """
        # Open input
        in_wb = openpyxl.load_workbook(self.input_files['econ_heritage'])
        in_sheet = in_wb.active
        out_sheet = self.out_workbook.get_sheet_by_name('data')
        # Prepare input and output rows
        in_rows = in_sheet.rows[1:]
        out_rows = out_sheet.rows[2:]
        last_out_row_n = 0
        # Get numeric column index for the first output item.
        out_first_col_idx = openpyxl.utils.column_index_from_string(self.source_pos['econ_heritage'])
        # Read all input rows.
        for row_in in in_rows:
            # Check if the column is not empty - the end.
            if not row_in[0].value:
                break
            # Read values.
            year_in = row_in[1].value
            cname_in = row_in[0].value.strip()
            #print('IN %d: %s') % (year_in, cname_in)
            # Get scores
            out_scores = [row_in[x].value if row_in[x].value != 'N/A' else '' for x in range(2, 13)]
            # Write scores to the correct output row.
            for row_n, row_out in enumerate(out_rows[last_out_row_n:], start=last_out_row_n):
                year_out = row_out[0].value
                cname_out = row_out[2].value.strip()
                #print('%d OUT %d: %s') % (row_n, year_out, cname_out)
                if year_out == year_in and cname_out == cname_in:
                    for i_s, score in enumerate(out_scores, out_first_col_idx):
                        row_out[i_s].value = score
                    last_out_row_n = row_n
                    break
                else:
                    continue
        # Save file
        self.out_workbook.save('output/all_data_edited.xlsx')


    def write_cpi(self):
        """
        Corruption Perceptions Index
        """
        # Prepared stuff
        in_file = open(self.input_files['cpi_db'], 'rb')
        cpi_reader = csv.reader(in_file)
        out_sheet = self.out_workbook.get_sheet_by_name('data')
        out_rows = out_sheet.rows[2:]
        last_out_row_n = 0
        out_first_col_idx = openpyxl.utils.column_index_from_string(self.source_pos['cpi']) - 1
        # Get the first (header) line.
        header_line = cpi_reader.next()
        first_year, last_year = int(header_line[1]), int(header_line[-1])

        # Read countries (lines) from file.
        for row in cpi_reader:
            cname_in = row[0]
            print('IN %s') % cname_in
            # Save yearly values (1998-2014).
            yearly_values = [float(row[x]) if row[x] != 'NA' else '' for x in range(1, len(header_line))]
            # Find the first position of the country in the out workbook.
            for row_n, row_out in enumerate(out_rows[last_out_row_n:], start=last_out_row_n):
                year_out = row_out[0].value
                cname_out = row_out[2].value.strip()
                #print('%d OUT %d: %s') % (row_n, year_out, cname_out)
                if year_out == first_year and cname_out == cname_in:
                    last_out_row_n = row_n
                    # Write values for all subsequent years.
                    for y_value in yearly_values:
                        row_out[out_first_col_idx].value = y_value
                        last_out_row_n += 1
                        row_out = out_rows[last_out_row_n]
                    break
                else:
                    continue
        # Save file
        self.out_workbook.save('output/all_data_edited.xlsx')


    def write_finopen(self):
        """
        http://web.pdx.edu/~ito/Chinn-Ito_website.htm
        format: excel
        years: 1970-2013
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