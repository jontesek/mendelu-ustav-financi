from src.IndicatorsProcessor import IndicatorsProcessor

file_paths = {
    'input_dir': 'input',
    'output_dir': 'output',
}

ip = IndicatorsProcessor(file_paths)

# ip.write_years_and_countries(1970, 2016)
# exit()

ip.write_cpi()
ip.write_finopen()
ip.write_global()
ip.write_polcon()
ip.write_shadow()
ip.write_ulc()
ip.write_econ_freedom()
ip.write_econ_heritage()
ip.write_dobiz()


ip.write_workbook_to_file()
