from src.IndicatorsProcessor import IndicatorsProcessor

file_paths = {
    'input_dir': 'input',
    'output_file': 'output/all_data.xlsx'
}

ip = IndicatorsProcessor(file_paths)

#ip.write_years_and_countries(1970, 2016)
#ip.write_econ_freedom()
#ip.write_econ_heritage()
#ip.write_cpi()
#ip.write_finopen()
#ip.write_global()
#ip.write_polcon()
#ip.write_shadow()
ip.write_dobiz()
