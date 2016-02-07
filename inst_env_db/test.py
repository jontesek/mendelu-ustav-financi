from src.IndicatorsProcessor import IndicatorsProcessor

file_paths = {
    'input_dir': 'input',
    'output_file': 'output/all_data.xlsx'
}

ip = IndicatorsProcessor(file_paths)

ip.write_years_and_countries(1990, 2016)
#ip.write_econ_freedom()
#ip.write_econ_heritage()
ip.write_cpi()
