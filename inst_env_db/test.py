from src.IndicatorsProcessor import IndicatorsProcessor

file_paths = {
    'input_dir': 'input',
    'output_file': 'output/all_data.xlsx'
}

ip = IndicatorsProcessor(file_paths)

#ip.write_years_and_countries()
ip.write_business_regulations()
