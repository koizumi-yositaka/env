import pandas as pd
import json
import os

# Load Excel file
excel_path = 'mnt/data/テストデータ全体.xlsx'
data = pd.read_excel(excel_path)

# Load template JSON file
template_path = 'mnt/data/test1.json'
with open(template_path, 'r', encoding='utf-8') as file:
    template = json.load(file)

# Ensure output directory exists
output_dir = 'mnt/data/output_json'
os.makedirs(output_dir, exist_ok=True)

# Iterate over each row in the Excel file
for index, row in data.iterrows():
    # Create a copy of the template
    json_data = template.copy()
    
    # Set the values based on the Excel row
    for item in json_data['inputdata']:
        if item['id'] == '予約語名':
            item['value'] = row['引数名']
        elif item['id'] == '引数項':
            item['value'] = row['定義文']
    
    # Determine output file name
    file_name = f"test{row['ID']}.json"
    output_path = os.path.join(output_dir, file_name)
    
    # Write the JSON data to the output file
    with open(output_path, 'w', encoding='utf-8') as output_file:
        json.dump(json_data, output_file, ensure_ascii=False, indent=4)

print(f"JSON files have been generated in {output_dir}")
