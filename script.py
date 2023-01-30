import os
import pandas as pd

def save_to_excel(excel_file_path):
    data = []
    line_number = 1
    script_folder = os.path.dirname(os.path.realpath(__file__))
    try:
        for root, dirs, files in os.walk(script_folder):
            for file in files:
                file_path = os.path.join(root, file)
                file_name, file_extension = os.path.splitext(file)
                data.append({
                    'line_number': line_number,
                    'folder': root,
                    'file_name': file_name,
                    'file_extension': file_extension,
                })
                line_number += 1

        df = pd.DataFrame(data)
        df.to_excel(excel_file_path, index=False)
    except Exception as e:
        print(f'An error occurred: {e}')

excel_file_path = 'result.xlsx'
save_to_excel(excel_file_path)
