import pandas as pd
import os

# Define relative file paths
input_file_path = 'Sample_Input.xlsx'
consolidated_file_path = 'Consolidated.xlsx'
output_file_path = 'Output.xlsx'

# Get the absolute paths based on the script location
script_dir = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.join(script_dir, input_file_path)
consolidated_file_path = os.path.join(script_dir, consolidated_file_path)
output_file_path = os.path.join(script_dir, output_file_path)

# Read the Input and Consolidated Excel sheets
input_df = pd.read_excel(input_file_path)
consolidated_df = pd.read_excel(consolidated_file_path)

# Merge the two dataframes based on 'CID' and identify the non-matching rows
output_df = input_df.merge(consolidated_df[['CID', 'Institute', 'Name', 'GATE Reg No']],
                            on='CID', how='left', indicator=True)
output_df = output_df[output_df['_merge'] == 'left_only']

# Drop the '_merge' column
output_df = output_df.drop(['_merge'], axis=1)

# Save the non-matching rows to the "Output" Excel sheet
output_df.to_excel(output_file_path, index=False)

print("Candidates for seat allocation saved to: ", output_file_path)
