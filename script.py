import pandas as pd

# Replace these with the actual file paths
input_file_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\Input.xlsx'
consolidated_file_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\Consolidated.xlsx'
output_file_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\Output.xlsx'

# Read the Input and Consolidated Excel sheets
input_df = pd.read_excel(input_file_path)
consolidated_df = pd.read_excel(consolidated_file_path)

# print(consolidated_df.columns)

# Merge the two dataframes based on 'CID' and identify the non-matching rows
output_df = input_df.merge(consolidated_df[['CID', 'Institute', 'Name', 'GATE Reg No']],
                           on='CID', how='left', indicator=True)
output_df = output_df[output_df['_merge'] == 'left_only']

# Drop the '_merge' column
output_df = output_df.drop(['_merge'], axis=1)

# Save the non-matching rows to the "Output" Excel sheet
output_df.to_excel(output_file_path, index=False)
