import pandas as pd
import re
import os

# Function to filter students with CGPA greater than 8 or percentage greater than 80
def process_data(file_path):
    xls = pd.ExcelFile(file_path)
    filtered_data = pd.DataFrame(columns=['Name', 'University', 'Graduation_Marks'])

    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name)
        sheet_data = sheet_data[(sheet_data['Graduation Marks'] > 10) & (sheet_data['Graduation Marks'] > 80) |
                                (sheet_data['Graduation Marks'] <= 10) & (sheet_data['Graduation Marks'] >= 8)]
        filtered_data = pd.concat([filtered_data, sheet_data])

    return filtered_data

def save_to_excel(dataframe, output_path):
    dataframe.to_excel(output_path, index=False)

# Define a regex pattern to match CFTI universities
keywords_pattern = r"(?i)(INDIAN\s*INSTITUTES?\s*OF\s*TECHNOLOGY|^IIT|INDIAN\s*INSTITUTES?\s*OF\s*MANAGEMENT|^IIM|INDIAN\s*INSTITUTE\s*OF\s*SCIENCE|^IISc|INDIAN\s*INSTITUTES?\s*OF\s*SCIENCE.*RESEARCH|IISER|Indian\s*Institute\s*of\s*Information\s*Technology.*Allahabad|Atal\s*Bihari\s*Vajpayee\s*Indian\s*Institute\s*of\s*Information\s*Technology.*Gwalior|^ABVIIITM|Pandit\s*Dwarka\s*Prasad\s*Mishra\s*Indian\s*Institute\s*of\s*Information\s*Technology.*Jabalpur|IIITDM|Indian\s*Institute\s*of\s*Information\s*Technology.*Kanchipuram|IITD&M|Indian\s*Institute\s*of\s*Information\s*Tehnology.*Kurnool.*Andhra\s*Pradesh|IIITDM|NATIONAL\s*INSTITUTES?\s*OF\s*TECHNICAL\s*TEACHERS?\s*TRAINING.*RESEARCH|NITTTR|NATIONAL\s*INSTITUTES?\s*OF\s*TECHNOLOGY|NIT|National\s*Institute\s*of\s*Industrial|NITIE|National\s*Institute\s*of\s*Foundry\s*Forge\s*Technology|NIFFT|School\s*of\s*Planning\s*and\s*Architecture|^SPA|Central\s*Institute\s*of\s*Technology\.Kokrajhar|Sant\s*Longowal|North\s*Eastern\s*Regional\s*Institute\s*of\s*Science.*Technology|Ghani\s*Khan\s*Choudhury\s*Institute\s*of\s*Engineering.*Technology)"

def is_cfti(university):
    return bool(re.search(keywords_pattern, university, re.IGNORECASE))

# Define relative file paths
input_file_path = 'DC PROJECT.xlsx'
output_excel_cfti_path = 'Result.xlsx'

# Get the absolute paths based on the script location
script_dir = os.path.dirname(os.path.abspath(__file__))
input_file_path = os.path.join(script_dir, input_file_path)
output_excel_cfti_path = os.path.join(script_dir, output_excel_cfti_path)

# Call the process_data function to generate the CFTI_Result.xlsx file
filtered_data = process_data(input_file_path)

# Filter rows where the university matches the criteria
filtered_data = filtered_data[filtered_data['Graduation University'].apply(is_cfti)]

# Save the filtered data to a new Excel file (CFTI_Result.xlsx)
filtered_data.to_excel(output_excel_cfti_path, index=False)

print("Eligible candidates saved to:", output_excel_cfti_path)
