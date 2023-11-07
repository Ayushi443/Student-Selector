import pandas as pd
import re
# Step 3: Filter and Merge Data
def process_data(file_path):  
    
    xls = pd.ExcelFile(file_path)

    # Create an empty DataFrame to store the filtered data
    filtered_data = pd.DataFrame(columns=['Name', 'University', 'Graduation_Marks'])

    # Loop through each sheet and filter students
    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name)
        
        # Check if "Graduation Marks" (CG) are greater than 10 and meet the additional conditions
        sheet_data = sheet_data[(sheet_data['Graduation Marks'] > 10) & (sheet_data['Graduation Marks'] > 80) |
                                (sheet_data['Graduation Marks'] <= 10) & (sheet_data['Graduation Marks'] >= 8)]

        filtered_data = pd.concat([filtered_data, sheet_data])



    return filtered_data

def save_to_excel(dataframe, output_path):
    dataframe.to_excel(output_path, index=False)

file_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\DC PROJECT-1.xlsx'
output_excel_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\CG_Result.xlsx'

# Call the process_data function to generate the CG_Result.xlsx file
filtered_data = process_data(file_path)
save_to_excel(filtered_data, output_excel_path)

output_excel_sheet = pd.read_excel(output_excel_path)

# Create an array of graduation universities from the output Excel data
graduation_universities_list = output_excel_sheet['Graduation University'].unique()

# Define a regex pattern to match universities with specific keywords
keywords_pattern = r"(?i)(INDIAN\s*INSTITUTES?\s*OF\s*TECHNOLOGY|^IIT|INDIAN\s*INSTITUTES?\s*OF\s*MANAGEMENT|^IIM|INDIAN\s*INSTITUTE\s*OF\s*SCIENCE|^IISc|INDIAN\s*INSTITUTES?\s*OF\s*SCIENCE.*RESEARCH|IISER|Indian\s*Institute\s*of\s*Information\s*Technology.*Allahabad|Atal\s*Bihari\s*Vajpayee\s*Indian\s*Institute\s*of\s*Information\s*Technology.*Gwalior|^ABVIIITM|Pandit\s*Dwarka\s*Prasad\s*Mishra\s*Indian\s*Institute\s*of\s*Information\s*Technology.*Jabalpur|IIITDM|Indian\s*Institute\s*of\s*Information\s*Technology.*Kanchipuram|IITD&M|Indian\s*Institute\s*of\s*Information\s*Tehnology.*Kurnool.*Andhra\s*Pradesh|IIITDM|NATIONAL\s*INSTITUTES?\s*OF\s*TECHNICAL\s*TEACHERS?\s*TRAINING.*RESEARCH|NITTTR|NATIONAL\s*INSTITUTES?\s*OF\s*TECHNOLOGY|NIT|National\s*Institute\s*of\s*Industrial|NITIE|National\s*Institute\s*of\s*Foundry\s*Forge\s*Technology|NIFFT|School\s*of\s*Planning\s*and\s*Architecture|^SPA|Central\s*Institute\s*of\s*Technology\.Kokrajhar|Sant\s*Longowal|North\s*Eastern\s*Regional\s*Institute\s*of\s*Science.*Technology|Ghani\s*Khan\s*Choudhury\s*Institute\s*of\s*Engineering.*Technology)"

def is_cfti(university):
    return bool(re.search(keywords_pattern, university, re.IGNORECASE))

# Load the filtered data from the output Excel file
output_excel_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\CG_Result.xlsx'
output_excel_sheet = pd.read_excel(output_excel_path)

# Filter rows in output_excel_sheet where the university matches the criteria
output_excel_sheet = output_excel_sheet[output_excel_sheet['Graduation University'].apply(is_cfti)]

# Save the filtered data to a new Excel file
filtered_output_excel_path = 'C:\\Users\\anugy\\Desktop\\VS code\\Student-Selector\\CFTI_Result.xlsx'
output_excel_sheet.to_excel(filtered_output_excel_path, index=False)

print("Filtered CFTI data saved to:", filtered_output_excel_path)