import pandas as pd
import re
# Step 3: Filter and Merge Data
def process_data(file_path, cfti_colleges):  
    
    xls = pd.ExcelFile(file_path)

    # Create an empty DataFrame to store the filtered data
    filtered_data = pd.DataFrame(columns=['Name', 'University', 'Graduation_Marks', 'CFTI YES/NO'])

    # Loop through each sheet and filter students
    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name)
        
        # Check if "Graduation Marks" (CG) are greater than 10 and meet the additional conditions
        sheet_data = sheet_data[(sheet_data['Graduation Marks'] > 10) & (sheet_data['Graduation Marks'] > 80) |
                                (sheet_data['Graduation Marks'] <= 10) & (sheet_data['Graduation Marks'] >= 8)]

        filtered_data = pd.concat([filtered_data, sheet_data])

    # Update "CFTI YES/NO" Column
    filtered_data['CFTI YES/NO'] = filtered_data['University'].apply(lambda uni: 'YES' if uni in cfti_colleges else 'NO')

    return filtered_data

def save_to_excel(dataframe, output_path):
    dataframe.to_excel(output_path, index=False)

# Provided array of CFTI colleges
provided_cfti_colleges = [
    "INDIAN INSTITUTES OF TECHNOLOGY",
    "IIT"
    "INDIAN INSTITUTES OF MANAGEMENT ",
    "INDIAN INSTITUTE OF SCIENCE (IISc), BANGALORE",
    "INDIAN INSTITUTES OF SCIENCE EDUCATION AND RESEARCH",
    "IISER",
    "Indian Institute of Information Technology Allahabad",
    "Atal Bihari Vajpayee Indian Institute of Information Technology Gwalior",
    "ABVIIITM",
    "Pandit Dwarka Prasad Mishra Indian Institute of Information Technology, Design and Manufacturing Jabalpur",
    "IIITDM",
    "Indian Institute of Information Technology, Design and Manufacturing Kanchipuram",  
    "IITD&M",
    "Indian Institute of Information Tehnology, Design and Manufacturing Kurnool,Andhra Pradesh",
    "IIITDM", 
    "NATIONAL INSTITUTES OF TECHNICAL TEACHERS TRAINING AND RESEARCH",
    "NITTTR",
    "NATIONAL INSTITUTES OF TECHNOLOGY"  
    "NIT",
    "National Institute of Industrial Engg., Mumbai",
    "NITIE, Mumbai",
    "National Institute of Foundry & Forge Technology, Ranchi",
    "NIFFT Ranchi"
    "School of Planning & Architecture New Delhi",
    "School of Planning & Architecture  Bhopal",
    "School of Planning & Architecture  Vijayawada",
    "SPA Delhi",
    "SPA Bhopal",
    "SPA Vijaywada",
    "Central Institute of Technology, Kokrajhar"
    "Sant Longowal Institute of Engineering & Technology Longowal, Punjab",
    "SLIET",
    "North Eastern Regional Institute of Science & Technology (NERIST), Itanagar" 
    "Ghani Khan Choudhury Institute of Engineering & Technology(GKCIET), Malda, West Bengal"
    
]

# Replace file paths with the actual paths to your files
file_path = 'C:\\Users\\AYUSHI\\OneDrive\\Desktop\\DC\\DC PROJECT-1.xlsx'
output_excel_path = 'C:\\Users\\AYUSHI\\OneDrive\\Desktop\\DC\\Filtered_Result.xlsx'

# Process the data and obtain the result DataFrame
result = process_data(file_path, provided_cfti_colleges)

# Save the result to a new Excel file
save_to_excel(result, output_excel_path)

#print("Filtered data saved to:", output_excel_path)

output_excel_sheet = pd.read_excel(output_excel_path)

# Create an array of graduation universities from the output Excel data
graduation_universities_list = output_excel_sheet['Graduation University'].unique()

# Define a regex pattern to match universities with specific keywords
keywords_pattern = r"(?i)(INDIAN\s*INSTITUTES?\s*OF\s*TECHNOLOGY|^IIT|INDIAN\s*INSTITUTES?\s*OF\s*MANAGEMENT|^IIM|INDIAN\s*INSTITUTE\s*OF\s*SCIENCE.BANGALORE|^IISC|INDIAN\s*INSTITUTES?\s*OF\s*SCIENCE.*RESEARCH|IISER|Indian\s*Institute\s*of\s*Information\s*Technology\s*Allahabad|Atal\s*Bihari\s*Vajpayee\s*Indian\s*Institute\s*of\s*Information\s*Technology.*Gwalior|ABVIIITM|Pandit\s*Dwarka\s*Prasad\s*Mishra\s*Indian\s*Institute\s*of\s*Information\s*Technology.*Jabalpur|IIITDM|Indian\s*Institute\s*of\s*Information\s*Technology.*Kanchipuram|IITD&M|Indian\s*Institute\s*of\s*Information\s*Tehnology.*Kurnool.*Andhra\s*Pradesh|IIITDM|NATIONAL\s*INSTITUTES?\s*OF\s*TECHNICAL\s*TEACHERS?\s*TRAINING.*RESEARCH|NITTTR|NATIONAL\s*INSTITUTES?\s*OF\s*TECHNOLOGY|^NIT|National\s*Institute\s*of\s*Industrial\s*Mumbai|NITIE|National\s*Institute\s*of\s*Foundry\s*Forge\s*Technology.Ranchi|NIFFT|School\s*of\s*Planning\s&\s*Architecture.Delhi|School\s*of\s*Planning\s&\s*Architecture.Bhopal|School\s*of\s*Planning\s&\s*Architecture.*Vijayawada|^SPA|Central\s*Institute\s*of\s*Technology.*Kokrajhar|Sant\s*Longowal\s*|North\s*Eastern\s*Regional\s*Institute\s*of\s*Science.*Technology.*Itanagar|Ghani\s*Khan\s*Choudhury\s*Institute\s*of\s*Engineering.*Technology.*Malda.*West\s*Bengal)"

def is_cfti(university):
    return bool(re.search(keywords_pattern, university, re.IGNORECASE))

# Load the filtered data from the output Excel file
output_excel_path = 'C:\\Users\\AYUSHI\\OneDrive\\Desktop\\DC\\Filtered_Result.xlsx'
output_excel_sheet = pd.read_excel(output_excel_path)

# Filter rows in output_excel_sheet where the university matches the criteria
output_excel_sheet = output_excel_sheet[output_excel_sheet['Graduation University'].apply(is_cfti)]

# Save the filtered data to a new Excel file
filtered_output_excel_path = 'C:\\Users\\AYUSHI\\OneDrive\\Desktop\\DC\\Filtered_CFTI_Result.xlsx'
output_excel_sheet.to_excel(filtered_output_excel_path, index=False)

print("Filtered CFTI data saved to:", filtered_output_excel_path)
