import pandas as pd

# Step 3: Filter and Merge Data
def process_data(file_path, cfti_colleges):
    # Read the Excel file with three sheets
    xls = pd.ExcelFile(file_path)

    # Create an empty DataFrame to store the filtered data
    filtered_data = pd.DataFrame(columns=['Name', 'University', 'Graduation_Marks', 'CFTI YES/NO'])

    # Loop through each sheet and fikmlncoienlter stumnb,dents
    for sheet_name in xls.sheet_names:
        sheet_data = pd.read_excel(xls, sheet_name)
        filtered_data = pd.concat([filtered_data, sheet_data[sheet_data['Graduation Marks'] > 8]])

    # Update "CFTI YES/NO" Column
    filtered_data['CFTI YES/NO'] = filtered_data['University'].apply(lambda uni: 'YES' if uni in cfti_colleges else 'NO')

    return filtered_data

def save_to_excel(dataframe, output_path):
    dataframe.to_excel(output_path, index=False)

# Provided array of CFTI colleges
provided_cfti_colleges = [
    "INDIAN INSTITUTES OF TECHNOLOGY",
    "INDIAN INSTITUTES OF MANAGEMENT ",
    "INDIAN INSTITUTE OF SCIENCE (IISc), BANGALORE",
    "INDIAN INSTITUTES OF SCIENCE EDUCATION AND RESEARCH (IISERs) ",
    "Indian Institute of Information Technology Allahabad",
    "Atal Bihari Vajpayee - Indian Institute of Information Technology Gwalior- (ABVIIITM)",
    "Pandit Dwarka Prasad Mishra Indian Institute of Information Technology, Design and Manufacturing (IIITDM) Jabalpur",
    "Indian Institute of Information Technology, Design and Manufacturing (IITD&M) Kanchipuram",  
    "Indian Institute of Information Tehnology, Design and Manufacturing (IIITDM) Kurnool,Andhra Pradesh", 
    "NATIONAL INSTITUTES OF TECHNICAL TEACHERS TRAINING AND RESEARCH (NITTTRs)",
    "NATIONAL INSTITUTES OF TECHNOLOGY"  
    "National Institute of Industrial Engg. (NITIE), Mumbai",
    "National Institute of Foundry & Forge Technology (NIFFT), Ranchi",
    "School of Planning & Architecture (SPA), New Delhi",
    "School of Planning & Architecture (SPA), Bhopal"
    "School of Planning & Architecture (SPA), Vijayawada",
    "Central Institute of Technology, Kokrajhar"
    "Sant Longowal Institute of Engineering & Technology (SLIET), Longowal, Punjab"
    "North Eastern Regional Institute of Science & Technology (NERIST), Itanagar" 
    "Ghani Khan Choudhury Institute of Engineering & Technology(GKCIET), Malda, West Bengal"
    
]

# Replace file paths with the actual paths to your files
file_path = 'C:\\Users\\AYUSHI\\OneDrive\\Desktop\\DC\\DC PROJECT.xlsx'
output_excel_path = 'C:\\Users\\AYUSHI\\OneDrive\\Desktop\\DC\\Filtered_Result.xlsx'

# Process the data and obtain the result DataFrame
result = process_data(file_path, provided_cfti_colleges)

# Save the result to a new Excel file
save_to_excel(result, output_excel_path)

print("Filtered data saved to:", output_excel_path)
