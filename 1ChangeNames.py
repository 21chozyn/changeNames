import os
import openpyxl

# Load the candidate information from the Excel file
wb = openpyxl.load_workbook('candidate_info.xlsx')
ws = wb.active
count = 0
ws3 = wb["Gen"]
# Loop through each .docx file in the folder
for filename in os.listdir('.'):
    if filename.endswith('.doc') or filename.endswith(".docx"):
        # Get the last name from the filename
        last_name = filename.split('_')[-1].split()[0]
        try:
            first_name = filename.split('_')[-1].split()[1]
        except:
            if len(filename.split('_')[-1].split() == 1) and len(filename.split() == 1): #this means file name is of format name.docx
                continue
            elif len(filename.split("_") == 2) and len(filename.split() == 1): #this means file name is of format Lname_Firstname
                first_name = filename.split('_')[-1].split()[1]

        if filename.endswith('.doc') :
            extension = ".doc"
        else:
            extension = ".docx"

        # Find the row in the Excel file that matches the last name
        for rowNum in range(1,678):
            if ((ws3[f'A{rowNum}']).value == None) or (ws3[f'A{rowNum}']).value.startswith("2") == False:
                continue
            Lname = (ws3[f'B{rowNum}']).value.split()[0]
            Fname = (ws3[f'B{rowNum}']).value.split()[1]
            print(Lname.lower(),Fname.lower())
            if Lname.lower() == last_name.lower() and Fname.lower() in first_name.lower() :
                # Get the candidate information from the Excel file
                candidate_number = (ws3[f'A{rowNum}']).value
                formated_candidate_number = candidate_number.replace("/", "-" )
                first_name = (ws3[f'B{rowNum}']).value.split()[1]
                last_name = Lname.upper()
                # Create the new filename with the candidate information
                new_filename = f'{formated_candidate_number}_{last_name}_{first_name}{extension}'
                # Rename the file
                os.renames(filename, new_filename)
                count+=1
                print(count, "renamed")
                break
            

# Save the changes to the Excel file
wb.save('candidate_info.xlsx')