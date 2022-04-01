# Importing libraries
import pandas as pd
import openpyxl
import xlwings as xw
import os
import glob

# Paths 
path_kti = "/Users/adap.anjali/Desktop/SC Affluent/Test Folder/Updates KTI" # Path to KTI folder
path_sr = "/Users/adap.anjali/Desktop/SC Affluent/Test Folder/Updates SR"   # Path to SR folder
path_Updates = "/Users/adap.anjali/Desktop/SC Affluent/Test Folder/Intermediate Updates" # Path to Intermediate Updates folder
path_List = "/Users/adap.anjali/Desktop/SC Affluent/Test Folder/Main Docs/Main List.xlsx" # Path to Main Combined List
path_Keys = "/Users/adap.anjali/Desktop/SC Affluent/Test Folder/Main Docs/Keywords.xlsx" # Path to Keywords List 

# Rename the sheet in each of the KTI workbooks to name of the document 
for filename in os.listdir(path_kti):
    if not filename.startswith('.') and os.path.isfile(os.path.join(path_kti, filename)):
        ss = openpyxl.load_workbook(path_kti + "/" + filename) # Read the documents 
        ss_sheet = ss["Keywords"] # Transfer sheet to variable 
        ss_sheet.title = filename[:-5] # Change name of sheet 
        os.remove(path_kti + "/" + filename) # Remove old document from the file
        ss.save(path_kti + "/" + filename) # Save the new document into the file

# Rename the sheet in each of the SR workbooks to name of the document
for filename in os.listdir(path_sr):
    if not filename.startswith('.') and os.path.isfile(os.path.join(path_sr, filename)):
        ss = openpyxl.load_workbook(path_sr + "/" + filename) # Read the documents 
        ss_sheet = ss["Keywords"] # Transfer sheet to variable 
        ss_sheet.title = filename[:-5] # Change name of sheet 
        os.remove(path_sr + "/" + filename) # Remove old document from the file
        ss.save(path_sr + "/" + filename) # Save the new document into the file

# Merge respective documents from KTI and SR folders (final documents will be in new folder called Updates)
for filename_kti in os.listdir(path_kti):
    if not filename_kti.startswith('.') and os.path.isfile(os.path.join(path_kti, filename_kti)):
        for filename_sr in os.listdir(path_sr):
            if not filename_sr.startswith('.') and os.path.isfile(os.path.join(path_sr, filename_sr)):
                if filename_kti == filename_sr:
                    df_kti = pd.read_excel(path_kti + "/" + filename_kti)
                    df_kti = df_kti.iloc[:,[0, 13]] # Keep only keywords and search volumes 
                    df_sr = pd.read_excel(path_sr + "/" + filename_sr)
                    df_sr = df_sr.iloc[:,[1, 4]] # Keep only keywords and search volumes
                    df_sr.columns = df_kti.columns # Rename columns
                    df_merge = pd.concat([df_kti, df_sr]) # Merge the two dataframes
                    df_merge.drop_duplicates(subset = "Keywords", keep = "last", inplace = True)
                    df_merge = df_merge.where(pd.notnull(df_merge), None)
                    writer = pd.ExcelWriter(path_Updates + "/" + filename_kti)
                    df_merge.to_excel(writer, index = False, sheet_name = filename_kti[:-5])
                    writer.save()
                else: 
                    df_kti = pd.read_excel(path_kti + "/" + filename_kti)
                    df_kti = df_kti.iloc[:,[0, 13]] # Keep only keywords and search volumes 
                    writer = pd.ExcelWriter(path_Updates + "/" + filename_kti)
                    df_kti.to_excel(writer, index = False, sheet_name = filename_kti[:-5])
                    writer.save()

# Remove documents from Updates KTI
for filename_kti in os.listdir(path_kti):
    os.remove(path_kti + "/" + filename_kti) 

# Remove documents from Updates SR
for filename_sr in os.listdir(path_sr):
    os.remove(path_sr + "/" + filename_sr) # Remove old documents from the file

# Join all the worksheets into one, creating a workbook with 14 sheets
excel_files = glob.glob(os.path.join(path_Updates, "*.xlsx"))

with xw.App(visible=False) as app:
    combined_wb = app.books.add()
    for excel_file in excel_files:
        wb = app.books.open(excel_file)
        for sheet in wb.sheets:
            sheet.copy(after = combined_wb.sheets[0])
        wb.close()
    combined_wb.sheets[0].delete()
    combined_wb.save(path_Updates + "/combined.xlsx")
    combined_wb.close()

# Remove all other files from the folder except "combined.xlsx"
for filename in os.listdir(path_Updates):
    if not filename.startswith('.') and os.path.isfile(os.path.join(path_Updates, filename)):
        if filename != "combined.xlsx":
            os.remove(path_Updates + "/" + filename)

# Reading the new Combined Updates, the Main Combined List and the Keywords List
df_Updates = pd.read_excel(path_Updates + "/" + "combined.xlsx", sheet_name= None)
df_List = pd.read_excel(path_List)
df_Keys = pd.read_excel(path_Keys, sheet_name= None)

# Dataframes with Updated search volumes in a list
Updates = []
for name, sheet in df_Updates.items():
    Updates.append(sheet)

# Dataframes with Main keywords and other columns that need to be merged with the respective Updates dataframes in a List 
Keys = []
for name, sheet in df_Keys.items():     
    Keys.append(sheet)

# Order of dataframes in the list of dataframes
all_Updates = list(df_Updates.keys())
all_Keys = list(df_Keys.keys())

# Changing the Keywords to lower case and setting index as Keywords for merging 
for j in range(len(Updates)):
    Updates[j]["Keywords"] = Updates[j]["Keywords"].str.lower()
    Updates[j].set_index("Keywords", inplace= True)

for i in range(len(Keys)):
    Keys[i]['Keywords'] = Keys[i]['Keywords'].str.lower()
    Keys[i].set_index("Keywords", inplace= True)

# Merge Keys and Updates based on the keywords and by name of the sheet if matching
for i in range(len(Keys)):
    for j in range(len(Updates)):
        if all_Keys[i] == all_Updates[j]:
            Updates[j] = Keys[i].join(Updates[j])

# Put all the necessary columns into every dataframe in the folder 
for i in range(len(Updates)):
    month = Updates[i].columns[-1][15:18]
    year = Updates[i].columns[-1][19:23]
    Updates[i]["Search Volume"] = Updates[i].iloc[:, -1]
    Updates[i].drop(columns=Updates[i].columns[-2], inplace= True)
    Updates[i]["Month"] = month
    Updates[i]["Year"] = year

    if "EN Translations" not in Updates[i].columns:
        Updates[i]["EN Translations"] = "-"
    if "Local Full Segment Name" not in Updates[i].columns:
        Updates[i]["Local Full Segment Specific Name"] = "-"
    if "English Full Segment Specific Name" not in Updates[i].columns:
        Updates[i]["English Full Segment Name"] = "-"
    if "Local Bank Name" not in Updates[i].columns:
        Updates[i]["Local Bank Name"] = "-"
    if "Bank Specific Name" not in Updates[i].columns:
        Updates[i]["Bank Specific Name"] = "-"

    Updates[i].reset_index(inplace= True)
    Updates[i] = Updates[i][["Bank", "Local Bank Name", "Segment", "English Full Segment Name", "Local Full Segment Specific Name", "Bank Specific Name", "Keywords Grouping", "EN Translations", "Keywords", "Country", "Language", "Month", "Year", "Search Volume"]]

# Merging all sheets into one sheet
df_Updates = pd.concat(Updates)

# Changing data type
df_List["Year"] = pd.to_numeric(df_List["Year"])
df_Updates["Year"] = pd.to_numeric(df_Updates["Year"])

# Final merge of all dataframes
df_final = pd.concat([df_Updates, df_List])

# Remove duplicates
df_final.drop_duplicates(subset=["Bank", "Local Bank Name", "Segment", "English Full Segment Name", "Local Full Segment Specific Name", "Bank Specific Name", "Keywords Grouping", "EN Translations", "Keywords", "Country", "Language", "Month", "Year"], keep= 'last', inplace= True, ignore_index= False)

# Convert NaN to None values 
df_final = df_final.where(pd.notnull(df_final), None)

# Remove the old Main Combined List
os.remove(path_List)

# Remove the old "combined.xlsx"
os.remove(path_Updates + "/" + "combined.xlsx")

# Write new Main Combined List into the file
writer = pd.ExcelWriter(path_List)
df_final.to_excel(writer, index = False, sheet_name = "Keywords")
writer.save()

# Check end process
print("Fin")