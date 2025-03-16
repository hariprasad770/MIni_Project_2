import pandas as pd

# Extract the data
forest_file = "C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Bird_Monitoring_Data_FOREST.XLSX"
grassland_file = "C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Bird_Monitoring_Data_GRASSLAND.XLSX"

# Function to merge all sheets into a single DataFrame
def merge_sheets(file_path):
    excel_file = pd.ExcelFile(file_path)
    merged_df = pd.DataFrame()
    
    for sheet in excel_file.sheet_names:
        temp_df = pd.read_excel(file_path, sheet_name=sheet)
        temp_df["Admin_Unit_Code"] = sheet  # Add sheet name as a column
        merged_df = pd.concat([merged_df, temp_df], ignore_index=True)
    
    return merged_df

# Merge data
forest_data = merge_sheets(forest_file)
grassland_data = merge_sheets(grassland_file)
# Handle Missing Values
for df in [forest_data, grassland_data]:
    df["ID_Method"].fillna("Unknown", inplace=True)
    df["Distance"].fillna("Unknown", inplace=True)
    df["Sex"].fillna("Undetermined", inplace=True)

# Drop columns with excessive missing values
forest_data.drop(columns=["Sub_Unit_Code", "AcceptedTSN"], inplace=True)
grassland_data.drop(columns=["Sub_Unit_Code", "AcceptedTSN", "TaxonCode"], inplace=True)

# Save Cleaned Data for Power BI
forest_data.to_csv("C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Cleaned_Bird_Forest.csv", index=False)
grassland_data.to_csv("C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Cleaned_Bird_Grassland.csv", index=False)



c_forest_data=pd.read_csv("C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Cleaned_Bird_Forest.csv")
c_grassland_data=pd.read_csv("C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Cleaned_Bird_Grassland.csv")

combined_data=pd.concat([c_forest_data,c_grassland_data],ignore_index=True)

combined_data.to_csv("C:/Users/DEEPADHARSHINI/OneDrive/Desktop/Combined_Bird_Dataset.csv", index=False)