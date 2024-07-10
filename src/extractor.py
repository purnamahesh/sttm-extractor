import pandas as pd

def read_excel(file_name: str, sheet_name: str, skiprows: int):
    return pd.read_excel(file_name,sheet_name=sheet_name,skiprows=skiprows)

source_df = read_excel('TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',sheet_name='Source Layout',skiprows=1)
target_df = read_excel('TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',sheet_name='Target Layout',skiprows=1)
transform_df = read_excel('TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',sheet_name='Transformation Logic',skiprows=0)

#############################################################################################################################
import pandas as pd
from fuzzywuzzy import process

def fuzzy_map_columns(df, column_mappings, threshold=20):
    actual_columns = list(df.columns)
    for standard_field, possible_columns in column_mappings.items():
        best_match, score = process.extractOne(standard_field, actual_columns)
        print(standard_field, best_match, score)
        if score >= threshold:
            df[standard_field] = df[best_match]
    return df


column_mappings = {
    'Source_Table_Name':['table_name', 'Table Name', 'tbl_name'], 
    'Source_Column_Name': ['Source Column Name', 'Field_Name', 'src_col_name', 'src col name'],
    'Transformation_Logic': ['Transformation'], 
    'Destination_Table_Name':['Target Table Name', 'Tgt Table Name', 'tgt_tbl_name'],
    'Destination_Column_Name':['Destination Column Name', 'Destination Field Name', 'tgt_col_name', 'tgt col name', 'Target Column Name'], 
    'Comments': ['comments'],
    # 'Destination_System':[], 
    # 'Destination_Data_Type':['Target Data Type', 'tgt data type'], 
    # 'Primary_Key': [''], 
    # 'Foreign_Key': [], 
    # 'Is_Nullable': ['Null / Not Null'], 
    # 'Transformation_Type': [], 
    # 'Joining_Table_Name': [], 
    # 'Joining_Type': [], 
    # 'Joining_Logic': [], 
    # 'Sample_Values': [], 
    # 'Source_Data_Type':['Data Type'], 
    # 'Source_Filter_Condition':[]
}

file_paths = ['TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls'] 
for file_path in file_paths:
    df = read_excel(
        'TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',
        sheet_name='Transformation Logic',
        skiprows=0
    )
    cols = df.columns
    df = fuzzy_map_columns(df, column_mappings)
    for col in df.columns:
        if col in cols: 
            df.drop(col, axis=1, inplace=True)
df
