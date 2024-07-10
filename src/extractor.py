import pandas as pd

def read_excel(file_name: str, sheet_name: str, skiprows: int):
    return pd.read_excel(file_name,sheet_name=sheet_name,skiprows=skiprows)

source_df = read_excel('TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',sheet_name='Source Layout',skiprows=1)
target_df = read_excel('TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',sheet_name='Target Layout',skiprows=1)
transform_df = read_excel('TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls',sheet_name='Transformation Logic',skiprows=0)

#############################################################################################################################
import pandas as pd
from fuzzywuzzy import process

def read_excel_file(file_path, sheet_name=None):
    """
    Read an Excel file and return a pandas DataFrame.
    """
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
    except Exception:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd', skiprows=1)
    return df

def fuzzy_map_columns(df, column_mappings, threshold=20):
    """
    Map the columns of a DataFrame based on fuzzy matching of column names.
    """
    actual_columns = list(df.columns)
    for standard_field, possible_columns in column_mappings.items():
        best_match, score = process.extractOne(standard_field, actual_columns)
        print(standard_field, best_match, score)
        if score >= threshold:
            df[standard_field] = df[best_match]
    return df

def extract_and_normalize(df, standard_fields):
    """
    Extract and normalize the specified standard fields from the DataFrame.
    """
    df = df[standard_fields]
    return df

def save_to_excel(df, output_path):
    """
    Save the DataFrame to an Excel file.
    """
    df.to_excel(output_path, index=False)

def process_excel_files(file_paths, column_mappings, standard_fields, output_path):
    """
    Process multiple Excel files to extract and normalize specified fields.
    """
    all_data = []
    for file_path in file_paths:
        df = read_excel_file(file_path)
        df = fuzzy_map_columns(df, column_mappings)
        normalized_data = extract_and_normalize(df, standard_fields)
        all_data.append(normalized_data)

    combined_data = pd.concat(all_data, ignore_index=True)
    save_to_excel(combined_data, output_path)

# Example usage
file_paths = ['TS_m_Xref_Delta_Update_XREF_CDH_BIG_DELTA_STI.xls'] 
# standard_fields = ['Source_System', 'Source_Table_Name', 'Source_Column_Name', 'Source_Data_Type', 'Source_Filter_Condition', 'Destination_Table_Name', 'Destination_System', 'Destination_Column_Name', 'Destination_Data_Type', 'Primary_Key', 'Foreign_Key', 'Is_Nullable', 'Transformation_Type', 'Transformation_Logic', 'Joining_Table_Name', 'Joining_Type', 'Joining_Logic', 'Sample_Values', 'Comments'] 
standard_fields = ['Source_Table_Name' ]
column_mappings = {
    'Source_Column_Name': ['Source Column Name', 'Field_Name', 'src_col_name', 'src col name']
                          }
# output_path = 'combined_output.xlsx'
# process_excel_files(file_paths, column_mappings, standard_fields, output_path) 
# ['Source Table/File Name', 'Source Column Name',
#        'Transformation logic in detail/steps', 'Target Table Name',
#        'Target Column Name', 'Comments & Issues']
all_data = []
for file_path in file_paths:
    df = read_excel_file(file_path,sheet_name='Source Layout')
    df = fuzzy_map_columns(df, column_mappings)
    print(df)
    # normalized_data = extract_and_normalize(df, standard_fields)
    # all_data.append(normalized_data)
#############################################################################################################################
# combined_data = pd.concat(all_data, ignore_index=True)
# save_to_excel(combined_data, output_path)
