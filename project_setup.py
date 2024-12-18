import re
from simpledbf import Dbf5
import pandas as pd

def sanitize(df):
    return df.map(lambda x: ''.join(filter(lambda c: c.isprintable(), str(x))) if isinstance(x, str) else x)

def get_unique_formats(file_path, column_name):
    # Function to replace specific parts with placeholders
    def generalize_format(s):
        if not isinstance(s, str):
            return ''
        
        # Replace numbers and single letters with placeholders
        s = re.sub(r'\d+', 'N', s)
        s = re.sub(r'-[A-Z]-', '-L-', s)
        # Replace consecutive letters with a single 'L'
        s = re.sub(r'[A-Z]{2,}', 'L', s)
        return s
    
    try:
        # Read the DBF file
        dbf = Dbf5(file_path)
        df = dbf.to_dataframe()
        df = sanitize(df) # Clean contents
        
        # Format the dataframe
        df = df[['CONTROLNO', 'SP4', 'SP5', 'SP7', 'SP36', 'SP38','SP37']].copy()
        df.rename(columns={'CONTROLNO': 'Spool', 'SP4': 'MAT. Group', 'SP5': 'VT', 'SP7': 
                           'AREA', 'SP36': 'Ref Drawing', 'SP37': 'Drawing Rev.' ,'SP38': 'Spool/Piece Mark' }, inplace=True)

        # Remove the first row
        df = df.iloc[1:]

        # Check if the specified column exists
        if column_name not in df.columns:
            raise ValueError(f"Column '{column_name}' not found in the DBF file.")

        # Extract data from the specified column
        string_list = df[column_name].tolist()

        # Apply the generalization to each string and get unique formats
        unique_formats = set(map(generalize_format, string_list))
        return list(unique_formats)

    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return []

    # Apply the generalization to each string and get unique formats
    unique_formats = set(map(generalize_format, string_list))
    return list(unique_formats)

# Example usage
file_path = r"C:\Users\cmccall\source\repos\cmccall95\FabManager_To_QB\Shop\30501\SpoolData.dbf" #'path/to/your/dbf/file.dbf'
column_name = 'Ref Drawing'

unique_formats = get_unique_formats(file_path, column_name)
for format in unique_formats:
    print(format)