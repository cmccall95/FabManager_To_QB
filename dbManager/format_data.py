import pandas as pd
from logger_setup import logger
from transform_values import map_sch_desc, map_base_material, map_joint_details, find_closest_match

'''
File for formatting Spool & Weld table data
'''


def format_spools(df, regex_patterns=None, select_fields=None): #OCI Setup JOBS: 30489

    """
    Format spools dataframe and set NDE groups based on Job column values.
    """
    
    # Initialize a list to store non-convertible values
    non_convertible_values = []
    
    # Initialize a dataframe to store non-convertible rows
    non_convertible_rows_df = pd.DataFrame()

    # Initialize NDE Group column if not present
    if 'NDE Group' not in df.columns:
        df['NDE Group'] = ''

    if 'Series' not in df.columns:
        df['Series'] = ''

    if 'Sheet' not in df.columns:
        df['Sheet'] = ''

    if 'Line ID' not in df.columns:
        df['Line ID'] = ''

    
    # Select and Reorder fields from dataframe
    if select_fields: 
        df = df[select_fields]

    # Ensure job numbers in 'Job' column end with "-"
    df['Job'] = df['Job'].apply(lambda x: x + "-" if not str(x).endswith("-") else x)

    # NDE Group mapping
    nde_group_mapping = {
        '30504-': 'AA1B',
        '30507P-': 'A',
        '30507T-': 'A'
        # Add more mappings as needed
    }

    # Apply NDE Group based on job number in 'Job' column
    df['NDE Group'] = df['Job'].map(nde_group_mapping).fillna(df['NDE Group'])

    # Fill blank Drawing Rev. with '0'
    df['Drawing Rev.'] = df['Drawing Rev.'].fillna('0')

    # Function to try conversion and catch errors
    def try_convert_to_int(value, row):
        # Initialize a dataframe to store non-convertible rows
        non_convertible_rows_df = pd.DataFrame()
        
        try:
            return int(float(value.lstrip('0')))
        except ValueError:
            logger.warning(f"Could not convert value: {value}")
            non_convertible_values.append(value)
            try:
                #non_convertible_rows_df.append(row)
                non_convertible_rows_df = pd.concat([non_convertible_rows_df, row.to_frame().T], ignore_index=True)
            except Exception as e:
                logger.error(f"Could not append to dataframe (Non Convertible Spools) {e}", exc_info=True)
            return None  # Or a placeholder value indicating conversion failure
    
    # Apply conversion with error handling
    df['Spool'] = df.apply(lambda row: try_convert_to_int(row['Spool'], row), axis=1)


    # Print the non-convertible values
    if non_convertible_values:
        print(f"({len(non_convertible_values)})Non-convertible values in 'Spool' column: {non_convertible_values}")

    # Drop rows with non-convertible values
    df = df.dropna(subset=['Spool'])

    # Ensure 'Ref Drawing' is of string type
    df['Ref Drawing'] = df['Ref Drawing'].astype(str)

    # Extract 'Series' and 'Sheet' from 'Ref Drawing' based on the logic provided
    df['Series'] = df['Ref Drawing'].apply(lambda x: x.rsplit('-', 1)[0] if '-' in x else '')
    df['Sheet'] = df['Ref Drawing'].apply(lambda x: x.rsplit('-', 1)[1] if '-' in x else '')

    # Export non-convertible rows to Excel
    filename="Dropped Rows - Spool.xlsx"
    try:
        non_convertible_rows_df.to_excel(filename, index=True)
    except Exception as e:
        logger.error(f"Could not export to dataframe {filename}")
        
    return df

def format_welds(df, tmp_spools_df, select_fields=None):
    logger.info("Accessed format_welds_df")
    df = df.copy()
    spools_df = tmp_spools_df.copy()

    non_convertible_values = []
    dropped_rows_df = pd.DataFrame()  # DataFrame to store dropped rows

    # Convert Field Names
    # Initialize 'Weld ID' and 'Related Record' column if not present
    if 'Weld ID' not in df.columns:
        df['Weld ID'] = ''

    if 'Related Record' not in df.columns: # Add new columns 'qb_id' and 'JOB' with empty values
        df['Related Record'] = ''

    if 'Material' not in df.columns:
        df['Material'] = ''

    if 'Pipe Spec' not in df.columns:
        df['Pipe Spec'] = ''


    print("\n\nIncoming Weld DF Fields: \n", df.columns)


    # Ensure job numbers in 'Job' column end with "-"
    df['Job'] = df['Job'].apply(lambda x: x + "-" if not str(x).endswith("-") else x)

    # Function to check if Spool can be converted to an integer
    def is_convertible_to_int(value):
        try:
            int(value)
            return True
        except ValueError:
            return False

    # Function to try conversion and catch errors
    def try_convert_to_int(value):
        try:
            return int(float(value.lstrip('0')))
        except ValueError:
            # logger.warning(f"Could not convert value: {value}") ORIGINAL
            # non_convertible_values.append(value)
            
            # NEW LINES 6-27
            non_convertible_values.append(value)
            row = pd.Series({'Spool': value})  # Create a row to append
            non_convertible_rows_df = pd.concat([non_convertible_rows_df, row.to_frame().T], ignore_index=True)
            
            return None  # Or a placeholder value indicating conversion failure

    # Apply the function to create a mask for rows with valid Spool
    valid_controlno_mask = df['Spool'].astype(str).str.lstrip('0').apply(is_convertible_to_int)
    
    # Store invalid Spool rows and filter the DataFrame
    dropped_rows_df = pd.concat([dropped_rows_df, df[~valid_controlno_mask]])
    df = df[valid_controlno_mask]

    # Filter the DataFrame based on the mask
    df = df[valid_controlno_mask]

    # Now it's safe to convert Spool to int
    logger.info(f"COLUMNS: {df.columns}")
    logger.info(f"Length: {len(df)}")
    df['Spool'] = df['Spool'].astype(str).str.lstrip('0').astype(int)

    # Strip Leading Zeroes from wekd_df before merging
    # welds_df['Spool'] = welds_df['Spool'].str.lstrip('0')

    # Perform merge to get Pipe Spec
    # First, temporarily rename Pipe Spec in welds_df to avoid the duplicate
    df['Pipe Spec_original'] = df['Pipe Spec']
    df.drop('Pipe Spec', axis=1, inplace=True)

    # Now do the merge
    df = df.merge(
        spools_df[['Job', 'Spool', 'Pipe Spec']], 
        on=['Job', 'Spool'], 
        how='left'
    )

    # If merge failed for any rows, use the original value
    df['Pipe Spec'] = df['Pipe Spec'].fillna(df['Pipe Spec_original'])

    # Clean up by dropping the temporary column
    df.drop('Pipe Spec_original', axis=1, inplace=True)
    
    # Generate 'Weld ID' column
    df['Weld ID'] = df['Spool'].astype(str) + '-' + df['Weld Label'].astype(str)
    

    #####################
    logger.info("Checking for highest revision, dropping duplicates, and welds not in latest revision.")
    
    # Determine the latest revision for each Spool
    latest_revisions = df.groupby('Spool')['Revision'].transform('max')
    
    # Filter the DataFrame to keep only the latest revisions
    dropped_rows_df = pd.concat([dropped_rows_df, df[df['Revision'] != latest_revisions]])
    df = df[df['Revision'] == latest_revisions]
    
    logger.info("Revision check complete and rows dropped.")
    
    # Remove '"' from 'SIZE_I' and ensure no leading or trailing zeroes
    df['Size'] = df['Size'].str.replace('"', '').astype(float).astype(str).str.strip('0').str.rstrip('.')
    
    # Replace 'BR' with 'BRANCH' and 'FW' with 'ATTACH' in 'Joint'
    df['Joint'] = df['Joint'].replace({'BR': 'BRANCH', 'FW': 'ATTACH'})

    # Record initial row count
    initial_row_count = df.shape[0]

    # Function to check if a value is NaN, blank or a variation of "nan"
    def is_invalid(value):
        value_str = str(value).strip().lower()
        return pd.isna(value) or value_str == '' or value_str == 'nan'
    
    # Store invalid Weld Label rows and drop them from the DataFrame
    try:
        invalid_weldlabel_mask = df['Weld Label'].apply(is_invalid)
        dropped_rows_df = pd.concat([dropped_rows_df, df[invalid_weldlabel_mask]])
        df = df[~invalid_weldlabel_mask]
    except Exception as e:
        logger.error(f"Could not create df_invalid: {e}", exc_info=True)
        # Drop rows where 'Weld Label' is invalid
        df = df[~df['Weld Label'].apply(is_invalid)]

    # Calculate the number of rows dropped
    final_row_count = df.shape[0]
    rows_dropped = initial_row_count - final_row_count
    logger.info(f"{rows_dropped} rows were dropped due to invalid values in 'Spool' or 'Weld Label'")
    
    # Export dropped rows to Excel
    try:
        filename = 'Dropped Rows-welds.xlsx'
        dropped_rows_df.to_excel(filename, index=False)
    except Exception as e:
        logger.error(f"Could not export {filename}. Error: {e}", exc_info=True)

    # # Select only the desired columns and rename them
    # df = df[['Related Record', 'Job', 'Revision','Weld ID', 'SPEC', 'SIZE_I', 'Joint', 'WDESCRIPT', 'WALL_I', 'MATGROUP','Spool', 'Weld Label', 'DIAINCH']]
    # df.rename(columns={'SIZE_I': 'Size', 'Joint': 'Joint', 'WDESCRIPT': 'Joint Detail' ,'SPEC': 'Pipe Spec', 'WALL_I':'SCH Desc.', 'MATGROUP':'MTRL Group','Spool': 'Spool', 'Revision':'Revision', 'DIAINCH':'DiaInch'}, inplace=True)

    # Ensure 'SCH Desc.' column is a string before applying the function
    df['SCH Desc.'] = df['SCH Desc.'].astype(str).apply(extract_float_from_parentheses)
    
    # Then convert to float for lookup
    df['SCH Desc.'] = df['SCH Desc.'].astype(float)
    
    # Format 'SCH Desc.' column as a string with the format "0.000"
    df['SCH Desc.'] = df['SCH Desc.'].map("{:0.3f}".format)
    
    # Call map_sch_desc after you have the 'SCH Desc.' as floats
    map_sch_desc(df, 'Pipe Dimensions and Weights.xlsx')
    
    # Call map_joint_details to add the 'Joint 2' column based on 'Joint Detail'
    df = map_joint_details(df, 'Joint Lookup.xlsx')

    # Then convert to float for lookup
    df['DiaInch'] = df['DiaInch'].astype(float)
    
    # Call map_base_material to add the 'BASE MATERIAL' column based on 'PIPE CLASS'
    # if job_number == "30489-":
    #     df = map_base_material(df, 'assets/30489 Material.xlsx')
    
        
    # # Replace '-' with a blank in job_number
    # cleaned_job_number = job_number.replace("-", "")

    # Replace '_job_number_' with cleaned_job_number in spool_detail_path
    # spool_detail_path = f"app_files/{cleaned_job_number} Spool Detail.xlsx"

    # Map statuses
    # Apply conversion with error handling
    df['Spool'] = df['Spool'].astype(str).apply(try_convert_to_int)

    # Select and Reorder fields from dataframe
    print("\n\nWELD DF Fields: \n", df.columns)

    if select_fields: 
        df = df[select_fields]

    #updated_df = add_status_column(df, spool_detail_path)
    
    return df, spools_df

def extract_float_from_parentheses(s):
    s = str(s)
    start = s.find('(')
    end = s.find(')')
    if start != -1 and end != -1 and start < end:
        return float(s[start+1:end])
    return "0.0"

def add_status_column(df, file_path):
    # Function to try conversion and catch errors
    def try_convert_to_int(value):
        try:
            return int(float(value.lstrip('0')))
        except ValueError:
            #logger.error(f"Could not convert value: {value}")
            return None  # Or a placeholder value indicating conversion failure

    # Load the Excel file
    status_df = pd.read_excel(file_path)

    # Strip spaces from column names
    status_df.columns = status_df.columns.str.strip()

    # Adjust columns if necessary, e.g., if the names are known to have extra spaces or similar issues
    # Ensure columns you will use are properly named here after stripping
    use_columns = ['Spool', 'Current_Status']  # Example, adjust as needed

    # Keep only necessary columns to avoid potential errors
    status_df = status_df[use_columns]

    # Convert the 'Spool_ID' column in the Excel DataFrame
    status_df['Spool'] = status_df['Spool ID'].astype(str).apply(try_convert_to_int)

    # Drop rows where conversion failed
    status_df = status_df.dropna(subset=['Spool ID'])

    # Create a dictionary for mapping 'Spool' to 'Status'
    status_map = status_df.set_index('Spool ID')['Current Status'].to_dict()

    # Map the 'Status' to the existing DataFrame
    df['Status'] = df['Spool'].map(status_map).fillna('DNF')