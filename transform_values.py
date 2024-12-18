import pandas as pd
import numpy as np

def find_closest_match(lookup_dict, size, wall_inch, tolerance=0.002):
    """
    Finds the closest match within the specified tolerance.
    """
    closest_key = None
    min_diff = float('inf')  # Initialize with infinity
    
    for (key_size, key_wall_inch), _ in lookup_dict.items():
        if key_size == size:  # Ensure we're comparing the same size
            diff = abs(key_wall_inch - wall_inch)  # Calculate absolute difference
            if diff <= tolerance and diff < min_diff:  # Check within tolerance and closer than previous
                closest_key = (key_size, key_wall_inch)
                min_diff = diff
                
    return closest_key

def map_sch_desc(df, lookup_table_path):
    try:
        lookup_df = pd.read_excel(lookup_table_path)
        lookup_df['WALL INCH'] = lookup_df['WALL INCH'].apply(lambda x: "{:.3f}".format(x))
        
        # Preparing the dictionary without 'MTRL Group' as it's not in the lookup table        
        lookup_dict = {(row['NPS INCH'], float(row['WALL INCH'])):  # Ensuring 'WALL INCH' is a float
               {'ASME1': row['ASME1'], 'ASME3': row['ASME3'], 'ASME2': row['ASME2']}
               for index, row in lookup_df.iterrows()}
    except Exception as e:
        print(f"Error loading lookup table: {e}")
        return df

    def get_asme_value_based_on_matgroup(matgroup, asme_values):
        if "Grp 0" in matgroup or "Grp 1" in matgroup:
            preferred_order = ['ASME1', 'ASME3', 'ASME2']
        else:
            # If not "Grp 0" or "Grp 1", start with ASME3, then check ASME1, then fallback to ASME2
            preferred_order = ['ASME3', 'ASME1', 'ASME2']

        for key in preferred_order:
            value = asme_values[key]
            # Check if the value is valid (not NaN or a blank string)
            if not pd.isnull(value) and str(value).strip().lower() not in ['nan', '']:
                return value

        # If none of the values are valid, return np.nan as a last resort
        return np.nan
    
    def get_asme1(row):
        size_float = float(row['Size'])
        sch_desc_float = float(row['SCH Desc.'])  # Ensure this is a float for comparison
        closest_key = find_closest_match(lookup_dict, size_float, sch_desc_float)
    
        if closest_key:
            asme_values = lookup_dict[closest_key]
            value = get_asme_value_based_on_matgroup(row['MTRL Group'], asme_values)
            # If the value is explicitly 0, return a blank string
            if value == 0:
                return ""
            else:
                return value
        else:
            #logger.info(f"Closest match not found for Size: {size_float}, Wall Inch: {sch_desc_float}")
            # Check if sch_desc_float is 0 and return blank if true, else format sch_desc_float
            return "" if sch_desc_float == 0 else "{:.3f}".format(sch_desc_float)

    df['SCH'] = df.apply(get_asme1, axis=1)
    return df

def map_joint_details(df, lookup_table_path):
    # Load the joint lookup table
    joint_lookup_df = pd.read_excel(lookup_table_path)

    # Function to find the joint based on joint detail description
    def find_joint(detail):
        for _, row in joint_lookup_df.iterrows():
            if row['Lookup Key'] in detail:
                return row['Joint']
        return ""  # Return an empty string if no match is found

    # Apply the function to create the 'Joint 2' column
    df['Joint 2'] = df['Joint Detail'].apply(find_joint)
    
    return df

def map_base_material(df, lookup_table_path):
    # Load the joint lookup table
    mtrl_lookup_df = pd.read_excel(lookup_table_path)

    # Function to find the joint based on joint detail description
    def find_spec(detail):
        for _, row in mtrl_lookup_df.iterrows():
            if row['PIPE CLASS'] in detail:
                return row['BASE MATERIAL']
        return ""  # Return an empty string if no match is found

    # Apply the function to create the 'Joint 2' column
    df['Material'] = df['Pipe Spec'].apply(find_spec)
    
    return df
