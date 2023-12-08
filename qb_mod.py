import pandas as pd
import os, sys, requests, json
from dotenv import load_dotenv

print("\n\n>>>MODULE ACTIVATED")

HOSTNAME = os.getenv('HOSTNAME')
TOKEN = os.getenv('TOKEN')
USER_AGENT = os.getenv('USER_AGENT')
APP_ID = os.getenv('APP_ID')
WELD_LOG_ID = os.getenv('WELD_LOG_ID')
JOB_STENCIL_LOTS_ID = os.getenv('JOB_STENCIL_LOTS_ID')
JOB_STENCIL_OPTION_LOTS_ID = os.getenv('JOB_STENCIL_OPTION_LOTS_ID')
TPSL_ID = os.getenv('TPSL_ID')
JOB_SETUP_TABLE = os.getenv('JOB_SETUP_TABLE')
WELDER_CONT_TABLE = os.getenv('WELDER_CONTINUITY_TABLE')
JOB_NDE = os.getenv('JOB_NDE')
QUICKBASE_API_URL = "https://api.quickbase.com/v1" 

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)



#Main QB call for app
def get_table_data(qb_fields, job_number, table_id):
    print("\nFetching data...")
    # Get table info and update field mapping
    headers = {
        "QB-Realm-Hostname": HOSTNAME,
        "User-Agent": USER_AGENT,
        "Authorization": TOKEN,
        'Content-Type': 'application/json'
        }

        # Construct where clause using tpsl_data
    where_clause = f"{{6.EX.{job_number}}}"

    query_body = {
        "from": table_id, 
        "select": qb_fields, 
        "where": where_clause,
        "orderBy": [],
        "skip": 0,
    }

    response = requests.post(
        f"{QUICKBASE_API_URL}/records/query",
        headers=headers,
        json=query_body
    )
    
    if response.status_code == 200:
        print("Request Successful... " + str(response.status_code))
        data = response.json()
        records = data.get('data', [])

        #Extract values from dictionaries
        for record in records:
            for field_id in record.keys():
                if isinstance(record[field_id], dict) and 'value' in record[field_id]:
                    record[field_id] = record[field_id]['value']

        # Get the number of records
        print("COUNT OF RECORDS: " + str(len(records)) + "\n")
        data['data'] = records

        df = pd.DataFrame(records)

        # Specify the filename and sheet name
        filename = 'Original Dataframe.xlsx'
        sheet_name = 'Original Dataframe'

        ## Export the DataFrame to an Excel file
        #df.to_excel(filename, sheet_name=sheet_name, index=False)
        #print("Dataframe Exported....")

        return df
        
    else:
        return {f"Resonse: {str(response.status_code)} error_get_table_data": "Unable to fetch table data"}
    print("Get existing QB table here")

#Exports shop 
def export_weld_log_to_excel(table_id):
    print("\nFetching data from WELD_LOG_ID table...")
    # load_dotenv()
    
    headers = {
        "QB-Realm-Hostname": HOSTNAME,
        "User-Agent": USER_AGENT,
        "Authorization": TOKEN,
        'Content-Type': 'application/json'
    }

    # Construct where clause for specific job numbers
    where_clause = "{'6'.EX.'30489-'}OR{'6'.EX.'30496-'}OR{'6'.EX.'30497-'}OR{'6'.EX.'30498-'}"

    query_body = {
        "from": table_id,
        "select": ['3','6','7','15', '20','21' ],  # Fetch all fields
        "where": where_clause,
        "orderBy": [],
        "skip": 0,
    }

    response = requests.post(
        f"{QUICKBASE_API_URL}/records/query",
        headers=headers,
        json=query_body
    )
    
    print(f"headers: {headers} \n\n query_body: {query_body}" )
    
    if response.status_code == 200:
        print("Request Successful... " + str(response.status_code))
        data = response.json()
        records = data.get('data', [])

        # Extract values from dictionaries
        for record in records:
            for field_id, field_data in record.items():
                if isinstance(field_data, dict) and 'value' in field_data:
                    record[field_id] = field_data['value']

        print("COUNT OF RECORDS: " + str(len(records)) + "\n")

        df = pd.DataFrame(records)

        # Specify the filename and sheet name
        filename = 'Weld_Log_Selected_Jobs.xlsx'
        sheet_name = 'Weld Log Data'

        # Export the DataFrame to an Excel file
        df.to_excel(filename, sheet_name=sheet_name, index=False)
        print(f"Data exported to {filename}")

        return df
        
    else:
        print(f"Error fetching data: {response.status_code}")
        return None

