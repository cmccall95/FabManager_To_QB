server_name='EXLBRD-V-SQL1'


'''
Database Id: EXLBRD-V-SQL1.FabManager

Server Name: EXLBRD-V-SQL1
DataBase: FabManager

Tables:

> FabManager.Tables

Spools: dbo.tbllISOSpools (Same data as 'SpoolData') Seems to also have Statuses
Also has dbo.tblSpools with very similar informaation

Welds:

'''

from msilib.text import tables
import pyodbc
import pandas as pd
print(pyodbc.drivers())
import urllib.parse
from sqlalchemy import create_engine
from contextlib import contextmanager
from dbManager.mapping_files import spool_column_mapping, weld_column_mapping
from logger_setup import logger


def clean_string_column(col):
    if col.dtype == "object":
        return col.astype(str).str.strip()
    return col

def create_db_engine(db_name):
    """
    Creates a SQLAlchemy engine for database connection.
    
    Args:
        db_name (str): Name of the database to connect to
        
    Returns:
        sqlalchemy.engine.Engine: Database engine
    """
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server_name};"
        f"DATABASE={db_name};"
        f"Trusted_Connection=yes;"
        f"Encrypt=yes;"
        f"TrustServerCertificate=yes"
    )
    
    return create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

@contextmanager
def db_connection(db_name):
    """
    Context manager for database connections.
    Ensures proper handling of connections and automatic cleanup.
    
    Args:
        db_name (str): Name of the database to connect to
        
    Yields:
        Connection: SQLAlchemy connection object
    """
    engine = create_db_engine(db_name)
    connection=None

    try:
        connection = engine.connect()
        yield connection

    except Exception as e:
        print("Error connecting to database: ", e)

    finally:
        if connection:
            connection.close()

        engine.dispose()

def get_welds_data(spools_df, db_name="FabManager"):
    """
    Retrieves weld data from FabManager database matching specific spool and job combinations.
    
    Args:
        spools_df (pandas.DataFrame): DataFrame containing spool data with 'Spool' and 'Job' columns
        db_name (str): Name of the database. Defaults to "FabManager"
        
    Returns:
        pandas.DataFrame: DataFrame containing the matching weld records
    """
    
    # Create conditions for each spool-job combination
    conditions = []
    for _, row in spools_df.iterrows():
        job = row['Job'].rstrip('-')  # Remove trailing dash
        spool = f"{int(row['Spool']):06d}"  # Format as 6-digit string
        conditions.append(f"(WLD_JobID = '{job}' AND WLD_ControlNo = '{spool}')")
    
    # Join conditions with OR
    conditions_str = " OR ".join(conditions)
    
    query = f"""
    SELECT *
    FROM dbo.tblWelds
    WHERE {conditions_str}
    """
    
    try:
        with db_connection(db_name) as conn:
            df = pd.read_sql(query, conn)
            
            
            # Clean leading/trailing spaces
            df = df.apply(clean_string_column)

            # Normalize the column/field name values
            df.rename(columns=weld_column_mapping, inplace=True)

            print(f"Weld Query Successful: {query}")
                
            return df
            
    except Exception as e:
        logger.error(f"Error executing welds query: {e}", exc_info=True)
        print(f"\n\nWelds Query: {query}")
        return pd.DataFrame() # Return blank dataframe

def get_spools_data(db_name="FabManager", job_numbers=None, all_jobs_after=None, active_jobs=None):

    # active_jobs = ['30489', '30501', '30506', '30507P']

    """
    Retrieves spool data from FabManager database with optional job number filtering.
    
    Args:
        db_name (str): Name of the database. Defaults to "FabManager"
        job_numbers (list): Optional list of job numbers as strings (e.g. ['30489', '30503-S'])
        all_jobs_after (str/int): Optional job number to get all jobs after (e.g. '30489' or 30489)
        
    Returns:
        pandas.DataFrame: DataFrame containing the spool records
    """


    # Base query
    query = """
    WITH StatusCTE AS (
        SELECT *,
            CASE 
                WHEN LTRIM(RTRIM(Voided)) LIKE '%Y%' THEN 'Void'
                WHEN On_Hold != '1900-01-01' AND Off_Hold = '1900-01-01' THEN 'On Hold'
                WHEN Shipped != '1900-01-01' THEN 'Shipped'
                WHEN Ship_Paint != '1900-01-01' THEN 'Ship Paint'
                WHEN OutsideService != '1900-01-01' THEN 'Outside Service'
                WHEN RdyToShip != '1900-01-01' THEN 'Ready to Ship'
                WHEN QC_Complete != '1900-01-01' THEN 'QC Complete'
                WHEN InHydro != '1900-01-01' THEN 'In Hydro'
                WHEN NDE != '1900-01-01' THEN 'NDE'
                WHEN QC != '1900-01-01' THEN 'QC'
                WHEN Check_Out != '1900-01-01' THEN 'Check Out'
                WHEN WeldOut != '1900-01-01' THEN 'Weld Out'
                WHEN Final_Fit != '1900-01-01' THEN 'Final Fit'
                WHEN First_Fit != '1900-01-01' THEN 'First Fit'
                WHEN Cut != '1900-01-01' THEN 'Cut'
                WHEN MC_Rel != '1900-01-01' THEN 'MC Rel'
                WHEN Eng_Rel != '1900-01-01' THEN 'Eng Rel'
                WHEN Drawn != '1900-01-01' THEN 'Drawn'
                ELSE 'Unknown'
            END AS Spool_Status
        FROM dbo.tblSpools
        WHERE (
               MC_Rel != '1900-01-01 00:00:00.000' OR
               Cut != '1900-01-01 00:00:00.000' OR 
               First_Fit != '1900-01-01 00:00:00.000' OR 
               Final_Fit != '1900-01-01 00:00:00.000' OR 
               WeldOut != '1900-01-01 00:00:00.000' OR 
               Check_Out != '1900-01-01 00:00:00.000' OR 
               QC != '1900-01-01 00:00:00.000' OR 
               NDE != '1900-01-01 00:00:00.000' OR 
               InHydro != '1900-01-01 00:00:00.000' OR 
               QC_Complete != '1900-01-01 00:00:00.000' OR 
               RdyToShip != '1900-01-01 00:00:00.000' OR 
               OutsideService != '1900-01-01 00:00:00.000' OR 
               Ship_Paint != '1900-01-01 00:00:00.000' OR 
               Shipped != '1900-01-01 00:00:00.000'
        )
    )
    SELECT * FROM StatusCTE WHERE 1=1
        """


    # WHERE MC_Rel IS NOT NULL
    #       AND MC_Rel != '1900-01-01 00:00:00.000'

    # Add job number filtering if a list of active jobs is provided
    if active_jobs is not None:
        job_list = "', '".join(str(job) for job in active_jobs)
        query += f"\n          AND Job_ID IN ('{job_list}')"
    
    # Add job number filtering if provided
    if job_numbers is not None:
        # Convert job numbers list to properly formatted SQL string
        job_list = "', '".join(str(job) for job in job_numbers)
        query += f"\n          AND Job_ID IN ('{job_list}')"
    
    # Add filtering for all jobs after a certain number
    elif all_jobs_after is not None:
        # Convert input to string if it's not already
        base_job_number = str(all_jobs_after)
        
        # Extract only the numeric portion for comparison
        query += f"""
          AND CAST(
                CASE 
                    WHEN PATINDEX('%[0-9]%', Job_ID) > 0 
                    THEN SUBSTRING(Job_ID, 
                                 PATINDEX('%[0-9]%', Job_ID), 
                                 PATINDEX('%[^0-9]%', Job_ID + 'X') - PATINDEX('%[0-9]%', Job_ID))
                    ELSE '0'
                END AS INT) >= {base_job_number}"""
    
    try:
        with db_connection(db_name) as conn:
            df = pd.read_sql(query, conn)
            

            # Clean leading/trailing spaces 
            df = df.apply(clean_string_column)

            # print("Exporting Debug Spool Query to xlsx")
            # df.to_excel("Debug Spool Query.xlsx")

            # Normalize the values
            df.rename(columns=spool_column_mapping, inplace=True)

            print(df.columns)

            print(f"Spool Query Successful: {query}")

            return df

    except Exception as e:
        logger.error(f"Error executing spools query: {e}", exc_info=True)
        print(f"Error in Spool Query: \nQuery\n: {query}")
        return pd.DataFrame() # Return blank dataframe
        # print(f"Error executing query: {e}")
        # raise

def test_connection(conn):
    # Test the connection
    try:
        cursor = conn.cursor()
        cursor.execute("SELECT @@version")
        version = cursor.fetchone()
        print("Connection successful!")
        print(f"SQL Server version: {version[0]}")
    except pyodbc.Error as e:
        print(f"Connection failed: {str(e)}")
    finally:
        if 'conn' in locals():
            conn.close()

def inspect_permissions(table_name=None):
    if not table_name:
        print("\n\nMust Specify a table name!")
        return
        
    conn = db_connection(table_name)
    cursor = conn.cursor()
    permissions = {}
    
    try:
        # Current user and context
        permissions['user_info'] = {
            'user': cursor.execute("SELECT USER_NAME()").fetchval(),
            'user_type': cursor.execute("SELECT USER_ID()").fetchval(),
            'session_user': cursor.execute("SELECT SYSTEM_USER").fetchval(),
            'current_database': cursor.execute("SELECT DB_NAME()").fetchval()
        }
        
        # Database roles
        roles_query = """
        SELECT 
            DP1.name AS DatabaseRoleName,
            ISNULL (DP2.name, 'No members') AS DatabaseUserName   
        FROM sys.database_role_members AS DRM  
        RIGHT OUTER JOIN sys.database_principals AS DP1  
            ON DRM.role_principal_id = DP1.principal_id  
        LEFT OUTER JOIN sys.database_principals AS DP2  
            ON DRM.member_principal_id = DP2.principal_id  
        WHERE DP1.type = 'R'
        ORDER BY DP1.name;
        """
        permissions['database_roles'] = [dict(zip(['role', 'user'], row)) 
                                       for row in cursor.execute(roles_query)]
        
        # Table info
        table_info_query = """
        SELECT 
            TABLE_CATALOG,
            TABLE_SCHEMA,
            TABLE_NAME,
            TABLE_TYPE
        FROM INFORMATION_SCHEMA.TABLES 
        WHERE TABLE_NAME = ?
        """
        table_info = cursor.execute(table_info_query, table_name).fetchone()
        if table_info:
            permissions['table_info'] = dict(zip(
                ['catalog', 'schema', 'name', 'type'], table_info))
            schema = table_info[1]  # Get schema from table info
            
            # Add permission verification with schema
            has_select = cursor.execute("""
                SELECT HAS_PERMS_BY_NAME(?, 'OBJECT', 'SELECT') as has_select_permission;
            """, f"{schema}.{table_name}").fetchval()
            
            try:
                nolock_count = cursor.execute(f"""
                    SELECT COUNT_BIG(*) FROM {schema}.{table_name} WITH (NOLOCK);
                """).fetchval()
                
                normal_count = cursor.execute(f"""
                    SELECT COUNT_BIG(*) FROM {schema}.{table_name};
                """).fetchval()
                
                permissions['access_info'] = {
                    'has_select_permission': bool(has_select),
                    'row_count_nolock': nolock_count,
                    'row_count_normal': normal_count
                }
            except Exception as count_error:
                print(f"Error getting row counts: {str(count_error)}")
            
            # Column permissions
            column_query = """
            SELECT 
                COLUMN_NAME,
                DATA_TYPE,
                CHARACTER_MAXIMUM_LENGTH,
                IS_NULLABLE
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_NAME = ? AND TABLE_SCHEMA = ?
            """
            permissions['column_info'] = [dict(zip(
                ['name', 'type', 'max_length', 'nullable'], row))
                for row in cursor.execute(column_query, (table_name, schema))]
                
            # Effective permissions
            effective_query = """
            SELECT DISTINCT permission_name
            FROM fn_my_permissions(?, 'OBJECT')
            """
            permissions['effective_permissions'] = [row[0] 
                for row in cursor.execute(effective_query, f"{schema}.{table_name}")]
            
    except Exception as e:
        print(f"Error checking permissions: {str(e)}")
    finally:
        conn.close()
    
    return permissions

def print_permissions(perms):
    print("\nDetailed Permissions Report:")
    print("-" * 50)
    
    print("\nUser Information:")
    for key, value in perms['user_info'].items():
        print(f"  {key}: {value}")
    
    print("\nDatabase Roles:")
    for role in perms['database_roles']:
        print(f"  Role: {role['role']}, User: {role['user']}")
    
    print("\nEffective Permissions:")
    for perm in perms['effective_permissions']:
        print(f"  - {perm}")
    
    if 'table_info' in perms:
        print("\nTable Information:")
        for key, value in perms['table_info'].items():
            print(f"  {key}: {value}")

    print("\nAccess Information:")
    if 'access_info' in perms:
        for key, value in perms['access_info'].items():
            print(f"  {key}: {value}")
    
    print("\nColumn Information:")
    for col in perms['column_info']:
        print(f"  Column: {col['name']}")
        print(f"    Type: {col['type']}")
        print(f"    Max Length: {col['max_length']}")
        print(f"    Nullable: {col['nullable']}")



