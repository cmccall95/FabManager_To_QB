#from os import waitid_result
from multiprocessing.reduction import duplicate
import re, os, sys, requests, json
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QMenuBar, QMenu,
    QFileDialog, QVBoxLayout, QInputDialog, QTableWidget, QTableWidgetItem, 
    QPushButton, QMessageBox, QErrorMessage
    )
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtCore import Qt
import pandas as pd
from simpledbf import Dbf5
from dotenv import load_dotenv



load_dotenv()  # take environment variables from .env.
from qb_mod import get_table_data, export_weld_log_to_excel

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

# print("Temporary Export - (Check Function). Ensure to disable when done")
# export_weld_log_to_excel(WELD_LOG_ID)

def sanitize(df):
    return df.map(lambda x: ''.join(filter(lambda c: c.isprintable(), str(x))) if isinstance(x, str) else x)

def transform_linde_specs(original_df, needed_df, job_number):
    # Create a copy of the needed_df to avoid modifying the original DataFrame
    transformed_df = needed_df.copy()
    
    # Loop through each row in the original DataFrame
    for idx, row in original_df.iterrows():
        # Find the corresponding row in the needed DataFrame
        needed_idx = transformed_df[transformed_df['Option'] == row['LINE ID']].index
        
        # Check if a matching row is found
        if not needed_idx.empty:
            # Extract the first index value
            first_needed_idx = needed_idx[0]
            
            # Loop through the inspection types ('RT/PAUT', 'PT/MT')
            for inspection_type in ['RT/PAUT', 'PT/MT']:
                # Construct the column name in the needed format
                col_name = f"{inspection_type}|{row['JOINT TYPE']}"
                
                # Check if the column exists in the needed DataFrame
                if col_name in transformed_df.columns:
                    # Transfer the value
                    transformed_df.at[first_needed_idx, col_name] = row[inspection_type]

    # Replace NaN and blank values with 0
    #transformed_df = transformed_df.fillna(0)
    #transformed_df = transformed_df.replace('', 0)
    transformed_df = transformed_df.fillna('')

    transformed_df['Job'] = job_number
                
    return transformed_df

# Function to apply the regex to the 'Ref Drawing' column
def extract_series(ref_drawing):
    # First pattern: Matches the string within parentheses
    first_pattern = r'\([^\)]*\)([^\s\.]+)'
    # Second pattern: Matches the string without parentheses and optional following period/dash with numbers
    second_pattern = r'(?:\(|\s)(4-\d{2}[A-Z]+\d*)(?:[.-]\d+)?'
    
    # Try the first pattern
    match = re.search(first_pattern, ref_drawing)
    if match:
        return match.group(1)
    
    # If the first pattern didn't match, try the second one
    match = re.search(second_pattern, ref_drawing)
    return match.group(1) if match else ''

#Fill blanks in pipe spec
def fill_blank_pipe_specs(df):
    # Iterate through each row in the dataframe
    for index, row in df.iterrows():
        # Check if 'Pipe Spec' is blank
        if pd.isna(row['Pipe Spec']) or row['Pipe Spec'] == '':
            ref_drawing = row['Series']
            # Find a row with the same 'Ref Drawing' and a non-blank 'Pipe Spec'
            match = df[(df['Series'] == ref_drawing) & (df['Pipe Spec'].notna()) & (df['Pipe Spec'] != '')]
            if not match.empty:
                # If a match is found, use its 'Pipe Spec' to fill the blank
                df.at[index, 'Pipe Spec'] = match.iloc[0]['Pipe Spec']
    return df

def format_spools_df(df, weld_df, job_number): #Linde Setup
    
    df = df[['CONTROLNO', 'SP4', 'SP5', 'SP7', 'SP36', 'SP38' ]].copy()
    df.rename(columns={'CONTROLNO': 'Spool', 'SP4': 'MAT. Group', 'SP5': 'VT', 'SP7': 
                       'AREA', 'SP36': 'Ref Drawing', 'SP38': 'Spool/Piece Mark' }, inplace=True)

    # Remove the first row
    df = df.iloc[1:]

    # Remove leading zeroes from 'CONTROLNO'
    #df['Spool'] = df['Spool'].astype(str).str.lstrip('0').astype(int)
    df['Spool'] = df['Spool'].astype(str).apply(lambda x: re.sub('[^\d]', '', x)).astype(int)

    # Extract 'Series' and 'Sheet' from 'Ref Drawing'
    #df['Series'] = df['Ref Drawing'].apply(lambda x: re.search(r'\)(.*?)\.', x).group(1) if re.search(r'\)(.*?)\.', x) else '')
    #df['Series'] = df['Ref Drawing'].apply(lambda x: re.search(r'\s(\w+-\w+)', x).group(1) if re.search(r'\s(\w+-\w+)', x) else '')
    df['Series'] = df['Ref Drawing'].apply(extract_series)


    df['Sheet'] = df['Ref Drawing'].apply(lambda x: int(re.search(r'\.(\d+)$', x).group(1)) if re.search(r'\.(\d+)$', x) else '')

    # Drop duplicates in weld_df, keeping the row with non-blank 'Pipe Spec' if present
    weld_df = weld_df.sort_values('Pipe Spec', na_position='last').drop_duplicates(subset='Spool', keep='first')

    # Vlookup 'Pipe Spec' from weld_df
    df['Pipe Spec'] = df['Spool'].map(weld_df.set_index('Spool')['Pipe Spec'])

    # Set 'NDE Group' column equal to 'Series' Removed for temporary update
    df['NDE Group'] = df['Series']
    df['Job'] = job_number

    
    #print("\n1. ", df)

    return df

def format_spools_df2(df, weld_df, job_number): #OCI Setup JOBS: 30489
    
    df = df[['CONTROLNO', 'SP4', 'SP5', 'SP7', 'SP36', 'SP38' ]].copy()
    df.rename(columns={'CONTROLNO': 'Spool', 'SP4': 'MAT. Group', 'SP5': 'VT', 'SP7': 
                       'AREA', 'SP36': 'Ref Drawing', 'SP38': 'Spool/Piece Mark' }, inplace=True)

    # Remove the first row
    df = df.iloc[1:]

    # Remove leading zeroes from 'CONTROLNO'
    df['Spool'] = df['Spool'].astype(str).str.lstrip('0').astype(int)

    # Ensure 'Ref Drawing' is of string type
    df['Ref Drawing'] = df['Ref Drawing'].astype(str)

    # Extract 'Series' and 'Sheet' from 'Ref Drawing' based on the logic provided
    df['Series'] = df['Ref Drawing'].apply(lambda x: x.rsplit('-', 1)[0] if '-' in x else '')
    df['Sheet'] = df['Ref Drawing'].apply(lambda x: x.rsplit('-', 1)[1] if '-' in x else '')


    # Drop duplicates in weld_df, keeping the row with non-blank 'Pipe Spec' if present
    weld_df = weld_df.sort_values('Pipe Spec', na_position='last').drop_duplicates(subset='Spool', keep='first')

    # Vlookup 'Pipe Spec' from weld_df
    df['Pipe Spec'] = df['Spool'].map(weld_df.set_index('Spool')['Pipe Spec'])

    # Set 'NDE Group' column equal to 'Series'
    #df['NDE Group'] = '' #df['Series']

    df['Job'] = job_number

    #print("\n1. ", df)

    return df

def format_spools_df3(df, weld_df, job_number): #Blue Tide Jobs: 30496
    df = df[['CONTROLNO', 'SP4', 'SP5', 'SP7', 'SP36', 'SP38']].copy()
    df.rename(columns={'CONTROLNO': 'Spool', 'SP4': 'MAT. Group', 'SP5': 'VT', 'SP7': 
                       'AREA', 'SP36': 'Ref Drawing', 'SP38': 'Spool/Piece Mark'}, inplace=True)

    # Remove the first row
    df = df.iloc[1:]

    # Remove leading zeroes from 'Spool'
    df['Spool'] = df['Spool'].astype(str).str.lstrip('0').astype(int)

    # Ensure 'Ref Drawing' is of string type
    df['Ref Drawing'] = df['Ref Drawing'].astype(str)

    # Extract 'Series' and 'Sheet' from 'Ref Drawing'
    # The part before the last '-' is 'Series' and after it is 'Sheet'
    df['Series'] = df['Ref Drawing'].apply(lambda x: x.rsplit('-', 1)[0] if '-' in x else '')
    df['Sheet'] = df['Ref Drawing'].apply(lambda x: x.rsplit('-', 1)[1] if '-' in x else '')

    # Existing logic for processing weld_df
    weld_df = weld_df.sort_values('Pipe Spec', na_position='last').drop_duplicates(subset='Spool', keep='first')

    # Vlookup 'Pipe Spec' from weld_df
    df['Pipe Spec'] = df['Spool'].map(weld_df.set_index('Spool')['Pipe Spec'])
    
    # Set 'NDE Group' column equal to 'A'
    df['NDE Group'] = 'A'

    df['Job'] = job_number

    return df

def format_welds_df(df, job_number): #Linde Setup: Jobs: 30497, 81213
    # Transformations on welds_df
    df = df.copy()

    # Remove leading zeroes from 'CONTROLNO'
    df['CONTROLNO'] = df['CONTROLNO'].astype(str).str.lstrip('0').astype(int)
    
    # Create 'Weld_ID' column
    df['Weld ID'] = df['CONTROLNO'].astype(str) + '-' + df['WELDLABEL'].astype(str)
    
    # Remove '"' from 'SIZE_I' and ensure no leading or trailing zeroes
    df['SIZE_I'] = df['SIZE_I'].str.replace('"', '').astype(float).astype(str).str.strip('0').str.rstrip('.')
    
    # Replace 'BR' with 'BRANCH' in 'GENRE'
    df['GENRE'] = df['GENRE'].replace('BR', 'BRANCH')
    
    # Add new columns 'qb_id' and 'JOB' with empty values
    df['Related Record'] = ''
    df['Job'] = job_number

    # Record initial row count
    initial_row_count = df.shape[0]

    # Function to check if a value is NaN, blank or a variation of "nan"
    def is_invalid(value):
        value_str = str(value).strip().lower()
        return pd.isna(value) or value_str == '' or value_str == 'nan'

    # Drop rows where 'CONTROLNO' or 'WELDLABEL' is invalid
    df = df[~df['CONTROLNO'].apply(is_invalid) & ~df['WELDLABEL'].apply(is_invalid)]

    # Calculate the number of rows dropped
    final_row_count = df.shape[0]
    rows_dropped = initial_row_count - final_row_count
    print(f"{rows_dropped} rows were dropped due to invalid values in 'CONTROLNO' or 'WELDLABEL'")

    # Select only the desired columns and rename them
    df = df[['Related Record', 'Job', 'Weld ID', 'SPEC', 'SIZE_I', 'GENRE', 'CONTROLNO', 'WELDLABEL']]
    df.rename(columns={'SIZE_I': 'Size', 'GENRE': 'Joint', 'SPEC': 'Pipe Spec', 'CONTROLNO': 'Spool'}, inplace=True)

    return df

class CustomTableWidget(QTableWidget):
    def keyPressEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_C:
            self.copy_to_clipboard()
        elif event.modifiers() == Qt.KeyboardModifier.ControlModifier and event.key() == Qt.Key.Key_V:
            self.paste_clipboard()
        elif event.key() == Qt.Key.Key_Delete:
            self.clear_selected_cells()
        elif event.key() == Qt.Key.Key_Enter or event.key() == Qt.Key.Key_Return:
            self.move_to_next_cell()
        else:
            super().keyPressEvent(event)
            
    def move_to_next_cell(self):
        current_row = self.currentRow()
        current_column = self.currentColumn()
        next_row = current_row + 1

        if next_row < self.rowCount():
            self.setCurrentCell(next_row, current_column)
        else:
            # Optionally add a new row at the end if you've reached the last row
            self.insertRow(self.rowCount())
            self.setCurrentCell(self.rowCount(), current_column)
            
    def clear_selected_cells(self):
        selected_ranges = self.selectedRanges()
        for selected_range in selected_ranges:
            for r in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                for c in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    self.setItem(r, c, QTableWidgetItem(""))


    def copy_to_clipboard(self):
        selected_range = self.selectedRanges()[0]
        if selected_range:
            clipboard_content = ''
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                row_data = []
                for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    item = self.item(row, col)
                    row_data.append(item.text() if item else '')
                clipboard_content += '\t'.join(row_data) + '\n'
            QApplication.clipboard().setText(clipboard_content)

    def paste_clipboard(self):
        clipboard = QApplication.clipboard()
        clipboard_content = clipboard.text()

        # Split clipboard content into rows and cells
        clipboard_rows = clipboard_content.split('\n')
        clipboard_data = [row.split('\t') for row in clipboard_rows if row]

        # Get the first selected range
        selected_ranges = self.selectedRanges()
        if not selected_ranges:
            return  # No selection, so nothing to paste into

        selected_range = selected_ranges[0]

        for r in range(selected_range.rowCount()):
            for c in range(selected_range.columnCount()):
                clipboard_row = clipboard_data[r % len(clipboard_data)]
                clipboard_cell = clipboard_row[min(c, len(clipboard_row) - 1)]
                row_index = selected_range.topRow() + r
                col_index = selected_range.leftColumn() + c
                if row_index < self.rowCount() and col_index < self.columnCount():
                    self.setItem(row_index, col_index, QTableWidgetItem(clipboard_cell))

                        
class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle('My App')
        #self.showMaximized()
        
        self.tabWidget = QTabWidget()
        self.setCentralWidget(self.tabWidget)
        
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tab3 = QWidget()
        
        self.tabWidget.addTab(self.tab1, "Spools")
        self.tabWidget.addTab(self.tab2, "Welds")
        self.tabWidget.addTab(self.tab3, "Client Specs")
        
        self.menubar = QMenuBar(self)
        self.setMenuBar(self.menubar)
        
        self.file_menu = QMenu("File", self.menubar)
        self.menubar.addMenu(self.file_menu)
        
        self.import_action = QAction("Import .dbf Files", self)
        self.import_specs = QAction("Import Linde Specs", self)
        self.import_action.triggered.connect(self.import_dbf_welds)
        self.import_specs.triggered.connect(self.import_linde_specs)
        self.file_menu.addAction(self.import_action)
        self.file_menu.addAction(self.import_specs)

        #Using CustomTableWidget Class
        self.table1 = CustomTableWidget(self.tab1)
        self.table2 = CustomTableWidget(self.tab2)
        self.table3 = CustomTableWidget(self.tab3)
        # Create a QVBoxLayout for each tab
        self.layout1 = QVBoxLayout(self.tab1)
        self.layout2 = QVBoxLayout(self.tab2)
        self.layout3 = QVBoxLayout(self.tab3)
        
        # Create "Push to QuickBase" button for each tab
        self.push_to_qb_button1 = QPushButton("Push Spools to QuickBase", self.tab1)
        self.push_to_qb_button2 = QPushButton("Push Welds to QuickBase", self.tab2)
        self.push_to_qb_button3 = QPushButton("Push NDE Specs to QuickBase", self.tab3)

        #Connect Buttons
        self.push_to_qb_button1.clicked.connect(self.push_to_tpsl)
        self.push_to_qb_button2.clicked.connect(self.push_to_weld_log)
        self.push_to_qb_button3.clicked.connect(self.push_to_nde)

        # Add QTableWidget and button to the layout for each tab
        self.layout1.addWidget(self.push_to_qb_button1, stretch=0)  # No extra space to button
        self.layout1.addWidget(self.table1, stretch=1)             # All extra space to table
        self.layout2.addWidget(self.push_to_qb_button2, stretch=0)
        self.layout2.addWidget(self.table2, stretch=1)
        self.layout3.addWidget(self.push_to_qb_button3, stretch=0)
        self.layout3.addWidget(self.table3, stretch=1)

        # Set the layout for each tab
        self.tab1.setLayout(self.layout1)
        self.tab2.setLayout(self.layout2)
        self.tab3.setLayout(self.layout3)

        self.current_job_number = None  # Initialize the job number holder

    def set_job_number(self):
        while True:
            job_number, ok = QInputDialog.getText(self, "Input", "Enter Job Number:")
            if ok and job_number:
                # Ensure job number ends with '-'
                job_number = job_number if job_number.endswith('-') else job_number + '-'
                self.current_job_number = job_number
                return job_number
            else:
                response = QMessageBox.question(self, "No Job Number", "You didn't enter a job number. Do you want to try again?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if response == QMessageBox.StandardButton.No:
                    return None

    def get_job_number(self):
        return self.current_job_number  # Retrieve the stored job number

    def push_to_qb(self, payload):
        print("\n\n>>>Pushing to QB")
        headers = {
        "QB-Realm-Hostname": HOSTNAME,
        "User-Agent": USER_AGENT,
        "Authorization": TOKEN,
        'Content-Type': 'application/json'
        }

        response = requests.post("https://api.quickbase.com/v1/records", json=payload, headers=headers)

        # Check and print the status code
        status_code = response.status_code
        print("Status Code:", status_code)
    
        # Print a custom message based on the status code
        if status_code == 200:
            print("Status 200: Records Pushed Successfully")
        else:
            print("Error: Failed to push records")

        print("\n\nRESPONSE:", response.json())

        return response

    def create_df_from_table(self, table_name):
        # Select the appropriate table based on the table_name
        if table_name == 'tpsl':
            table_widget = self.table1

        elif table_name == 'weld log':
            table_widget = self.table2

        elif table_name == 'nde':
            table_widget = self.table3
        # Add more conditions here for other tables
        else:
            raise ValueError(f"Unknown table name: {table_name}")

        # Get the number of rows and columns from the table
        rows = table_widget.rowCount()
        cols = table_widget.columnCount()

        # Extract the header names
        headers = [table_widget.horizontalHeaderItem(c).text() for c in range(cols)]

        # Extract the data from the table
        data = []
        for r in range(rows):
            row_data = []
            for c in range(cols):
                item = table_widget.item(r, c)
                if item and item.text():
                    row_data.append(item.text())
                else:
                    row_data.append('')
            data.append(row_data)

        # Create a DataFrame from the extracted data
        df = pd.DataFrame(data, columns=headers)
        
        return df

   
    def prepare_payload(self, df, mapping, qb_table, existing_df=None, compare_columns=None):
        # Convert the df columns to field IDs using mapping
        df_mapped = df.rename(columns=mapping)

        payload_data = []

        # Check if existing_df is None or empty
        print("\n\nEXISTING DF: \n", existing_df)
        if existing_df is not None and not existing_df.empty:
            # Create a unique identifier for each row in df_mapped and existing_df based on the compare_columns
            df_mapped['unique_id'] = df_mapped[compare_columns].astype(str).agg('-'.join, axis=1)
            existing_df['unique_id'] = existing_df[compare_columns].astype(str).agg('-'.join, axis=1)

            # Make a set of the unique_ids in existing_df for efficient lookup
            existing_unique_ids = set(existing_df['unique_id'].values)

            duplicate_map = []

            # Iterate over df_mapped and prepare the payload data and duplicate_map
            for index, row in df_mapped.iterrows():
                if row['unique_id'] in existing_unique_ids:
                    # If the row is in existing_df, add to duplicate_map
                    existing_row = existing_df[existing_df['unique_id'] == row['unique_id']].iloc[0]
                    duplicate_map.append(f"{existing_row[compare_columns[1]]}:{existing_row['3']}")
                else:
                    # If the row is not in existing_df, add to payload
                    record = {str(field_id): {"value": row[field_id]} for field_id in mapping.values() if field_id in row}
                    payload_data.append(record)
        else:
            # If existing_df is None or empty, add all rows to payload_data
            duplicate_map = []
            for index, row in df_mapped.iterrows():
                record = {str(field_id): {"value": row[field_id]} for field_id in mapping.values() if field_id in row}
                payload_data.append(record)

        # The final payload
        payload = {
            "to": qb_table,
            "data": payload_data,
            "fieldsToReturn": [3] + [int(fid) for fid in mapping.values()]
        }

        # Logging the number of rows processed and ignored
        print(f"Original row count: {len(df_mapped)}, Payload row count: {len(payload_data)}, Duplicates ignored: {len(duplicate_map)}")
        print(f"Duplicate Map: {duplicate_map}")

        return payload, len(payload_data), duplicate_map


    def push_to_nde(self):
        #Create a dataframe from tab1
        df = self.create_df_from_table('nde')
        print("\n\nPUSH NDE DF:\n", "COLUMNS: \n", df.columns, "\n", df)

        # Reduce the dataframe to the first 5 rows
        #df = df.head(5)

        # Load the mapping from the JSON file
        with open('specs_map.json', 'r') as f:
            mapping = json.load(f)

        payload, row_count, excluded_records = self.prepare_payload(df, mapping, JOB_NDE)
        print("\n\nPayload Prepared - 'NDE - JOB LOT NDE'\n", payload)

        print("--SIMULATED")

        ## Push to QuickBase and capture the response
        #response = self.push_to_qb(payload)

        ##with pd.ExcelWriter('output.xlsx') as writer:
        ##    df.to_excel(writer, sheet_name='NDE Specs')
        
        ## Check if the response status code is 200
        #if response.status_code == 200:
        #    print("\n\nResponse Status: 200\nNDE % Pushed to Quickbase!.")
        #else:
        #    print("Bad Response:'n", response)

    def push_to_weld_log(self):
        #Create a dataframe from tab1
        df = self.create_df_from_table('weld log')
        print("\n\nPUSH DF: \n", df)

        # Reduce the dataframe to the first 5 rows
        #df = df.head(20)

        # Load the mapping from the JSON file
        with open('weld_map.json', 'r') as f:
            mapping = json.load(f)

        job_number = self.get_job_number()
        qb_fields=['3','6','15']
        existing_records_df = get_table_data(qb_fields, job_number, WELD_LOG_ID)

        if existing_records_df.shape[0] == 0:
            existing_records_df = None
        
        #altered_df = get_table_data()
        payload, row_count, excluded_records = self.prepare_payload(df, mapping, WELD_LOG_ID, existing_records_df, ['6', '15'])
        print("\n\nPayload Prepared - 'Welds - Weld Log'\n", payload)
        #print("--SIMULATED")

        #Abort if no rows
        if row_count < 1:
            error_dialog = QErrorMessage()
            error_dialog.showMessage('Payload is empty. \nNote: Duplicate records are automatically removed. \n0 Rows added to QuickBase.')
            error_dialog.exec()  # Use exec_() to make sure the dialog is modal and waits for user input
            return

         #Push to QuickBase and capture the response
        response = self.push_to_qb(payload)

        # Check if the response status code is 200
        if response.status_code == 200:
            QMessageBox.information(None, 'Success', f'Respond Status: 200 \n\n({row_count}) records added to QuickBase successfully.')
        
    def push_to_tpsl(self):
        #Create a dataframe from tab1
        df = self.create_df_from_table('tpsl')
        # Reduce the dataframe to the first 5 rows
        #df = df.head(10)
        print("\n\nPUSH DF:", df)

        # Load the mapping from the JSON file
        with open('spool_map.json', 'r') as f:
            mapping = json.load(f)

         #Check if duplicate records exists
        job_number = self.get_job_number()
        qb_fields=['3','6','108']
        existing_records_df = get_table_data(qb_fields, job_number, TPSL_ID)

        payload, row_count, duplicate_map = self.prepare_payload(df, mapping, TPSL_ID, existing_records_df, ['6', '108'])
        print("\n\nPayload Prepared - 'Spools - TPSL'\n", payload)

        #Abort if no rows
        if row_count < 1:
            if duplicate_map:
                # Convert duplicate_map list to a dictionary
                duplicate_map_dict = dict(item.split(":") for item in duplicate_map)

                # Add a new column 'TPSL Record' to the DataFrame
                df['TPSL Record'] = df['Spool'].map(duplicate_map_dict)

                # Reload the table1 with updated DataFrame
                self.load_data_into_table(self.table1, df)
                print("\n\n'Spools' table reloaded with updated existing records...\nExcluded Mapping: ", duplicate_map)

                # Update 'Related Record' in table2 (Welds)
                df_welds = self.create_df_from_table('weld log')
                df_welds['Related Record'] = df_welds['Spool'].map(duplicate_map_dict)

                # Reload the table2 (Welds) with updated DataFrame
                self.load_data_into_table(self.table2, df_welds)
                print("\n\n'Welds' table reloaded with updated existing spool records...")

            error_dialog = QErrorMessage()
            error_dialog.showMessage('Payload is empty. \nNote: Duplicate records are automatically removed. \n0 Rows added to QuickBase.')
            error_dialog.exec()  # Use exec_() to make sure the dialog is modal and waits for user input
            return

        # Push to QuickBase and capture the response
        response = self.push_to_qb(payload)
        
        # Check if the response status code is 200
        if response.status_code in [200, 207]:
            response_data = response.json()

            # Initialize a map to hold the mapping of spool numbers to record IDs
            spool_to_record_id = {}

                # Handle partial success (207)
            if response.status_code == 207:
                if 'lineErrors' in response_data:
                    line_errors = response_data['lineErrors']
                    for line, errors in line_errors.items():
                        # Log the errors along with the corresponding payload data
                        print(f"Errors on line {line}: {errors}")
                        print(f"Data on line {line}: {payload['data'][int(line)]}")

            QMessageBox.information(None, 'Success', f'Respond Status: {response.status_code} \n\n({row_count}) records added to QuickBase successfully.')

            
            if 'metadata' in response_data and 'createdRecordIds' in response_data['metadata']:
                print("\n\nResponse Status: 200\nRe-Rendering tables with Related Record IDs.")
                # Extract created Record IDs and corresponding Spool numbers
                created_record_ids = response_data['metadata']['createdRecordIds']
                spool_numbers = [record['108']['value'] for record in payload['data']]

                # Map Spool numbers to created Record IDs
                spool_to_record_id = dict(zip(spool_numbers, created_record_ids))
            else:
                # Handle cases where no record IDs are returned
                spool_to_record_id = {}

            # If duplicate_map has items, convert it to a dictionary and combine it with spool_to_record_id
            if duplicate_map:
                # Convert duplicate_map list to a dictionary
                duplicate_map_dict = dict(item.split(":") for item in duplicate_map)
                # Combine the two dictionaries, with duplicate_map_dict taking priority
                combined_record_map = {**spool_to_record_id, **duplicate_map_dict}
            else:
                combined_record_map = spool_to_record_id

            # Add a new column 'TPSL Record' to the DataFrame
            df['TPSL Record'] = df['Spool'].map(combined_record_map)

            # Reload the table1 with updated DataFrame
            self.load_data_into_table(self.table1, df)
            print("\n\n'Spools' table reloaded with updated records...\nExcluded Mapping: ", duplicate_map)

            # Update 'Related Record' in table2 (Welds)
            df_welds = self.create_df_from_table('weld log')
            df_welds['Related Record'] = df_welds['Spool'].map(combined_record_map)

            with pd.ExcelWriter('output.xlsx') as writer:
                df_welds.to_excel(writer, sheet_name='Welds')
                df.to_excel(writer, sheet_name='Spools')

            # Reload the table2 (Welds) with updated DataFrame
            self.load_data_into_table(self.table2, df_welds)
            print("\n\n'Welds' table reloaded with updated records...")

    def load_data_into_table(self, table_widget, df):
        print("/n/nLoading Data into Tables")
        # Set the number of rows and columns in the table
        table_widget.setRowCount(df.shape[0])
        table_widget.setColumnCount(df.shape[1])

        # Set the column headers
        table_widget.setHorizontalHeaderLabels(df.columns.tolist())

        # Populate the table with data
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                value = df.iloc[i, j]
        
                # Check if the value is numeric and not NaN
                if pd.notna(value) and isinstance(value, float):
                    # Convert float values to int if they are whole numbers
                    value = int(value) if value.is_integer() else value
                elif pd.notna(value) and isinstance(value, int):
                    # If value is already an integer, do nothing
                    pass
                elif pd.isna(value) or value == 'nan':
                    # Replace 'nan' with an empty string
                    value = ''
            
                table_widget.setItem(i, j, QTableWidgetItem(str(value)))


    def import_linde_specs(self):
        # Define the path to the workbook
        #job_number, ok = QInputDialog.getText(self, "Input", "Enter Job Number:")
        #if ok:
        #    # Ensure job number ends with '-'
        #    job_number = job_number if job_number.endswith('-') else job_number + '-'
        # If job number is not set, prompt the user to set it
        if not self.get_job_number():
            if self.set_job_number() is None:
                print("Operation cancelled by the user --Linde Specs.")
                return  # Exit the function if no job number is set

        job_number = self.get_job_number()

        workbook_path = 'Linde NDE.xlsx'

        # Load data from the 'Original Format' sheet
        original_df = pd.read_excel(workbook_path, sheet_name='Original Format')

        # Load data from the 'Needed Format' sheet
        needed_df = pd.read_excel(workbook_path, sheet_name='Reformat')

        # Apply the transformation
        transformed_df = transform_linde_specs(original_df, needed_df, job_number)

        self.load_data_into_table(self.table3, transformed_df)
        print("\n\nLoaded 'Client Spec'...")

        # Write dataframes to Excel
        #with pd.ExcelWriter('output_spec.xlsx') as writer:
        #    transformed_df.to_excel(writer, sheet_name='Spec')

    def import_dbf_welds(self):
        if not self.get_job_number():
            if self.set_job_number() is None:
                print("Operation cancelled by the user.")
                return  # Exit the function if no job number is set

        job_number = self.get_job_number()

        options = QFileDialog.Option.ReadOnly
        file, _ = QFileDialog.getOpenFileName(self, "Import 'Welds' .dbf File", "", "Database Files (*.dbf);;All Files (*)", options=options)
        if file:
            print(file)
            welds_df = self.read_dbf(file)
            self.import_dbf_spools(welds_df, job_number)

    def import_dbf_spools(self, welds_df, job_number):
        options = QFileDialog.Option.ReadOnly
        file, _ = QFileDialog.getOpenFileName(self, "Import 'Spools' .dbf File", "", "Database Files (*.dbf);;All Files (*)", options=options)
        if file:
            print(file)
            spools_df = self.read_dbf(file)
            self.export_to_excel(welds_df, spools_df, job_number)

    def read_dbf(self, file_path):
        dbf = Dbf5(file_path)
        return dbf.to_dataframe()

    def export_to_excel(self, welds_df, spools_df, job_number):
        # Sanitize data
        welds_df = sanitize(welds_df)
        spools_df = sanitize(spools_df)

        formatted_welds_df = format_welds_df(welds_df, job_number)
        if job_number == "30489-":
            formatted_spools_df = format_spools_df2(spools_df, formatted_welds_df, job_number) #OCI Format
            
        elif job_number == "30497-":
            formatted_spools_df = format_spools_df(spools_df, formatted_welds_df, job_number) #Linde Format
            
        elif job_number == "30496-":
            formatted_spools_df = format_spools_df3(spools_df, formatted_welds_df, job_number) #Blue Tide Format
            
        else:
            formatted_spools_df = format_spools_df(spools_df, formatted_welds_df, job_number) #Linde Format
            QMessageBox.information(None, 'Success', f'Not configured for job: {job_number}\nDrawing formats will be assumed as format 1(format_spools_df).')
            #return
        
        print("\n\n>>> Attempting to fill blanks in 'Pipe Spec'....")
        filled_df = fill_blank_pipe_specs(formatted_spools_df)
        # Load data into tables
        self.load_data_into_table(self.table1, formatted_spools_df)
        self.load_data_into_table(self.table2, formatted_welds_df)
        print("\n\nLoaded 'Spools' - (TPSL)...\nLoaded 'Welds' - (Weld Log)...")

        # Write dataframes to Excel
        #with pd.ExcelWriter('output.xlsx') as writer:
        #    formatted_welds_df.to_excel(writer, sheet_name='Welds')
        #    formatted_spools_df.to_excel(writer, sheet_name='Spools')
 

if __name__ == "__main__":
    app = QApplication([])
    window = MyWindow()
    window.showMaximized()
    app.exec()
