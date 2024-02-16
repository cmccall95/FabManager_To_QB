#from os import waitid_result
from multiprocessing.reduction import duplicate
import re, os, sys, requests, json
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QWidget, QMenuBar, QMenu,
    QFileDialog, QVBoxLayout, QInputDialog, QTableWidget, QTableWidgetItem, 
    QPushButton, QMessageBox, QErrorMessage, QHeaderView, QTableView, QWidgetAction, QListWidget,
   QLineEdit, QListWidgetItem, QLabel, QCheckBox)
from PyQt6.QtGui import QAction, QKeySequence, QStandardItemModel, QStandardItem, QIcon, QGuiApplication, QColor, QFont
from PyQt6.QtCore import Qt, pyqtSignal, QRect, QPoint, QTimer
import pandas as pd
from simpledbf import Dbf5
from openpyxl import load_workbook
#from openpyxl.utils.exceptions import FileInUseError
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

# Function to apply the regex to the 'Ref Drawing' column #Linde: 30497
def extract_series(ref_drawing):
    # Dictionary of patterns with sample strings as keys and tuples as values (pattern, match group)
    regex_patterns = {
        "sample1 (Series in parentheses)": (r'\([^\)]*\)([^\s\.]+)', 1),
        "sample2 (4-XXALPHANUM)": (r'(?:\(|\s)(4-\d{2}[A-Z]+\d*-[A-Z]*\d*)(?=\.)', 1), #sample2 (4-XXALPHANUM)": (r'(?:\(|\s)(4-\d{2}[A-Z]+\d*)(?:[.-]\d+)?', 1),
        "sample3 (After 'ZL' before '.')": (r'ZL([^-]+-\d+[^.]*)', 1),
        "sample4 (After 'ZL6-' or 'ZL5-')": (r'(?<=ZL6-)\d+[A-Z]+\d+[A-Z]*\b|(?<=ZL5-)\d+[A-Z]+\d+[A-Z]*\b', 0),
    }

    # Loop through the dictionary and apply each regex
    for sample_string, (pattern, m_group) in regex_patterns.items():
        match = re.search(pattern, ref_drawing)
        if match:
            # Extracted series
            extracted_series = match.group(m_group).strip()
            # Clean up the extracted series by removing parenthetical content
            cleaned_series = re.sub(r'\(.*?\)', '', extracted_series).strip()
            # Return the cleaned series and the pattern that matched
            return cleaned_series, pattern

    # If no pattern matched, return None
    return None, None

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

def format_spools_df(df, weld_df, job_number): #Linde Setup: 30497
    
    df = df[['CONTROLNO', 'SP4', 'SP5', 'SP7', 'SP36', 'SP38' ]].copy()
    df.rename(columns={'CONTROLNO': 'Spool', 'SP4': 'MAT. Group', 'SP5': 'VT', 'SP7': 
                       'AREA', 'SP36': 'Ref Drawing', 'SP38': 'Spool/Piece Mark' }, inplace=True)

    # Remove the first row
    df = df.iloc[1:]

    # Remove leading zeroes from 'CONTROLNO'
    #df['Spool'] = df['Spool'].astype(str).str.lstrip('0').astype(int)
    df['Spool'] = df['Spool'].astype(str).apply(lambda x: re.sub('[^\d]', '', x)).astype(int)

    # Extract 'Series' and 'Sheet' from 'Ref Drawing'
    df['Series'], df['re Match Pattern'] = zip(*df['Ref Drawing'].apply(extract_series))
    
    # Extract 'Sheet' by first checking for digits after a period, then checking for a 3-character string starting with '0' between hyphens
    def extract_sheet(ref_drawing):
        # First, try to match three digits after a period
        match = re.search(r'\.(\d{3})\b', ref_drawing)
        if match:
            return int(match.group(1))
        # If that fails, try to match a 3-character string starting with '0' between hyphens
        match = re.search(r'-(0\d{2})-', ref_drawing)
        return int(match.group(1)) if match else None

    df['Sheet'] = df['Ref Drawing'].apply(extract_sheet)
    
    #df['Sheet'] = df['Ref Drawing'].apply(lambda x: int(re.search(r'\.(\d{3})\b', x).group(1)) if re.search(r'\.(\d{3})\b', x) else None)

    #df['Sheet'] = df['Ref Drawing'].apply(lambda x: int(re.search(r'\.(\d+)$', x).group(1)) if re.search(r'\.(\d+)$', x) else '')

    # Drop duplicates in weld_df, keeping the row with non-blank 'Pipe Spec' if present
    weld_df = weld_df.sort_values('Pipe Spec', na_position='last').drop_duplicates(subset='Spool', keep='first')

    # Vlookup 'Pipe Spec' from weld_df
    df['Pipe Spec'] = df['Spool'].map(weld_df.set_index('Spool')['Pipe Spec'])

    # Set 'NDE Group' column equal to 'Series' Removed for temporary update
    df['NDE Group'] = df['Series']
    df['Job'] = job_number
    #print("\n\nFormat 1 df:\n", df)
    return df

def format_spools_df2(df, weld_df, job_number): #OCI Setup JOBS: 30489
    
    df = df[['CONTROLNO', 'SP4', 'SP5', 'SP7', 'SP36', 'SP38' ]].copy()
    df.rename(columns={'CONTROLNO': 'Spool', 'SP4': 'MAT. Group', 'SP5': 'VT', 'SP7': 
                       'AREA', 'SP36': 'Ref Drawing', 'SP38': 'Spool/Piece Mark' }, inplace=True)

    # Initialize a list to store non-convertible values
    non_convertible_values = []
    
    # Function to try conversion and catch errors
    def try_convert_to_int(value):
        try:
            return int(float(value.lstrip('0')))
        except ValueError:
            non_convertible_values.append(value)
            return None  # Or a placeholder value indicating conversion failure

    # Apply conversion with error handling
    df['Spool'] = df['Spool'].astype(str).apply(try_convert_to_int)

    # Print the non-convertible values
    if non_convertible_values:
        print(f"({len(non_convertible_values)})Non-convertible values in 'Spool' column: {non_convertible_values}")

    # Drop rows with non-convertible values
    df = df.dropna(subset=['Spool'])

    # Remove the first row
    df = df.iloc[1:]

    # Remove leading zeroes from 'CONTROLNO'
    #df['Spool'] = df['Spool'].astype(str).str.lstrip('0').astype(int) Handled in function

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

# def format_welds_df(df, job_number): #Linde Setup: Jobs: 30497, 81213
#     # Transformations on welds_df
#     df = df.copy()

#     # Remove leading zeroes from 'CONTROLNO'
#     df['CONTROLNO'] = df['CONTROLNO'].astype(str).str.lstrip('0').astype(int)
    
#     # Create 'Weld_ID' column
#     df['Weld ID'] = df['CONTROLNO'].astype(str) + '-' + df['WELDLABEL'].astype(str)
    
#     # Remove '"' from 'SIZE_I' and ensure no leading or trailing zeroes
#     df['SIZE_I'] = df['SIZE_I'].str.replace('"', '').astype(float).astype(str).str.strip('0').str.rstrip('.')
    
#     # Replace 'BR' with 'BRANCH' in 'GENRE'
#     df['GENRE'] = df['GENRE'].replace('BR', 'BRANCH')
    
#     # Add new columns 'qb_id' and 'JOB' with empty values
#     df['Related Record'] = ''
#     df['Job'] = job_number

#     # Record initial row count
#     initial_row_count = df.shape[0]

#     # Function to check if a value is NaN, blank or a variation of "nan"
#     def is_invalid(value):
#         value_str = str(value).strip().lower()
#         return pd.isna(value) or value_str == '' or value_str == 'nan'

#     # Drop rows where 'CONTROLNO' or 'WELDLABEL' is invalid
#     df = df[~df['CONTROLNO'].apply(is_invalid) & ~df['WELDLABEL'].apply(is_invalid)]

#     # Calculate the number of rows dropped
#     final_row_count = df.shape[0]
#     rows_dropped = initial_row_count - final_row_count
#     print(f"{rows_dropped} rows were dropped due to invalid values in 'CONTROLNO' or 'WELDLABEL'")

#     # Select only the desired columns and rename them
#     df = df[['Related Record', 'Job', 'Weld ID', 'SPEC', 'SIZE_I', 'GENRE', 'CONTROLNO', 'WELDLABEL']]
#     df.rename(columns={'SIZE_I': 'Size', 'GENRE': 'Joint', 'SPEC': 'Pipe Spec', 'CONTROLNO': 'Spool'}, inplace=True)

#     return df

def format_welds_df(df, job_number):
    df = df.copy()

    # Function to check if CONTROLNO can be converted to an integer
    def is_convertible_to_int(value):
        try:
            int(value)
            return True
        except ValueError:
            return False

    # Apply the function to create a mask for rows with valid CONTROLNO
    valid_controlno_mask = df['CONTROLNO'].astype(str).str.lstrip('0').apply(is_convertible_to_int)

    # Filter the DataFrame based on the mask
    df = df[valid_controlno_mask]

    # Now it's safe to convert CONTROLNO to int
    df['CONTROLNO'] = df['CONTROLNO'].astype(str).str.lstrip('0').astype(int)
    
    # Create 'Weld ID' column
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

    # Drop rows where 'WELDLABEL' is invalid
    df = df[~df['WELDLABEL'].apply(is_invalid)]

    # Calculate the number of rows dropped
    final_row_count = df.shape[0]
    rows_dropped = initial_row_count - final_row_count
    print(f"{rows_dropped} rows were dropped due to invalid values in 'CONTROLNO' or 'WELDLABEL'")

    # Select only the desired columns and rename them
    df = df[['Related Record', 'Job', 'Weld ID', 'SPEC', 'SIZE_I', 'GENRE', 'CONTROLNO', 'WELDLABEL']]
    df.rename(columns={'SIZE_I': 'Size', 'GENRE': 'Joint', 'SPEC': 'Pipe Spec', 'CONTROLNO': 'Spool'}, inplace=True)

    return df

def pre_check(df, required_columns, integer_columns=[], hyphen_columns=[]):
    df['Error'] = ''
    df['Error Description'] = ''
    df['Comments'] = ''
    
    # Check for specific columns not being blank
    for column in required_columns:
        if column in df.columns:
            df.loc[df[column].isnull() | (df[column] == ''), 'Error'] = True
            df.loc[df[column].isnull() | (df[column] == ''), 'Error Description'] += f'{column} is blank; '

    # Check for columns that must end with a hyphen
    for column in hyphen_columns:
        if column in df.columns:
            df.loc[~df[column].astype(str).str.endswith('-'), 'Error'] = True
            df.loc[~df[column].astype(str).str.endswith('-'), 'Error Description'] += f'{column} does not end with "-"; '
            
    # Check and update 'Sheet' column if it contains 'VOID'
    if 'Sheet' in df.columns:
        #contains_void_mask = df['Sheet'].str.contains('VOID', case=False, na=False)
        contains_void_mask = df['Sheet'].astype(str).str.contains('VOID', case=False, na=False)
        df.loc[contains_void_mask, 'Comments'] = "This Spool has been marked as 'VOID'."
        #df.loc[contains_void_mask, 'Sheet'] = df.loc[contains_void_mask, 'Sheet'].str.replace('VOID', '', case=False).str.strip()
        df.loc[contains_void_mask, 'Sheet'] = df.loc[contains_void_mask, 'Sheet'].astype(str).str.replace('VOID', '', case=False).str.strip()

    # Check for columns that must be integers
    for column in integer_columns:
        if column in df.columns:
            df.loc[~df[column].fillna('0').astype(str).str.isdigit(), 'Error'] = True
            df.loc[~df[column].fillna('0').astype(str).str.isdigit(), 'Error Description'] += f'{column} is not an integer; '

    # Check for 'Nan' or 'nan' in any column
    for col in df.columns:
        # Check if the column type is string; if so, then perform the 'Nan' check
        if pd.api.types.is_string_dtype(df[col]):
            df.loc[df[col].str.lower() == 'nan', 'Error'] = True
            df.loc[df[col].str.lower() == 'nan', 'Error Description'] += f'{col} has Nan; '
        # For non-string columns, check if they are not null before converting to string
        else:
            df.loc[df[col].notnull() & (df[col].astype(str).str.lower() == 'nan'), 'Error'] = True
            df.loc[df[col].notnull() & (df[col].astype(str).str.lower() == 'nan'), 'Error Description'] += f'{col} has Nan; '

    return df

def post_check(df, required_columns=[], integer_columns=[], invalid_format_columns=[]):
    # Assuming 'Error' and 'Error Description' columns already exist
    # If not, uncomment the following lines
    # df['Error'] = ''
    # df['Error Description'] = ''

    # Check for required columns being blank
    for column in required_columns:
        if column in df.columns:
            df.loc[df[column].isnull() | (df[column] == ''), 'Error'] = True
            df.loc[df[column].isnull() | (df[column] == ''), 'Error Description'] += f'{column} is blank; '

    # Check for columns that must be integers
    for column in integer_columns:
        if column in df.columns:
            df.loc[~df[column].fillna('0').astype(str).str.isdigit(), 'Error'] = True
            df.loc[~df[column].fillna('0').astype(str).str.isdigit(), 'Error Description'] += f'{column} is not an integer; '

    # Check for invalid formats in specified columns
    for column in invalid_format_columns:
        if column in df.columns:
            invalid_format_condition = df[column].astype(str).str.contains('value:|{|}', regex=True)
            df.loc[invalid_format_condition, 'Error'] = True
            df.loc[invalid_format_condition, 'Error Description'] += f'{column} has invalid format; '

    return df

class CustomHeaderView(QHeaderView):
    iconClicked = pyqtSignal(int)  # Signal emitted when an icon is clicked
    filterRequested = pyqtSignal(int)  # Signal emitted when filtering is requested

    def __init__(self, orientation, parent=None):
        super().__init__(orientation, parent)
        print("CustomHeaderView accessed")
        self.icon = QIcon("assets/filter_icon.svg")
        
        if self.icon.isNull():
            print("Filter icon failed to load. Check the file path.")
            
    def setModel(self, model):
        super().setModel(model)
        for col in range(model.columnCount()):
            self.setIconForColumn(col)

    def setIconForColumn(self, col):
        icon = QIcon("assets/filter_icon.svg")
        self.model().setHeaderData(col, Qt.Orientation.Horizontal, icon, Qt.ItemDataRole.DecorationRole) #Actually makes the icon appear but is non adjustable

    def paintSection(self, painter, rect, logicalIndex):
        super().paintSection(painter, rect, logicalIndex)

        #print("PaintSection has been called")
        # Ensure icon is loaded
        if self.icon.isNull():
            print("Icon is not valid.")
            return

        # Calculate icon position
        icon_size = 20
        icon_x = rect.right() - icon_size - 5  # 5 pixels from the right edge
        icon_y = (rect.height() - icon_size) // 2  # Vertically centered
        icon_rect = QRect(icon_x, icon_y, icon_size, icon_size)

        # Paint the icon
        self.icon.paint(painter, icon_rect)
        
    def sectionCountChanged(self, oldCount, newCount):
        super().sectionCountChanged(oldCount, newCount)
        # Update icons for new columns
        for col in range(newCount):
            self.setIconForColumn(col)

    def mousePressEvent(self, event):
        super().mousePressEvent(event)

        # Check if the click event occurred on the icon area
        index = self.logicalIndexAt(event.position().toPoint())
        icon_rect = QRect(self.sectionViewportPosition(index) + 2, self.height() - 22, 20, 20)  # Bottom-left position
        
        if icon_rect.contains(event.position().toPoint()):
            self.iconClicked.emit(index)
            
    def contextMenuEvent(self, event):
        index = self.logicalIndexAt(event.pos())
        column_name = self.model().headerData(index, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
        menu = QMenu(self)
        filter_action = QAction(f"Filter Column: {column_name}", self)
        menu.addAction(filter_action)

        filter_action.triggered.connect(lambda: self.filterRequested.emit(index))
        
        menu.exec(event.globalPos())  # Display the context menu at the global position

class StandardTable(QWidget):
    def __init__(self, parent=None):  # Add parent=None to accept an optional parent argument
        super().__init__(parent)  # Pass the parent to the QWidget constructor
        self.resetFilters()  # Reset filters on initialization
        
        self.layout = QVBoxLayout(self)
        
        # Create a model and view for the standard table
        self.model = QStandardItemModel()
        self.view = QTableView(self)
        self.view.setModel(self.model)
        
        self.appliedFilters = {}  # Dictionary to store applied filters
        self.loadFilters()  # Load existing filters on initialization

        # Set the custom header view for the table
        header_view = CustomHeaderView(Qt.Orientation.Horizontal, self.view)
        self.view.setHorizontalHeader(header_view)
        
        header_view.iconClicked.connect(self.filter_click)
        header_view.filterRequested.connect(self.filter_click)
        
        # Set up the QTableView
        # self.view.setEditTriggers(QTableView.NoEditTriggers)  # Optional: make cells not editable
        # self.view.setSelectionBehavior(QTableView.SelectRows)  # Optional: change selection behavior

        self.layout.addWidget(self.view)
        
    def resetFilters(self):
        print("Filters Reset - StandardTableClass")
        self.appliedFilters = {}  # Clear the applied filters dictionary
        if os.path.exists("filters.json"):
            os.remove("filters.json")  # Remove the filters.json file if it exists
        
    def build_filter_menu(self, column_index):
        menu = QMenu(self)

        # Custom widget action for the list
        widgetAction = QWidgetAction(menu)
        containerWidget = QWidget()
        layout = QVBoxLayout(containerWidget)
        
        # Retrieve and display the column name in bold
        column_name = self.model.headerData(column_index, Qt.Orientation.Horizontal, Qt.ItemDataRole.DisplayRole)
        columnNameLabel = QLabel(column_name)
        boldFont = QFont()
        boldFont.setBold(True)
        columnNameLabel.setFont(boldFont)
        layout.addWidget(columnNameLabel)

        # Optional search field
        searchEdit = QLineEdit()
        searchEdit.setPlaceholderText("Search...")
        layout.addWidget(searchEdit)
        
        # "Select All" checkbox
        selectAllCheckbox = QCheckBox("Select All")
        selectAllCheckbox.setCheckState(Qt.CheckState.Checked)  # Set the initial state to Checked
        #selectAllCheckbox.setTristate(False)  # Enable tri-state for partial selections
        layout.addWidget(selectAllCheckbox)

        # List widget for the values
        listWidget = QListWidget()
        layout.addWidget(listWidget)
        widgetAction.setDefaultWidget(containerWidget)

        # Populate list with sorted unique values
        unique_values = set()
        for row in range(self.model.rowCount()):
            value = self.model.item(row, column_index).text()
            unique_values.add(value)

        sorted_values = sorted(unique_values, key=self.alphanumeric_key)
        for value in sorted_values:
            item = QListWidgetItem(value)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            listWidget.addItem(item)
            
        # Additional button for adding to filter
        addToFilterButton = QPushButton("Add to Filter")
        layout.addWidget(addToFilterButton)

        # Ok and Cancel buttons
        okButton = QPushButton("Apply")
        cancelButton = QPushButton("Cancel")
        layout.addWidget(okButton)
        layout.addWidget(cancelButton)
        
        # Connect the searchEdit signal
        searchEdit.textChanged.connect(lambda text: self.filterList(listWidget, text))
        
        # Set the initial state of selectAllCheckbox based on the items
        #self.updateSelectAllState(listWidget, selectAllCheckbox)
        
        # Connect the checkbox signal with a lambda function to pass the correct state
        #selectAllCheckbox.stateChanged.connect(lambda state=selectAllCheckbox.checkState(): self.selectAllItems(listWidget, state, False))
        
        # Inside build_filter_menu method
        selectAllCheckbox.stateChanged.connect(lambda state: self.onSelectAllStateChanged(listWidget, state))
        #selectAllCheckbox.stateChanged.connect(lambda state: self.onSelectAllStateChanged(listWidget, state, selectAllCheckbox))

        # Connect button signals
        okButton.clicked.connect(lambda: self.applyFilter(listWidget, column_index, False))  # Replace menu.close with applyFilter
        cancelButton.clicked.connect(menu.close)
        addToFilterButton.clicked.connect(lambda: self.applyFilter(listWidget, column_index, True))  # Keep the dialog open
        cancelButton.clicked.connect(menu.close)

        menu.addAction(widgetAction)
        menu.addSeparator()
        menu.addAction("Sort Smallest to Largest")
        menu.addAction("Sort Largest to Smallest")
        menu.addAction("Clear Filter")

        return menu
    
    def saveFilters(self):
        with open("filters.json", "w") as file:
            json.dump(self.appliedFilters, file)
            print("Filters Added")

    def loadFilters(self):
        try:
            with open("filters.json", "r") as file:
                self.appliedFilters = json.load(file)
        except (FileNotFoundError, json.JSONDecodeError):
            self.appliedFilters = {}
    
    def applyFilter(self, listWidget, column_index, keepOpen):
        selected_values = [listWidget.item(i).text() for i in range(listWidget.count()) if listWidget.item(i).checkState() == Qt.CheckState.Checked]
        self.appliedFilters[column_index] = selected_values
        self.saveFilters()

        if not keepOpen:
            self.sender().parent().close()  # Close the menu if not adding to filter

    def filterList(self, listWidget, text):
        print("\nTriggered Filter Search")
        for i in range(listWidget.count()):
            item = listWidget.item(i)
            item.setHidden(text.lower() not in item.text().lower())
            
    def alphanumeric_key(self, s):
        """
        Generate a key for sorting strings that may contain numbers (including floats).
        """
        def convert(text):
            try:
                return float(text)
            except ValueError:
                return text.lower()
    
        # Adjust the regular expression to capture floating-point numbers starting with a dot
        return [convert(c) for c in re.split('(\d+\.\d+|\.\d+|\d+)', s)]
    
    def updateSelectAllCheckboxState(self, listWidget, selectAllCheckbox):
        total_visible_items = sum(1 for i in range(listWidget.count()) if not listWidget.item(i).isHidden())
        checked_items = sum(listWidget.item(i).checkState() == Qt.CheckState.Checked for i in range(listWidget.count()) if not listWidget.item(i).isHidden())

        if checked_items == total_visible_items:
            selectAllCheckbox.setCheckState(Qt.CheckState.Checked)
        elif checked_items == 0:
            selectAllCheckbox.setCheckState(Qt.CheckState.Unchecked)
        else:
            selectAllCheckbox.setCheckState(Qt.CheckState.PartiallyChecked)


    def onSelectAllStateChanged(self, listWidget, state):
        print(f"\nonSelectAllStateChanged Triggered. State: {state}")
        
        selectAllCheckbox = self.sender()  # This will be the selectAllCheckbox that triggered the signal
        
        # Correctly determine new_state based on the received state value
        new_state = Qt.CheckState.Checked if state == 2 else Qt.CheckState.Unchecked
        print(f"New State: {new_state}")
        
        self.updateListWidgetItems(listWidget, new_state, selectAllCheckbox)

    def updateListWidgetItems(self, listWidget, new_state, selectAllCheckbox):
        print("Updating List Widget Items")
        listWidget.blockSignals(True)
        for i in range(listWidget.count()):
            item = listWidget.item(i)
            if not item.isHidden():
                print(f"Item {i} CheckState: {item.checkState()} -> {new_state}")
                item.setCheckState(new_state)
                # Connect item check state change to updateSelectAllCheckboxState
                item.setCheckState.connect(lambda: self.updateSelectAllCheckboxState(listWidget, selectAllCheckbox))
                
        listWidget.blockSignals(False)


    # def selectAllItems(self, listWidget, state, updateCheckboxState=True):
    #     print(f"\nSelectAllItems Triggered. State: {state}")
    #     listWidget.blockSignals(True)

    #     new_state = Qt.CheckState.Checked if state == Qt.CheckState.Checked else Qt.CheckState.Unchecked

    #     for i in range(listWidget.count()):
    #         item = listWidget.item(i)
    #         if not item.isHidden():
    #             item.setCheckState(new_state)

    #     listWidget.blockSignals(False)
    #     if updateCheckboxState:
    #         self.updateSelectAllState(listWidget, self.sender())

    # def updateSelectAllState(self, listWidget, selectAllCheckbox):
    #     visible_count = sum(1 for i in range(listWidget.count()) if not listWidget.item(i).isHidden())
    #     checked_count = sum(not listWidget.item(i).isHidden() and listWidget.item(i).checkState() == Qt.CheckState.Checked for i in range(listWidget.count()))

    #     print(f"\nupdateSelectAllState Triggered\nVisible: {visible_count}\nChecked Count: {checked_count}")

    #     if checked_count == 0:
    #         selectAllCheckbox.setCheckState(Qt.CheckState.Unchecked)
    #     elif checked_count == visible_count:
    #         selectAllCheckbox.setCheckState(Qt.CheckState.Checked)
    #     else:
    #         selectAllCheckbox.setCheckState(Qt.CheckState.Unchecked)

    def filter_click(self, column_index):
        header = self.view.horizontalHeader()
        menu = self.build_filter_menu(column_index)

        # Temporary show the menu off-screen to calculate its size
        temp_pos = QPoint(-10000, -10000)
        menu.move(temp_pos)
        menu.show()
        menu_height = menu.height()
        menu_width = menu.width()
        menu.hide()

        # Calculate the position for the menu
        header_rect = header.geometry()
        pos = self.view.mapToGlobal(QPoint(header.sectionViewportPosition(column_index), header_rect.bottom()))

        # Get the screen geometry
        screen = QGuiApplication.screenAt(self.window().frameGeometry().center()).geometry()

        # Adjust horizontal position
        if pos.x() + menu_width > screen.right():
            pos.setX(screen.right() - menu_width)
        elif pos.x() < screen.left():
            pos.setX(screen.left())

        # Adjust vertical position
        if pos.y() + menu_height > screen.bottom():
            pos.setY(screen.bottom() - menu_height)
        elif pos.y() < screen.top():
            pos.setY(screen.top())

        # Show the menu at the adjusted position
        menu.exec(pos)

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
            
    def move_to_next_cell(self, horizontal=False): #tab key
        current = self.view.currentIndex()
        if horizontal:
            next_index = self.view.model().index(current.row(), current.column() + 1)
        else:
            next_index = self.view.model().index(current.row() + 1, current.column())
        if next_index.isValid():
            self.view.setCurrentIndex(next_index)

    def move_to_previous_cell(self): #
        current = self.view.currentIndex()
        previous_index = self.view.model().index(current.row(), current.column() - 1)
        if previous_index.isValid():
            self.view.setCurrentIndex(previous_index)
            
    def copy_to_clipboard(self):
        selection = self.view.selectionModel()
        indexes = selection.selectedIndexes()
        if not indexes:
            return

        clipboard_content = ''
        previous = indexes[0]
        for i in indexes:
            if i.row() != previous.row():
                clipboard_content += '\n'
            elif i != indexes[0]:
                clipboard_content += '\t'
            clipboard_content += i.data()
            previous = i

        QApplication.clipboard().setText(clipboard_content)

    def paste_clipboard(self):
        clipboard = QApplication.clipboard()
        clipboard_content = clipboard.text()

        # Split clipboard content into rows and cells
        clipboard_rows = clipboard_content.split('\n')
        clipboard_data = [row.split('\t') for row in clipboard_rows if row]

        # Get the first selected cell or current cell
        selection = self.view.selectionModel()
        if selection.hasSelection():
            start_index = selection.selectedIndexes()[0]
        else:
            start_index = self.view.currentIndex()

        if not start_index.isValid():
            return  # No starting point for pasting

        for r, row in enumerate(clipboard_data):
            for c, cell in enumerate(row):
                row_index = start_index.row() + r
                col_index = start_index.column() + c
                if row_index < self.model.rowCount() and col_index < self.model.columnCount():
                    item = QStandardItem(cell)
                    self.model.setItem(row_index, col_index, item)

    def clear_selected_cells(self):
        selection = self.view.selectionModel()
        for index in selection.selectedIndexes():
            self.model.setItem(index.row(), index.column(), QStandardItem(""))
                   
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
        
        self.update_job = QAction("Update Job Number", self)
        self.import_action = QAction("Import .dbf Files", self)
        self.import_specs = QAction("Import Linde Specs", self)
        
        self.update_job.triggered.connect(self.set_job_number)
        self.import_action.triggered.connect(self.import_dbf_welds)
        self.import_specs.triggered.connect(self.import_linde_specs)
        
        self.file_menu.addAction(self.update_job)
        self.file_menu.addAction(self.import_action)
        self.file_menu.addAction(self.import_specs)

        #Using CustomTableWidget Class
        self.table1 = StandardTable(self.tab1)
        self.table2 = StandardTable(self.tab2)
        self.table3 = StandardTable(self.tab3)
        
        #Clear/Reset Filters
        self.table1.resetFilters()
        
        # Create a QVBoxLayout for each tab
        self.layout1 = QVBoxLayout(self.tab1)
        self.layout2 = QVBoxLayout(self.tab2)
        self.layout3 = QVBoxLayout(self.tab3)
        
        # Create "Push to QuickBase" button for each tab
        self.push_to_qb_button1 = QPushButton("Push Spools to QuickBase", self.tab1)
        self.export_button1 = QPushButton("Export", self.tab1)
        self.push_to_qb_button2 = QPushButton("Push Welds to QuickBase", self.tab2)
        self.export_button2 = QPushButton("Export", self.tab1)
        self.push_to_qb_button3 = QPushButton("Push NDE Specs to QuickBase", self.tab3)

        #Connect Buttons
        self.push_to_qb_button1.clicked.connect(self.push_to_tpsl)
        self.export_button1.clicked.connect(lambda: self.export_table_to_excel('tpsl'))
        
        self.push_to_qb_button2.clicked.connect(self.push_to_weld_log)
        self.export_button2.clicked.connect(lambda: self.export_table_to_excel('weld log'))
        
        self.push_to_qb_button3.clicked.connect(self.push_to_nde)

        # Add QTableWidget and button to the layout for each tab
        self.layout1.addWidget(self.push_to_qb_button1, stretch=0)  # No extra space to button
        self.layout1.addWidget(self.export_button1, stretch=0)  # No extra space to button
        self.layout1.addWidget(self.table1, stretch=1)             # All extra space to table
        
        self.layout2.addWidget(self.push_to_qb_button2, stretch=0)
        self.layout2.addWidget(self.export_button2, stretch=0)
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
                #Clear/Reset Filters
                self.table1.resetFilters()
                return job_number
            else:
                response = QMessageBox.question(self, "No Job Number", "You didn't enter a job number. Do you want to try again?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if response == QMessageBox.StandardButton.No:
                    return None

    def get_job_number(self):
        return self.current_job_number  # Retrieve the stored job number
    
    def confirm_push_to_quickbase(self, table): #Allow confirm
        job_number = self.get_job_number()
        message = f"Are you sure you want to push these records to QuickBase for Job '{job_number}'?"
    
        reply = QMessageBox.question(self, f"{job_number} - Confirm Push - {table}", message, 
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel, 
                                     QMessageBox.StandardButton.Cancel)

        if reply == QMessageBox.StandardButton.Yes:
            #return True
            return self.show_countdown(5)  # Countdown duration in seconds
        else:
            return False
    
    def show_countdown(self, duration):
        self.countdown = duration
        self.msg = QMessageBox(self)
        self.msg.setIcon(QMessageBox.Icon.Information)
        self.msg.setWindowTitle("Countdown")
        self.msg.setText(f"Proceeding in {self.countdown} seconds...\nClick Cancel to abort.")
        self.msg.setStandardButtons(QMessageBox.StandardButton.Cancel)
        self.msg.buttonClicked.connect(self.abort_countdown)

        self.timer = QTimer(self)
        self.timer.timeout.connect(lambda: self.update_countdown())
        self.timer.start(1000)

        result = self.msg.exec()
        self.timer.stop()

        if result == QMessageBox.StandardButton.Cancel:
            QMessageBox.information(self, "Aborted", "Push Aborted")
            return False
        elif self.countdown <= 0:
            return True
        else:
            return False

    def update_countdown(self):
        self.countdown -= 1
        if self.countdown <= 0:
            self.msg.done(0)
        else:
            self.msg.setText(f"Proceeding in {self.countdown} seconds...\nClick Cancel to abort.")

    def abort_countdown(self, i):
        self.timer.stop()
        self.countdown = 0
        self.msg.done(QMessageBox.StandardButton.Cancel)

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

        #print("\n\nRESPONSE:", response.json())

        return response

    def create_df_from_table(self, table_name):
        # Select the appropriate table based on the table_name
        if table_name == 'tpsl':
            standard_table = self.table1

        elif table_name == 'weld log':
            standard_table = self.table2

        elif table_name == 'nde':
            standard_table = self.table3
        # Add more conditions here for other tables
        else:
            raise ValueError(f"Unknown table name: {table_name}")
        
        model = standard_table.model  # Access the model of the StandardTable

        # Get the number of rows and columns from the model
        rows = model.rowCount()
        cols = model.columnCount()

        # Extract the header names
        headers = [model.horizontalHeaderItem(c).text() for c in range(cols)]

        # Extract the data from the table
        data = []
        for r in range(rows):
            row_data = []
            for c in range(cols):
                item = model.item(r, c)
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
        #print("\n\nEXISTING DF: \n", existing_df)
        # Verify that compare_columns are in df_mapped
        # for col in compare_columns:
        #     if col not in df_mapped.columns:
        #         raise ValueError(f"Column '{col}' not found in DataFrame")

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
                    # existing_row = existing_df[existing_df['unique_id'] == row['unique_id']].iloc[0]
                    # duplicate_map.append(f"{existing_row[compare_columns[1]]}:{existing_row['3']}")
                    existing_row = existing_df[existing_df['unique_id'] == row['unique_id']].iloc[0]
                    duplicate_map.append(f"{existing_row['unique_id']}:{existing_row['3']}")
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
            "fieldsToReturn": [3, 6, 108] #+ [int(fid) for fid in mapping.values()]
        }

        # Logging the number of rows processed and ignored
        print(f"Original row count: {len(df_mapped)}, Payload row count: {len(payload_data)}, Duplicates ignored: {len(duplicate_map)}")
        #print(f"Duplicate Map: {duplicate_map}")

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
        #print("\n\nPayload Prepared - 'NDE - JOB LOT NDE'\n", payload)

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
        if not self.confirm_push_to_quickbase("Welds"):
            return  # Immediately return if the user decides not to proceed or cancels during countdown     

        #Create a dataframe from tab1
        df = self.create_df_from_table('weld log')
        
        # Initialize a list to hold skipped records
        skipped_records = []
        
        # Iterate over the DataFrame to validate 'Related Record' column
        for index, row in df.iterrows():
            related_record = row.get('Related Record')
            try:
                # Attempt to convert 'Related Record' to an integer
                int_related_record = int(related_record) if related_record is not None else None

                # Check if conversion is successful and value is non-negative
                if int_related_record is not None and int_related_record >= 0:
                    continue  # Valid entry, proceed to next row
                else:
                    # Conversion failed or value is negative, skip this row
                    skipped_records.append((index, related_record))
                    df.drop(index, inplace=True)  # Remove invalid row from DataFrame
            except ValueError:
                # Conversion to integer failed, skip this row
                skipped_records.append((index, related_record))
                df.drop(index, inplace=True)  # Remove invalid row from DataFrame
        
        #print("\n\nDATAFRAME VALID: \n", df)

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
        # print("\n\nPayload Prepared - 'Welds - Weld Log'\n", payload)
        # print("\n\nPayload Prepared - 'Welds - Weld Log'\n", existing_records_df)
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
            
        # Print skipped records if any
        if skipped_records:
            print("\nSkipped Records (Invalid Related Records):")
            for index, invalid_value in skipped_records:
                print(f"Row {index}, Invalid Value: {invalid_value}")
        
    def push_to_tpsl(self):
        if not self.confirm_push_to_quickbase("Spools"):
            return  # Immediately return if the user decides not to proceed or cancels during countdown 
        
        # Post-check rules
        weld_log_post_checks = {
            'required_columns': ['Related Record'],
            'integer_columns': ['Related Record'],
            'invalid_format_columns': ['Related Record']
        }

        tpsl_log_post_checks = {
            'invalid_format_columns': ['TPSL Record']
        }

        #Create a dataframe from tab1
        df = self.create_df_from_table('tpsl')

        # Load the mapping from the JSON file
        with open('spool_map.json', 'r') as f:
            mapping = json.load(f)

         #Check if duplicate records exists
        job_number = self.get_job_number()
        qb_fields=['3','6','108']
        existing_records_df = get_table_data(qb_fields, job_number, TPSL_ID)

        payload, row_count, duplicate_map = self.prepare_payload(df, mapping, TPSL_ID, existing_records_df, ['6', '108'])
        #print("\n\nPayload Prepared - 'Spools - TPSL'\n", payload)
        
        df['unique_id'] = df['Job'].astype(str) + '-' + df['Spool'].astype(str)

        # If duplicate_map has items, convert it to a dictionary
        if duplicate_map:
            duplicate_map_dict = dict(item.split(":") for item in duplicate_map)
        else:
            duplicate_map_dict = {}
            
        #Create Welds Dataframe  
        df_welds = self.create_df_from_table('weld log')

        #Abort if no rows
        if row_count < 1:
            # Update the 'TPSL Record' column using the duplicate_map_dict
            df['TPSL Record'] = df['unique_id'].map(duplicate_map_dict)
            
            # Applying post-checks to Spool Table
            print("\nPerforming Post-Checks - Spools...")
            df = post_check(df, **tpsl_log_post_checks)
            print("\nPost-Checks Complete - Spools...")
            
            # Reload the table1 with updated DataFrame
            self.load_data_into_table(self.table1, df)
            print("\n\n'Spools' table reloaded with updated existing records...")
            #print(f"\n\nDuplicate Map: \n{duplicate_map_dict}")

            # Update the Welds dataframe with record IDs
            df_welds['unique_id'] = df_welds['Job'].astype(str) + '-' + df_welds['Spool'].astype(str)
            df_welds['Related Record'] = df_welds['unique_id'].map(duplicate_map_dict)
            
            # Performing Welds Post-Check...
            print("\nPerforming Post-Check - Welds...")
            df_welds = post_check(df_welds, **weld_log_post_checks)
            print("\nPost-Check Complete - Welds...")
            
            # Reload the table2 (Welds) with updated DataFrame
            self.load_data_into_table(self.table2, df_welds)
            print("\n\n'Welds' table reloaded with updated existing records...")

            error_dialog = QErrorMessage()
            error_dialog.showMessage('Payload is empty. \nNote: Duplicate records are automatically removed. \n0 Rows added to QuickBase.')
            error_dialog.exec()  # Use exec_() to make sure the dialog is modal and waits for user input
            #print(f"\n\nDuplicate Map: \n{duplicate_map_dict}")
            return
        
        # Push to QuickBase and capture the response
        response = self.push_to_qb(payload)
        
        # Check if the response status code is 200
        if response.status_code in [200, 207]:
            response_data = response.json()
            QMessageBox.information(None, 'Success', f'Respond Status: {response.status_code} \n\n({row_count}) records added to QuickBase successfully.')

             # Initialize a map to hold the mapping of unique identifiers to record IDs
            unique_id_to_record_id = {}
            
            # Extract data from the response
            if 'data' in response_data:
                for item in response_data['data']:
                    job = item['6']['value']
                    spool = item['108']['value']
                    record_id = item['3']['value']
                    unique_id = f"{job}-{spool}"
                    unique_id_to_record_id[unique_id] = record_id

            # If duplicate_map has items, convert it to a dictionary
            if duplicate_map:
                duplicate_map_dict = dict(item.split(":") for item in duplicate_map)
            else:
                duplicate_map_dict = {}

            # Create a unique identifier in the original dataframe
            df['unique_id'] = df['Job'].astype(str) + '-' + df['Spool'].astype(str)
            
            # Combine the unique_id_to_record_id and duplicate_map_dict for the final mapping
            combined_record_map = {**unique_id_to_record_id, **duplicate_map_dict}
            
            # Map the record IDs to the DataFrame using the unique identifier
            df['TPSL Record'] = df['unique_id'].map(combined_record_map)
            
            # Performing spools Post-Check...
            print("\nPerforming Post-Check - Spools...")
            df_welds = post_check(df_welds, **weld_log_post_checks)
            print("\nPost-Check Complete - Spools...")

            self.load_data_into_table(self.table1, df)
            # print("\n\n'Spools' table reloaded with updated records...\nExcluded Mapping: ", duplicate_map)

            # Update 'Related Record' in table2 (Welds)
            df_welds['unique_id'] = df_welds['Job'].astype(str) + '-' + df_welds['Spool'].astype(str)
            df_welds['Related Record'] = df_welds['unique_id'].map(combined_record_map)
            
            # with pd.ExcelWriter('output.xlsx') as writer:
            #     df_welds.to_excel(writer, sheet_name='Welds')
            #     df.to_excel(writer, sheet_name='Spools')

            # Performing Welds Post-Check...
            print("\nPerforming Post-Check - Welds...")
            df_welds = post_check(df_welds, **weld_log_post_checks)
            print("\nPost-Check Complete - Welds...")
            
            # Reload the table2 (Welds) with updated DataFrame
            self.load_data_into_table(self.table2, df_welds)
            
            print("\n\n'Welds' table reloaded with updated records...")

    def load_data_into_table(self, standard_table, df):
        print("\n\nLoading Data into Tables")
        model = standard_table.model  # Access the model of the StandardTable

        # Clear existing data from the model
        model.clear()

        # Load the SVG icon
        icon_path = "assets/filter_icon.svg"
        icon = QIcon(icon_path)
        
        # # Set the column headers
        model.setHorizontalHeaderLabels(df.columns.tolist())
        
        # Populate the model with data
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
                # Create a QStandardItem for each cell
                item = QStandardItem(str(value))

                # Add the item to the model
                model.setItem(i, j, item)
                
        # Autosize columns and update icons in one loop
        icon_width = 20  # Approximate width of the icon
        padding = 10     # Additional padding
        header_view = standard_table.view.horizontalHeader()
        if isinstance(header_view, CustomHeaderView):
            for col_number in range(df.shape[1]):
                # Update the icon for the column
                header_view.setIconForColumn(col_number)

                # Autosize column
                standard_table.view.resizeColumnToContents(col_number)

                # Adjust column width for icon width
                current_width = standard_table.view.columnWidth(col_number)
                standard_table.view.setColumnWidth(col_number, current_width + icon_width + padding)
                     
        # Redraw the header view
        standard_table.view.horizontalHeader().viewport().update()

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
            print("\n", file)
            spools_df = self.read_dbf(file)
            self.load_dbf_files(welds_df, spools_df, job_number)
            
    def read_dbf(self, file_path):
        dbf = Dbf5(file_path)
        return dbf.to_dataframe()
    
    def load_dbf_files(self, welds_df, spools_df, job_number):
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

        # Define required columns for pre-check
        welds_required_columns = ['Job', 'Spool', 'Weld ID']
        spools_required_columns = ['Spool', 'Ref Drawing', 'Series', 'NDE Group', 'Job', 'Sheet']

        # Perform pre-checks
        print("\nInitiating Pre-Check - Spools...")
        formatted_spools_df = pre_check(formatted_spools_df, spools_required_columns, integer_columns=['Sheet'])
        print("\nPre-Check Complete - Spools...")
        
        print("\nInitiating Pre-Check - Welds...")
        formatted_welds_df = pre_check(formatted_welds_df, welds_required_columns, hyphen_columns=['Job'])
        print("\nPre-Check Complete - Welds...")
        
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

    def export_table_to_excel(self, table_name):
        # Create DataFrame from table

        print("\nTriggered Export:", table_name)
        df = self.create_df_from_table(table_name)

        # Define initial filename based on table_name
        if table_name == 'tpsl':
            base_filename = 'FabManager Spool Data'
        elif table_name == 'weld log':
            base_filename = 'FabManager Weld Data'
        elif table_name == 'nde':
            base_filename = 'FabManager NDE Data'
        else:
            raise ValueError(f"Unknown table name: {table_name}")

        # Check for existing file and handle naming
        filename = base_filename
        count = 1
        while os.path.exists(f'{filename}.xlsx'):
            filename = f"{base_filename}({count})"
            count += 1
        full_filename = f'{filename}.xlsx'

        # Try to save the DataFrame to an Excel file
        try:
            # Attempt to open the file to check if it's in use
            with open(full_filename, 'a') as f:
                pass
            with pd.ExcelWriter(full_filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Data', index=True)
            print(f"Data exported successfully to {full_filename}")
        except IOError as e:
            print(f"Could not write to {full_filename}. The file may be open or in use. Please close the file and try again.")

if __name__ == "__main__":
    app = QApplication([])
    window = MyWindow()
    window.showMaximized()
    app.exec()
