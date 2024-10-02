import sys
import os
import pandas as pd
import pyodbc
import numpy as np
from datetime import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QTextEdit, QLabel, QComboBox, QLineEdit, QMainWindow, QPlainTextEdit, QTextBrowser, QHBoxLayout
from PyQt5.QtGui import QIcon, QFont, QDesktopServices, QColor, QTextCharFormat
from PyQt5.QtCore import Qt, QUrl
import ctypes
import markdown

 

class ScriptWindow(QMainWindow):

    def __init__(self, script):

        super().__init__()
        
        self.initUI(script)

 

    def initUI(self, script):

        self.setWindowTitle('Generated T-SQL Merge Script')

        self.setGeometry(800, 800, 800, 600)

        icon_path = os.path.join(os.path.dirname(__file__), 'savicon.png')

        if os.path.exists(icon_path):

            self.setWindowIcon(QIcon(icon_path))

 

        # Create a QPlainTextEdit widget to display the script

        self.scriptEdit = QPlainTextEdit(self)

        self.scriptEdit.setPlainText(script)

        self.scriptEdit.setReadOnly(True)

 

        # Set a monospace font for better code readability

        font = QFont("Courier")

        font.setStyleHint(QFont.Monospace)

        self.scriptEdit.setFont(font)

 

        self.setCentralWidget(self.scriptEdit)

 

class AboutWindow(QMainWindow):

    def __init__(self):

        super().__init__()

        self.initUI()

 

    def initUI(self):

        self.setWindowTitle('About T-SQL Merge Script Generator')

        self.setGeometry(600, 600, 600, 600)

       

        # Set the window icon

        icon_path = os.path.join(os.path.dirname(__file__), 'savicon.png')

        if os.path.exists(icon_path):

            self.setWindowIcon(QIcon(icon_path))

 

        # Create a QTextBrowser widget to display the README content

        self.textBrowser = QTextBrowser(self)

        self.textBrowser.setOpenExternalLinks(True)  # Allow opening external links

       

        # Load and display the README.md content

        self.load_readme()

 

        self.setCentralWidget(self.textBrowser)

 

    def load_readme(self):

        readme_path = os.path.join(os.path.dirname(__file__), 'README.md')

        if os.path.exists(readme_path):

            with open(readme_path, 'r', encoding='utf-8') as file:

                content = file.read()

                # Convert Markdown to HTML

                html_content = markdown.markdown(content)

                self.textBrowser.setHtml(html_content)

        else:

            self.textBrowser.setPlainText("README.md file not found.")

 

    def openUrl(self, url):

        QDesktopServices.openUrl(QUrl(url))

 

class SQLInsertGenerator(QWidget):

    def __init__(self):

        super().__init__()

        self.initUI()

       

    def initUI(self):

 

        # Set the window icon

        icon_path = os.path.join(os.path.dirname(__file__), 'savicon.png')

        if os.path.exists(icon_path):

            self.setWindowIcon(QIcon(icon_path))

        else:

            print(f"Warning: Icon file not found at {icon_path}")

 

        layout = QVBoxLayout()

        # Server name input

        self.server_label = QLabel('Server Name:')

        layout.addWidget(self.server_label)

        self.server_input = QLineEdit()

        self.server_input.setText('Localhost')

        layout.addWidget(self.server_input)

 

        # Database name input

        self.db_label = QLabel('Database Name:')

        layout.addWidget(self.db_label)

        self.db_input = QLineEdit()

        self.db_input.setText('TestDB')

        layout.addWidget(self.db_input)

 

        # Schema name input

        self.schema_label = QLabel('Schema Name:')

        layout.addWidget(self.schema_label)

        self.schema_input = QLineEdit()

        self.schema_input.setText('dbo')

        layout.addWidget(self.schema_input)

 

        # Table name input

        self.table_label = QLabel('Table Name:')

        layout.addWidget(self.table_label)

        self.table_input = QLineEdit()

        self.table_input.setText('testTable')

        layout.addWidget(self.table_input)

 

        # File selection

        self.file_btn = QPushButton('Select Source Excel or CSV File Provided')

        self.file_btn.setStyleSheet("color: blue; font-weight: bold;")

        self.file_btn.clicked.connect(self.select_file)

        layout.addWidget(self.file_btn)

 

        # File label

        self.file_label = QLabel('No file selected')

        self.file_label.setAlignment(Qt.AlignCenter)

        self.file_label.setStyleSheet("color: red; font-weight: bold;")

        layout.addWidget(self.file_label)

 

        # Sheet selection for Excel files

        self.sheet_combo = QComboBox()

        self.sheet_combo.setVisible(False)

        self.sheet_combo.currentIndexChanged.connect(self.update_preview)

        layout.addWidget(self.sheet_combo)

 

        # Data Preview

        self.preview = QTextEdit()

        self.preview.setReadOnly(True)

        self.preview.setVisible(False)

        layout.addWidget(self.preview)

 

        # Data preview

        self.output = QTextEdit()

        self.output.setReadOnly(True)

        self.output.setAcceptRichText(True)

        self.output.setPlaceholderText('Validation of files and table logging, if everything is ok, a script will be generated.')

        layout.addWidget(self.output)

 

        # Create a horizontal layout for buttons

        button_layout = QHBoxLayout()

 

        # Generate script button

        self.generate_btn = QPushButton('Validate Source and Destination')

        self.generate_btn.setStyleSheet("color: green; font-weight: bold;")

        self.generate_btn.clicked.connect(self.validate_and_generate)

        button_layout.addWidget(self.generate_btn)

 

        # Add Reset button

        self.reset_btn = QPushButton('Reset')

        self.reset_btn.setStyleSheet("color: orange; font-weight: bold;")

        self.reset_btn.clicked.connect(self.reset_application)

        button_layout.addWidget(self.reset_btn)

 

        # Add About button

        self.about_btn = QPushButton('About')

        self.about_btn.setStyleSheet("color: black; font-weight: bold;")

        self.about_btn.clicked.connect(self.show_about)

        button_layout.addWidget(self.about_btn)

 

        # Add the button layout to the main layout

        layout.addLayout(button_layout)

 

        self.setLayout(layout)

        self.setGeometry(800, 800, 800, 800)

        self.setWindowTitle('T-SQL Merge Script Generator v1.0')

        self.show()

 

        # # Log and script output

        # self.output = QTextEdit()

        # self.output.setReadOnly(True)

        # self.output.setPlaceholderText('Validation of files and table logging, if everything is ok, a script will be generated.')

        # layout.addWidget(self.output)

 

        self.setLayout(layout)

        self.setGeometry(800, 800, 800, 800)

        self.setWindowTitle('T-SQL Merge Script Generator v1.0')

        self.show()

 

    def reset_application(self):

    # Clear the file selection

        self.file_label.setText('No file selected')

        self.file_label.setStyleSheet("color: red; font-weight: bold;")

        if hasattr(self, 'file_path'):

            delattr(self, 'file_path')

 

        # Clear the DataFrame

        if hasattr(self, 'df'):

            delattr(self, 'df')

 

        # Clear the preview

        self.preview.clear()

        self.preview.setVisible(False)

 

        # Clear the output logs

        self.output.clear()

 

        # Reset the sheet combo box

        self.sheet_combo.clear()

        self.sheet_combo.setVisible(False)

 

        # Log the reset action

        self.log_message("Application has been reset.")

 

    def log_message(self, message, message_type='info'):

        if message_type == 'warning':

            format = QTextCharFormat()

            format.setForeground(QColor('orange'))

            format.setFontWeight(QFont.Bold)

            self.output.mergeCurrentCharFormat(format)

            self.output.append(f"WARNING: {message}")

        elif message_type == 'error':

            format = QTextCharFormat()

            format.setForeground(QColor('red'))

            format.setFontWeight(QFont.Bold)

            self.output.mergeCurrentCharFormat(format)

            self.output.append(f"ERROR: {message}")

        else:

            format = QTextCharFormat()

            format.setForeground(QColor('green'))

            self.output.mergeCurrentCharFormat(format)

            self.output.append(f"INFO: {message}")

       

        # Reset the format to default after appending

        format = QTextCharFormat()

        format.setForeground(QColor('black'))

        format.setFontWeight(QFont.Normal)

        self.output.mergeCurrentCharFormat(format)

 

    def show_about(self):

        self.about_window = AboutWindow()

        self.about_window.show()

 

    def get_connection_string(self):

        server_name = self.server_input.text()

        db_name = self.db_input.text()

        return f'DRIVER={{SQL Server}};SERVER={server_name};DATABASE={db_name};Trusted_Connection=yes;Integrated Security=SSPI;'

 

    def select_file(self):

        file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel or CSV File", "", "Excel Files (*.xlsx *.xls);;CSV Files (*.csv)")

        if file_path:

            file_name = os.path.basename(file_path)

            self.file_label.setText(f"Selected file: {file_name}")

            self.file_label.setStyleSheet("color: blue; font-weight: bold;")

            self.output.append(f"Selected file: {file_path}")

            self.load_file(file_path)

        else:

            self.file_label.setText("No file selected")

            self.file_label.setStyleSheet("color: red; font-weight: bold;")

 

    def load_file(self, file_path):

        self.file_path = file_path  # Store the file path

        file_extension = os.path.splitext(file_path)[1].lower()

        try:

            if file_extension in ['.xlsx', '.xls']:

                try:

                    # Read Excel file

                    self.excel = pd.ExcelFile(file_path)

                    self.sheet_combo.clear()

                    self.sheet_combo.addItems(self.excel.sheet_names)

                    self.sheet_combo.setVisible(True)

                   

                    # Read the first sheet by default

                    self.df = pd.read_excel(self.excel, sheet_name=0, dtype_backend='numpy_nullable')

                   

                    # Convert float columns that are actually integers

                    for col in self.df.columns:

                        if self.df[col].dtype == 'float64':

                            if self.df[col].apply(lambda x: x.is_integer()).all():

                                self.df[col] = self.df[col].astype('Int64')

                   

                    self.update_preview()

                except ImportError as e:

                    if 'openpyxl' in str(e):

                        self.output.append("Error: Missing 'openpyxl' package. Please install it using 'pip install openpyxl' and restart the application.")

                        self.file_label.setText("Error loading file (missing openpyxl)")

                    else:

                        raise

            elif file_extension == '.csv':

                # Read CSV file

                self.df = pd.read_csv(file_path, dtype_backend='numpy_nullable')

               

                # Convert float columns that are actually integers

                for col in self.df.columns:

                    if self.df[col].dtype == 'float64':

                        if self.df[col].apply(lambda x: x.is_integer()).all():

                            self.df[col] = self.df[col].astype('Int64')

               

                self.sheet_combo.setVisible(False)

                self.update_preview()

            else:

                self.output.append(f"Unsupported file type: {file_extension}")

                self.file_label.setText("Unsupported file type")

        except Exception as e:

            self.output.append(f"Error loading file: {str(e)}")

            self.file_label.setText("Error loading file")

 

    def update_preview(self):

        try:

            if hasattr(self, 'excel'):

                sheet_name = self.sheet_combo.currentText()

                self.df = self.excel.parse(sheet_name)

            if hasattr(self, 'df'):

                self.preview.setVisible(True)

                self.preview.setText(self.df.head().to_string())

            else:

                self.preview.setVisible(False)

        except Exception as e:

            self.output.append(f"Error updating preview: {str(e)}")

 

    def validate_table_structure(self):

        connection_string = self.get_connection_string()

        schema = self.schema_input.text()

        table = self.table_input.text()

 

        try:

            with pyodbc.connect(connection_string) as conn:

                cursor = conn.cursor()

               

                # Check if the table exists

                cursor.execute(f"""

                    SELECT TABLE_NAME

                    FROM INFORMATION_SCHEMA.TABLES

                    WHERE TABLE_SCHEMA = ? AND TABLE_NAME = ?

                """, (schema, table))

               

                if cursor.fetchone() is None:

                    self.log_message(f"Error: Table [{schema}].[{table}] does not exist.", 'error')

                    return False

 

                # Get column information, excluding identity columns

                cursor.execute(f"""

                    SELECT

                        c.COLUMN_NAME,

                        c.DATA_TYPE,

                        c.CHARACTER_MAXIMUM_LENGTH,

                        c.IS_NULLABLE,

                        COLUMNPROPERTY(OBJECT_ID(c.TABLE_SCHEMA + '.' + c.TABLE_NAME), c.COLUMN_NAME, 'IsIdentity') AS IS_IDENTITY,

                        CASE WHEN c.COLUMN_DEFAULT IS NOT NULL THEN 1 ELSE 0 END AS HAS_DEFAULT

                    FROM INFORMATION_SCHEMA.COLUMNS c

                    WHERE c.TABLE_SCHEMA = ? AND c.TABLE_NAME = ?

                    ORDER BY c.ORDINAL_POSITION

                """, (schema, table))

               

                db_columns = cursor.fetchall()

               

                if not db_columns:

                    self.output.append(f"Error: No columns found in table [{schema}].[{table}].")

                    return False

 

                # Filter out identity columns and separate required and optional columns

                required_columns = []

                optional_columns = []

                for col in db_columns:

                # if not col.IS_IDENTITY:  # Ignore identity columns

                    if col.IS_NULLABLE == 'NO' and not col.HAS_DEFAULT:

                        required_columns.append(col.COLUMN_NAME)

                    else:

                        optional_columns.append(col.COLUMN_NAME)

 

                # Compare with DataFrame columns

                df_columns = set(self.df.columns)

                required_db_columns = set(required_columns)

                optional_db_columns = set(optional_columns)

 

                missing_required = required_db_columns - df_columns

                extra_columns = df_columns - (required_db_columns | optional_db_columns)

 

                if missing_required:

                    self.log_message(f"Error: Missing required columns in input data: {', '.join(missing_required)}", 'error')

                    return False

 

                if extra_columns:

                    self.log_message(f"Warning: Extra columns in input data (will be ignored): {', '.join(extra_columns)}", 'warning')

 

                # Store column information for later use

                self.table_columns = {col.COLUMN_NAME: {

                    'data_type': col.DATA_TYPE,

                    'max_length': col.CHARACTER_MAXIMUM_LENGTH,

                    'is_nullable': col.IS_NULLABLE,

                    'has_default': col.HAS_DEFAULT

                } for col in db_columns} #if not col.IS_IDENTITY}

 

                # Get foreign key information including nullability

                cursor.execute(f"""

                    SELECT

                        COL_NAME(fc.parent_object_id, fc.parent_column_id) AS ColumnName,

                        OBJECT_SCHEMA_NAME(f.referenced_object_id) AS ReferencedSchemaName,

                        OBJECT_NAME(f.referenced_object_id) AS ReferencedTableName,

                        COL_NAME(fc.referenced_object_id, fc.referenced_column_id) AS ReferencedColumnName,

                        c.is_nullable AS IsNullable

                    FROM

                        sys.foreign_keys AS f

                    INNER JOIN

                        sys.foreign_key_columns AS fc ON f.object_id = fc.constraint_object_id

                    INNER JOIN

                        sys.columns AS c ON fc.parent_object_id = c.object_id AND fc.parent_column_id = c.column_id

                    WHERE

                        OBJECT_SCHEMA_NAME(f.parent_object_id) = ? AND OBJECT_NAME(f.parent_object_id) = ?

                """, (schema, table))

               

                self.foreign_keys = cursor.fetchall()

 

                self.output.append("Table structure and foreign key validation passed.")

                return True

 

        except pyodbc.Error as e:

            self.output.append(f"Database error: {str(e)}")

            return False

        except Exception as e:

            self.output.append(f"Error validating table structure: {str(e)}")

            return False

 

    def validate_data(self):

        self.validation_errors = []

        if not hasattr(self, 'table_columns'):

            self.log_message("Table structure not validated. Please run structure validation first.", 'error')

            return False

 

        null_values = [None, '', 'NULL', 'null', '<N/A>', '<n/a>']

 

        for column, info in self.table_columns.items():

            if column not in self.df.columns:

                self.add_validation_error(f"Column '{column}' not in the destinations table columns. Please check the file again.")

                continue

 

            # Replace null-like values with None for consistent handling

            self.df[column] = self.df[column].replace(null_values, None)

 

            # Check data types

            if info['data_type'] in ['int', 'bigint', 'smallint', 'tinyint']:

                non_numeric = self.df[~self.df[column].apply(lambda x: pd.isnull(x) or np.issubdtype(type(x), np.number))][column]

                if not non_numeric.empty:

                    self.add_validation_error(f"Column '{column}' contains non-numeric values in rows: {non_numeric.index.tolist()}")

           

            # Check BITS data type

            elif info['data_type'] == 'bit':

                valid_bits = [1, 0, '1', '0', 'True', 'False', True, False]

                invalid_bits = self.df[~self.df[column].isin(valid_bits) & ~self.df[column].isnull()]

                if not invalid_bits.empty:

                    self.add_validation_error(f"Column '{column}' contains invalid BIT values in rows: {invalid_bits.index.tolist()}")

           

            # Check string length

            if info['data_type'] in ['varchar', 'nvarchar', 'char', 'nchar'] and info['max_length']:

                too_long = self.df[self.df[column].notna() & (self.df[column].astype(str).str.len() > info['max_length'])]

                if not too_long.empty:

                    self.add_validation_error(f"Data in column '{column}' exceeds maximum length of {info['max_length']} in rows: {too_long.index.tolist()}")

 

            # Check for nulls in non-nullable columns

            if info['is_nullable'] == 'NO':

                null_rows = self.df[self.df[column].isnull()].index.tolist()

                if null_rows:

                    if info['has_default'] == 0:

                        self.add_validation_error(f"Column '{column}' cannot contain null values. Null found in rows {null_rows}")

                    else:

                        self.add_validation_error(f"Column '{column}' cannot contain null values even though it has a default. Null found in rows: {null_rows}")

 

        # Validate foreign keys

        if hasattr(self, 'foreign_keys'):

            for fk in self.foreign_keys:

                if fk.ColumnName in self.df.columns:

                    self.validate_foreign_key(fk)

 

        if len(self.validation_errors) >= 1:

            self.log_message("Data validation failed. Please check the errors below:", 'error')

            for error in self.validation_errors:

                self.log_message(error, 'error')

            return False

        else:

            self.log_message("Data validation passed.")

            return True

   

    def validate_foreign_key(self, fk):

        connection_string = self.get_connection_string()

        try:

            with pyodbc.connect(connection_string) as conn:

                cursor = conn.cursor()

               

                # Check for NULL values in non-nullable foreign key column

                if not fk.IsNullable:

                    null_rows = self.df[self.df[fk.ColumnName].isnull()]

                    if not null_rows.empty:

                        error_rows = null_rows.index.tolist()

                        self.add_validation_error(f"Non-nullable foreign key column '{fk.ColumnName}' contains NULL values in rows: {error_rows}")

                        self.log_message(f"Non-nullable foreign key column '{fk.ColumnName}' contains NULL values in rows: {error_rows}", 'error')

               

                # Get non-null values from the DataFrame column

                non_null_values = self.df[fk.ColumnName].dropna()

               

                if len(non_null_values) == 0:

                    # If all values are NULL and the column is nullable, it's valid

                    if not fk.IsNullable:

                        self.add_validation_error(f"Non-nullable foreign key column '{fk.ColumnName}' contains only NULL values.")

                        self.log_message(f"Non-nullable foreign key column '{fk.ColumnName}' contains only NULL values.", 'error')

                    return

               

                # Convert numpy types to Python native types

                non_null_values = [self.numpy_to_python(val) for val in non_null_values]

               

                # Check if all non-null values exist in the referenced table

                placeholders = ', '.join(['?' for _ in non_null_values])

                query = f"""

                    SELECT [{fk.ReferencedColumnName}]

                    FROM [{fk.ReferencedSchemaName}].[{fk.ReferencedTableName}]

                    WHERE [{fk.ReferencedColumnName}] IN ({placeholders})

                """

               

                cursor.execute(query, non_null_values)

                existing_values = set([row[0] for row in cursor.fetchall()])

               

                invalid_values = set(non_null_values) - existing_values

                if invalid_values:

                    error_rows = self.df[self.df[fk.ColumnName].isin(invalid_values)].index.tolist()

                    self.log_message(f"Foreign key constraint violation in column '{fk.ColumnName}'. "

                                              f"The following values do not exist in the referenced table "

                                              f"[{fk.ReferencedSchemaName}].[{fk.ReferencedTableName}]:", 'error')

                    self.add_validation_error(f"Foreign key constraint violation in column '{fk.ColumnName}'. "

                                              f"The following values do not exist in the referenced table "

                                              f"[{fk.ReferencedSchemaName}].[{fk.ReferencedTableName}]:")

                    for value in invalid_values:

                        rows = self.df[self.df[fk.ColumnName] == value].index.tolist()

                        self.log_message(f"  - Value '{value}' in rows: {rows}", 'error')

                        self.add_validation_error(f"  - Value '{value}' in rows: {rows}")

 

        except pyodbc.Error as e:

            self.log_message(f"Database error while validating foreign key '{fk.ColumnName}': {str(e)}", 'error')

            self.add_validation_error(f"Database error while validating foreign key '{fk.ColumnName}': {str(e)}")

        except Exception as e:

            self.log_message(f"Error validating foreign key '{fk.ColumnName}': {str(e)}", 'error')

            self.add_validation_error(f"Error validating foreign key '{fk.ColumnName}': {str(e)}")

 

    def numpy_to_python(self, value):

        if isinstance(value, np.integer):

            return int(value)

        elif isinstance(value, np.floating):

            return float(value)

        elif isinstance(value, np.ndarray):

            return value.tolist()

        else:

            return value

 

    def add_validation_error(self, error_message):

        self.validation_errors.append(error_message)

        # self.log_message(error_message, 'error')

 

    def generate_merge_script(self):

        schema = self.schema_input.text()

        table = self.table_input.text()

 

        # Get the current date and time

        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

       

        # Get the file name from the stored file path

        file_name = os.path.basename(self.file_path) if hasattr(self, 'file_path') else "Unknown file"

 

        # Create the header comment

        header_comment = f"""

    /************************************************************************************

    * Populating [{schema}].[{table}] with Lookup data from {file_name}                 *

    *                                                                                   *

    * Created Date: {current_time}                                                      *

    * Created By: T-SQL Merge Script Generator                                          *

    *                                                                                   *

    * Updated Date:         By:                         Comment:                        *

    * {current_time}        Merge Script Generator      Initial script generation       *

    ************************************************************************************/

 

    """

       

        # Identify the primary key column(s)

        pk_columns = self.get_primary_key_columns(schema, table)

       

        # Check if primary key columns exist in the source file

        pk_columns_in_source = [col for col in pk_columns if col in self.df.columns]

       

        # If no primary key is found in the source data,

        # use the first column of the DataFrame as the matching column

        if not pk_columns_in_source:

            pk_columns_in_source = [self.df.columns[0]]

       

        # Use only the columns that exist in both the DataFrame and the table structure

        valid_columns = [col for col in self.df.columns if col in self.table_columns]

       

        # Identify the date column for soft delete, if it exists

        # soft_delete_column = next((col for col in ['InvalidatedDate', 'ValidToDate', 'DeletedDate'] if col in self.table_columns), None)

 

        # Prepare the MERGE statement

        merge_script = header_comment + f"""

    SET IDENTITY_INSERT [{schema}].[{table}] ON;

    GO

   

    SET NOCOUNT ON;

    GO

 

    DECLARE @SummaryOfChanges TABLE(Change VARCHAR(20));

 

    MERGE INTO [{schema}].[{table}] AS Target

    USING (VALUES

    """

       

        # Generate the VALUES clause

        values_rows = []

        for _, row in self.df.iterrows():

            row_values = []

            for col in valid_columns:

                value = row[col]

                col_info = self.table_columns[col]

               

                if pd.isna(value):

                    row_values.append('NULL')

                elif col_info['data_type'] in ['varchar', 'nvarchar', 'char', 'nchar', 'text', 'ntext', 'xml', 'date', 'datetime', 'datetime2', 'time']:

                    escaped_value = str(value).replace("'", "''")

                    row_values.append(f"'{escaped_value}'")

                elif col_info['data_type'] in ['bit']:

                    row_values.append('1' if value else '0')

                else:

                    row_values.append(str(value))

           

            values_rows.append(f"({', '.join(row_values)})")

       

        merge_script += ',\n'.join(values_rows)

       

        # Add the column names to the USING clause

        merge_script += f"\n) AS Source ({', '.join(valid_columns)})\n"

       

        # Add the ON clause using the primary key columns that exist in the source

        merge_script += f"ON ({' AND '.join([f'Target.{col} = Source.{col}' for col in pk_columns_in_source])})\n"

       

        # Add the WHEN MATCHED clause

        update_columns = [col for col in valid_columns if col not in pk_columns_in_source]

        if update_columns:

            merge_script += "WHEN MATCHED AND (\n    "

            merge_script += " OR \n    ".join([f"NULLIF(Source.{col}, Target.{col}) IS NOT NULL OR NULLIF(Target.{col}, Source.{col}) IS NOT NULL" for col in update_columns])

            merge_script += "\n) THEN\n UPDATE SET\n    "

            merge_script += ",\n    ".join([f"{col} = Source.{col}" for col in update_columns])

       

        # Add the WHEN NOT MATCHED BY TARGET clause

        merge_script += f"\nWHEN NOT MATCHED BY TARGET THEN\n INSERT({', '.join(valid_columns)})\n VALUES({', '.join([f'Source.{col}' for col in valid_columns])})"

       

        # # Add the WHEN NOT MATCHED BY SOURCE clause

        # if soft_delete_column:

        #     merge_script += f"\nWHEN NOT MATCHED BY SOURCE AND ISNULL((Target.{soft_delete_column}, '1900-01-01') <> ISNULL((Source.{soft_delete_column}), '1900-01-01') THEN\n UPDATE SET {soft_delete_column} = GETDATE()"

        # else:

        #     merge_script += "\n-- **** NO SOFT DELETE COLUMN FOUND. RECORDS NOT IN SOURCE WILL NOT BE MODIFIED. ****"

       

        # Add the OUTPUT clause

        merge_script += "\nOUTPUT $action INTO @SummaryOfChanges;"

       

        # Add the summary select

        merge_script += f"""

    SELECT Change, COUNT(*) AS CountPerChange

    FROM @SummaryOfChanges

    GROUP BY Change;

    GO

    SET NOCOUNT OFF;

    GO

 

    SET IDENTITY_INSERT [{schema}].[{table}] OFF;

    GO

    """

       

        return merge_script

       

    def get_primary_key_columns(self, schema, table):

        connection_string = self.get_connection_string()

        try:

            with pyodbc.connect(connection_string) as conn:

                cursor = conn.cursor()

                cursor.execute(f"""

                    SELECT COLUMN_NAME

                    FROM INFORMATION_SCHEMA.KEY_COLUMN_USAGE

                    WHERE OBJECTPROPERTY(OBJECT_ID(CONSTRAINT_SCHEMA + '.' + CONSTRAINT_NAME), 'IsPrimaryKey') = 1

                    AND TABLE_SCHEMA = ? AND TABLE_NAME = ?

                """, (schema, table))

                pk_columns = [row.COLUMN_NAME for row in cursor.fetchall()]

                return pk_columns

        except pyodbc.Error as e:

            self.output.append(f"Error retrieving primary key information: {str(e)}")

            return []

 

    def validate_and_generate(self):

        if not hasattr(self, 'df'):

            self.log_message("Please select a file first.", 'warning')

            return

 

        if not self.server_input.text() or not self.db_input.text() or not self.schema_input.text() or not self.table_input.text():

            self.log_message("Please enter server name, database, schema, and table names.", 'warning')

            return

 

        try:

            self.log_message("Validating table structure and foreign keys...")

            if not self.validate_table_structure():

                self.log_message("Table structure validation failed. Please check the errors above.", 'error')

                return

 

            self.log_message("Validating data...")

            if not self.validate_data():

                self.log_message(f"Found {len(self.validation_errors)} validation errors.", 'error')

                return

 

            self.log_message("Validation passed. Generating MERGE script...")

            script = self.generate_merge_script()

            if script:

                self.log_message("MERGE script generated successfully.")

                self.show_script_window(script)

            else:

                self.log_message("Failed to generate MERGE script.", 'error')

 

        except Exception as e:

            self.log_message(f"Error: {str(e)}", 'error')

 

    def show_script_window(self, script):

        self.script_window = ScriptWindow(script)

        self.script_window.show()

 

def set_windows_app_id():

    app_id = "SQLInsertGenerator.1.0"

    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)

 

if __name__ == '__main__':

    app = QApplication(sys.argv)

    app_icon = QIcon(os.path.join(os.path.dirname(__file__), 'savicon.png'))

    app.setWindowIcon(app_icon)

    if sys.platform == "win32":

        set_windows_app_id()

    ex = SQLInsertGenerator()

    sys.exit(app.exec_())