import argparse # Used for parsing command line arguments
import sys # Used for accessing command line arguments
import openpyxl # Used for reading from the XLSX file
import json # Used to parse metadata in cell comments and file description, because it doesn't rely on whitespace like YAML
import os, platform, subprocess # Used for opening the generated XLSX file
import send2trash # Used for deleting the XLSX file, so the user can get it back if this was done by accident
from .constants import AUTOGEN_PREFIX, DEFAULT_METADATA, get_default_xlsx_path, get_default_sql_path

def xlsx_to_sql(arglist=sys.argv[1:]):
  # Parse command line arguments
  parser = argparse.ArgumentParser()
  parser.add_argument("-i", "--input", help="Path to the XLSX script to convert", default=get_default_xlsx_path())
  parser.add_argument("-o", "--output", help="Path to the SQL file to create", default=get_default_sql_path())

  parser.add_argument("-ltn", "--log-table-names", help="Log names of detected tables", action="store_true")
  parser.add_argument("-lr", "--log-records", help="Log detected records", action="store_true")
  parser.add_argument("-lf", "--log-fields", help="Log names of detected fields in tables", action="store_true")

  parser.add_argument("-sw", "--suppress-warnings", help="Do not display warnings", action="store_true")

  output_options = parser.add_argument_group("XLSX file options")
  parser.add_argument("-d", "--delete-xlsx", help="Do not delete the XLSX file", action="store_true")
  parser.add_argument("-p", "--preserve-xlsx", help="Do not delete the XLSX file", action="store_false", dest='delete_xlsx')
  # By default, delete the file at the end, so that there aren't 2 versions of the data
  parser.set_defaults(delete_xlsx=True)

  output_options = parser.add_argument_group("Output options")
  output_options.add_argument("-of", "--open-file", help="Automatically open the generated SQL file", action="store_true")
  output_options.add_argument("-nof", "--no-open-file", help="Do not open the generated SQL file", action="store_false", dest='open_file')
  # By default, don't open the file at the end: assume the user will manually do that when using the file in CCI
  parser.set_defaults(open_file=False)

  args = parser.parse_args(arglist)

  # Output options
  LOG_TABLE_NAMES = args.log_table_names
  LOG_RECORDS = args.log_records
  LOG_FIELDS = args.log_fields

  DISPLAY_WARNINGS = not args.suppress_warnings

  OPEN_FILE = args.open_file
  DELETE_XLSX = args.delete_xlsx

  INPUT_FILE = args.input
  OUTPUT_FILE = args.output

  # Logging functions
  def log_table_name(text):
    if LOG_TABLE_NAMES: print(text)
  def log_record(text):
    if LOG_RECORDS: print(text)
  def log_field(text):
    if LOG_FIELDS: print(text)
  def log_warning(text):
    if DISPLAY_WARNINGS: print(text)

  # Escape any single quotes in a string
  def escape_string_sql(string):
    if not isinstance(string, str):
      return string
    
    esc = string.replace("'", "''")
    return esc

  class XLSXParseError(Exception): pass

  # Load the XLSX file
  workbook = openpyxl.load_workbook(INPUT_FILE)

  # Get the document's description, where we're storing metadata about the tables
  desc_text = workbook.properties.description
  prefix = desc_text[:len(AUTOGEN_PREFIX)]
  suffix = desc_text[len(AUTOGEN_PREFIX):]

  if prefix != AUTOGEN_PREFIX:
    raise XLSXParseError(f"Comment verification failed: '{desc_text}'")
  
  metadata = json.loads(suffix)
  table_names = metadata["table_names"]

  # Variable to hold contents of the SQL script
  sql_script = "BEGIN TRANSACTION;\n"

  # For each worksheet, create the corresponding table in the SQL script
  for sheet in workbook.worksheets:
    sheet_name = sheet.title
    table_name = table_names[sheet_name]
    log_table_name(f"Extracting table: {table_name} ({sheet_name})")

    # Collect list of fields in this table
    fields = []
    mapping_fields = [] # Fields which are not the primary key
    field_metadata = []
    column = 1
    cell = sheet.cell(row=1, column=column)
    while cell.value:
      field_name = cell.value
      fields.append(field_name)

      # Extract field metadata
      if cell.comment:
        comment_text = cell.comment.text
        prefix = comment_text[:len(AUTOGEN_PREFIX)]
        suffix = comment_text[len(AUTOGEN_PREFIX):]

        if prefix != AUTOGEN_PREFIX:
          raise XLSXParseError(f"Comment verification failed: '{comment_text}'")

        metadata = json.loads(suffix)
        field_metadata.append(metadata)
      else:
        log_warning(f"WARNING: No comment found for field {field_name}. Using default type: VARCHAR(255), not primary key.")
        metadata = DEFAULT_METADATA
        field_metadata.append(metadata)

      column += 1
      cell = sheet.cell(row=1, column=column)
      log_field(f"Detected field: {field_name} with metadata {metadata}")
    log_field(f"Detected {len(fields)} fields. The final, empty cell's value's type was: {str(type(cell.value))}")

    # Create the table
    sql_script += f'CREATE TABLE "{escape_string_sql(table_name)}" (\n'

    primary_key_field = None

    # Add the fields
    for column, field_name in enumerate(fields):
      metadata = field_metadata[column]
      if metadata["pk"]:
        if primary_key_field is not None:
          raise XLSXParseError(f"Table {table_name}: multiple primary key fields detected!")
        
        primary_key_field = field_name

      # Primary key field doesn't get quotes around its name
      if primary_key_field == field_name:
        sql_script += f'\t{escape_string_sql(field_name)}'
      else:
        sql_script += f'\t"{escape_string_sql(field_name)}"'
        mapping_fields.append(field_name)

      sql_script += f' {metadata["type"]}'
      if metadata["not_null"]: sql_script += " NOT NULL"
      sql_script += ",\n"

    # Add the primary key
    if primary_key_field is not None:
      sql_script += f'\tPRIMARY KEY ({escape_string_sql(primary_key_field)})\n'
    else:
      raise XLSXParseError(f"Table {table_name}: no primary key field detected!")

    # End the table with close-paren
    sql_script += ");\n"

    # Insert records into the table
    for row_index, row_obj in enumerate(sheet.rows):
      if row_index == 0: continue # Skip the header row

      # Skip empty rows
      if not any(cell.value for cell in row_obj):
        continue

      values = []
      for column_index, field in enumerate(fields):
        cell = sheet.cell(row=row_index+1, column=column_index+1) # Remember, 1-based indexing!
        value = cell.value

        # CCI uses an empty string to indicate NULL, but OpenPyXL uses None
        if value is None: value = ''

        # CCI stores data, including dates, as strings. For dates already present in the original SQL file, this is handled correctly.
        # But for dates in fields that were previously empty, we need to convert them to strings.
        if cell.is_date:
          value = cell.value.strftime("%Y-%m-%d")

        target_type = field_metadata[column_index]["type"]
        if target_type == "INTEGER":
          value = int(value)
        else:
          value = str(value) # Ensure the value is a string, since Excel likes to auto-format booleans and numbers

        values.append(value)

      sql_script += f'INSERT INTO "{escape_string_sql(table_name)}" VALUES('
      for column, value in enumerate(values):
        if column > 0: sql_script += ","
        if isinstance(value, str): sql_script += f"'{escape_string_sql(value)}'"
        else: sql_script += f"{escape_string_sql(value)}"
      sql_script += ");\n"

      log_record(f"Record in table {escape_string_sql(table_name)}: {escape_string_sql(values)}")

  # End the SQL file
  sql_script += "COMMIT;\n"

  # Save the SQL file
  with open(OUTPUT_FILE, 'w') as f:
    f.write(sql_script)

  # Delete the XLSX file
  if DELETE_XLSX:
    send2trash.send2trash(INPUT_FILE)

  # Open the created SQL file
  if OPEN_FILE:
    if platform.system() == "Windows":
      os.startfile(OUTPUT_FILE)
    else:
      # subprocess.call() is safer than os.system(), as it prevents injection attacks. Of course, if an attacker can run this script, they probably have command-line access already, but why take the risk?
      subprocess.call(('open', OUTPUT_FILE))

if __name__ == "__main__":
  xlsx_to_sql()