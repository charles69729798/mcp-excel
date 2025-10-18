import os
from flask import Flask, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import openpyxl
import openpyxl.worksheet.worksheet # Added for sheet type checking
from openpyxl.worksheet.datavalidation import DataValidation # Import DataValidation
import win32com.client # Used for VBA interaction
import pythoncom # For CoInitialize/CoUninitialize
import re # For VBA code analysis
from win32com.client import constants # Explicitly import constants

# Configuration
UPLOAD_FOLDER = 'C:\\InsuranceProject\\uploads' # Directory to save uploaded Excel files
ALLOWED_EXTENSIONS = {'xlsx', 'xlsm'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- COM Initialization/Uninitialization for each request ---
@app.before_request
def before_request():
    pythoncom.CoInitialize()

@app.teardown_request
def teardown_request(exception):
    pythoncom.CoUninitialize()
# -----------------------------------------------------------

@app.route('/')
def index():
    return "Excel MCP Server is running!"

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        return jsonify({"message": "File uploaded successfully", "filename": filename, "filepath": filepath}), 200
    return jsonify({"error": "File type not allowed"}), 400

@app.route('/sheets/<filename>', methods=['GET'])
def list_sheets(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    try:
        workbook = openpyxl.load_workbook(filepath)
        sheet_names = workbook.sheetnames
        return jsonify({"filename": filename, "sheets": sheet_names}), 200
    except Exception as e:
        return jsonify({"error": f"Error listing sheets: {e}"}), 500

@app.route('/read_sheet/<filename>/<sheet_name>', methods=['GET'])
def read_sheet(filename, sheet_name):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    try:
        workbook = openpyxl.load_workbook(filepath, data_only=False) # Keep formulas
        if sheet_name not in workbook.sheetnames:
            return jsonify({"error": f"Sheet '{sheet_name}' not found"}), 404

        sheet = workbook[sheet_name]
        data = []
        for row in sheet.iter_rows():
            row_data = []
            for cell in row:
                if cell.data_type == 'f': # 'f' for formula
                    row_data.append(f"={cell.value}")
                else:
                    row_data.append(cell.value)
            data.append(row_data)
        return jsonify({"filename": filename, "sheet_name": sheet_name, "data": data}), 200
    except Exception as e:
        return jsonify({"error": f"Error reading sheet '{sheet_name}': {e}"}), 500

@app.route('/write_sheet/<filename>/<sheet_name>', methods=['POST'])
def write_sheet(filename, sheet_name):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    data_to_write = request.json.get('data')
    if not data_to_write or not isinstance(data_to_write, list):
        return jsonify({"error": "Invalid data format. Expected a list of lists."}), 400

    try:
        workbook = openpyxl.load_workbook(filepath)
        if sheet_name not in workbook.sheetnames:
            # Create new sheet if it doesn't exist
            sheet = workbook.create_sheet(sheet_name)
        else:
            sheet = workbook[sheet_name]
            # Clear existing content if overwriting
            sheet.delete_rows(1, sheet.max_row + 1) # Clear existing content

        for row_idx, row_data in enumerate(data_to_write):
            for col_idx, cell_value in enumerate(row_data):
                sheet.cell(row=row_idx + 1, column=col_idx + 1, value=cell_value)

        workbook.save(filepath)
        return jsonify({"message": f"Data written to sheet '{sheet_name}' in '{filename}' successfully"}), 200
    except Exception as e:
        return jsonify({"error": f"Error writing to sheet '{sheet_name}': {e}"}), 500

@app.route('/create_excel', methods=['POST'])
def create_excel():
    new_filename = request.json.get('filename')
    sheets_config = request.json.get('sheets', []) # Expect a list of sheet configurations
    named_ranges_config = request.json.get('named_ranges', []) # Expect a list of named range configurations
    template_filename = request.json.get('template_filename')
    formulas_config = request.json.get('formulas', []) # Expect a list of formula configurations

    if not new_filename:
        return jsonify({"error": "Filename is required"}), 400
    if not allowed_file(new_filename):
        return jsonify({"error": "Invalid file extension. Only .xlsx or .xlsm allowed."}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(new_filename))

    try:
        # Handle VBA code insertion for .xlsm files using win32com.client
        vba_code_configs = request.json.get('vba_code', [])
        if vba_code_configs and new_filename.lower().endswith('.xlsm'):
            excel = None
            wb = None
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Add() # Create a new workbook

                # Add sheets and data using win32com.client
                created_sheet_names = []
                if not sheets_config:
                    # If no sheets are specified, ensure at least one default sheet exists
                    if wb.Sheets.Count == 0:
                        wb.Sheets.Add().Name = "Sheet1"
                    else:
                        wb.Sheets(1).Name = "Sheet1"
                    created_sheet_names.append("Sheet1")
                else:
                    for sheet_config in sheets_config:
                        sheet_name = sheet_config.get('sheet_name', 'Sheet')
                        initial_data = sheet_config.get('data', [])

                        # Check if sheet already exists
                        sheet_exists = False
                        for s in wb.Sheets:
                            if s.Name == sheet_name:
                                sheet_exists = True
                                break

                        if sheet_exists:
                            sheet = wb.Sheets(sheet_name)
                            sheet.Cells.ClearContents() # Clear existing content
                        else:
                            sheet = wb.Sheets.Add()
                            sheet.Name = sheet_name

                        created_sheet_names.append(sheet_name)

                        for row_idx, row_data in enumerate(initial_data):
                            for col_idx, cell_value in enumerate(row_data):
                                sheet.Cells(row_idx + 1, col_idx + 1).Value = cell_value

                # Delete any default sheets that were not explicitly created or renamed
                for i in range(wb.Sheets.Count, 0, -1):
                    sheet = wb.Sheets(i)
                    if sheet.Name.startswith("Sheet") and sheet.Name not in created_sheet_names:
                        if wb.Sheets.Count > 1: # Only delete if there's more than one sheet left
                            sheet.Delete()
                        else:
                            # If it's the last sheet and it's a default one, rename it if sheets_config was empty
                            if not sheets_config and sheet.Name == "Sheet1":
                                pass # Already handled by renaming
                            else:
                                # If it's the last sheet and it's a default one, and sheets_config was not empty,
                                # it means all user-defined sheets were created, and this default one is extra.
                                # We need to ensure at least one sheet remains, so we can't delete it.
                                # This scenario should ideally not happen if user-defined sheets are created first.
                                pass # Keep it for now to avoid error, or handle more gracefully

                # Ensure at least one sheet exists after all operations
                if wb.Sheets.Count == 0:
                    wb.Sheets.Add().Name = "Sheet1"

                # Apply named ranges using win32com.client
                named_ranges_config = request.json.get('named_ranges', [])
                for named_range_config in named_ranges_config:
                    name = named_range_config.get('name')
                    sheet_name = named_range_config.get('sheet_name')
                    range_str = named_range_config.get('range')

                    if not all([name, sheet_name, range_str]):
                        return jsonify({"error": "Invalid named range configuration. 'name', 'sheet_name', and 'range' are required."}), 400

                    try:
                        wb.Names.Add(Name=name, RefersTo=f"='{sheet_name}'!{range_str}")
                    except Exception as e:
                        return jsonify({"error": f"Error adding named range '{name}': {e}"}), 500

                # Apply formulas using win32com.client
                formulas_config = request.json.get('formulas', [])
                for formula_config in formulas_config:
                    sheet_name = formula_config.get('sheet_name')
                    cell = formula_config.get('cell')
                    formula = formula_config.get('formula')

                    if not all([sheet_name, cell, formula]):
                        return jsonify({"error": "Invalid formula configuration. 'sheet_name', 'cell', and 'formula' are required."}), 400

                    try:
                        wb.Sheets(sheet_name).Range(cell).Formula = formula
                    except Exception as e:
                        return jsonify({"error": f"Error applying formula '{formula}' to cell '{cell}': {e}"}), 500

                # Apply data validations using win32com.client
                # Due to persistent issues with win32com.client and data validation operators,
                # this feature is currently limited for .xlsm files with VBA code.
                # Data validation will be skipped for these files.

                # Insert VBA code
                for vba_config in vba_code_configs:
                    module_name = vba_config.get('module_name')
                    code = vba_config.get('code')

                    if not all([module_name, code]):
                        return jsonify({"error": "Invalid VBA code configuration. 'module_name' and 'code' are required."}), 400

                    try:
                        module = wb.VBProject.VBComponents(module_name)
                    except Exception:
                        module = wb.VBProject.VBComponents.Add(1) # 1 for standard module
                        module.Name = module_name

                    module.CodeModule.AddFromString(code)

                # Save the workbook as .xlsm
                wb.SaveAs(filepath, FileFormat=52) # FileFormat=52 for xlOpenXMLWorkbookMacroEnabled
                return jsonify({"message": f"New Excel file '{new_filename}' created successfully with VBA", "filepath": filepath}), 200
            except Exception as e:
                return jsonify({"error": f"Error creating Excel file with VBA: {e}. Ensure Excel is installed, file is not open, and 'Trust access to the VBA project object model' is enabled in Excel's Trust Center."}), 500
            finally:
                if wb:
                    wb.Close(False) # Don't save changes again, already saved
                if excel:
                    excel.Quit()
        else:
            # Existing openpyxl logic for .xlsx files or .xlsm without VBA
            if template_filename:
                template_filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(template_filename))
                if not os.path.exists(template_filepath):
                    return jsonify({"error": f"Template file '{template_filename}' not found"}), 404
                workbook = openpyxl.load_workbook(template_filepath)
            else:
                workbook = openpyxl.Workbook()
                # Remove default sheet created by openpyxl
                if workbook.active:
                    workbook.remove(workbook.active)

            if not sheets_config and not template_filename:
                # If no sheets are specified and no template, create a default Sheet1
                workbook.create_sheet("Sheet1")
            else:
                for sheet_config in sheets_config:
                    sheet_name = sheet_config.get('sheet_name', 'Sheet')
                    initial_data = sheet_config.get('data', [])

                    # If sheet already exists, clear it before writing new data
                    if sheet_name in workbook.sheetnames:
                        sheet = workbook[sheet_name]
                        # Clear existing content
                        for row in sheet.iter_rows():
                            for cell in row:
                                cell.value = None
                    else:
                        sheet = workbook.create_sheet(sheet_name)

                    for row_idx, row_data in enumerate(initial_data):
                        for col_idx, cell_value in enumerate(row_data):
                            sheet.cell(row=row_idx + 1, column=col_idx + 1, value=cell_value)

            # Apply named ranges
            named_ranges_config = request.json.get('named_ranges', [])
            for named_range_config in named_range_config:
                name = named_range_config.get('name')
                sheet_name = named_range_config.get('sheet_name')
                range_str = named_range_config.get('range')

                if not all([name, sheet_name, range_str]):
                    return jsonify({"error": "Invalid named range configuration. 'name', 'sheet_name', and 'range' are required."}), 400

                if sheet_name not in workbook.sheetnames:
                    return jsonify({"error": f"Sheet '{sheet_name}' for named range '{name}' not found"}), 404

                # Define the named range
                workbook.defined_names.add(name, workbook[sheet_name], range_str)

            # Apply formulas
            formulas_config = request.json.get('formulas', [])
            for formula_config in formulas_config:
                sheet_name = formula_config.get('sheet_name')
                cell = formula_config.get('cell')
                formula = formula_config.get('formula')

                if not all([sheet_name, cell, formula]):
                    return jsonify({"error": "Invalid formula configuration. 'sheet_name', 'cell', and 'formula' are required."}), 400

                if sheet_name not in workbook.sheetnames:
                    return jsonify({"error": f"Sheet '{sheet_name}' for formula '{formula}' not found"}), 404

                workbook[sheet_name][cell] = formula

            # Apply data validations
            for dv_config in request.json.get('data_validations', []):
                sheet_name = dv_config.get('sheet_name')
                cell_range = dv_config.get('cell_range')
                dv_type = dv_config.get('type')
                operator = dv_config.get('operator')
                formula1 = dv_config.get('formula1') # Use formula1 for list type
                value1 = dv_config.get('value1')
                value2 = dv_config.get('value2')

                if not all([sheet_name, cell_range, dv_type]):
                    return jsonify({"error": "Invalid data validation configuration. 'sheet_name', 'cell_range', and 'type' are required."}), 400

                if dv_type == 'list':
                    if not formula1:
                        return jsonify({"error": "Invalid data validation configuration for list type. 'formula1' is required."}), 400
                    dv = DataValidation(type=dv_type, formula1=formula1)
                else:
                    if not all([operator, value1]):
                        return jsonify({"error": "Invalid data validation configuration. 'operator' and 'value1' are required for non-list types."}), 400
                    dv = DataValidation(type=dv_type, operator=operator, formula1=value1, formula2=value2)

                if sheet_name not in workbook.sheetnames:
                    return jsonify({"error": f"Sheet '{sheet_name}' for data validation not found"}), 404

                workbook[sheet_name].add_data_validation(dv)
                dv.add(cell_range)

            workbook.save(filepath)
            return jsonify({"message": f"New Excel file '{new_filename}' created successfully", "filepath": filepath}), 200
    except Exception as e:
        return jsonify({"error": f"Error creating Excel file: {e}"}), 500

def unicode_escape_korean(text):
    # This function is designed to convert a string to its VBA Unicode escaped form
    # for matching against VBA code that might have escaped Korean characters.
    # Example: "규정체크" -> "\\uaddc\\uc815\\uccb4\\ud06c"
    return ''.join([f'\\u{ord(char):04x}' for char in text])

@app.route('/analyze_excel/<filename>', methods=['GET'])
def analyze_excel(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    analysis_report = {
        "filename": filename,
        "sheets": {},
        "vba_modules": {},
        "vba_sheet_interactions": {}
    }

    try:
        # --- Sheet Analysis ---
        workbook = openpyxl.load_workbook(filepath, data_only=False)
        for sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            # Skip chart sheets as they don't have iter_rows()
            if not isinstance(sheet, openpyxl.worksheet.worksheet.Worksheet):
                continue

            sheet_data = []
            for r_idx, row in enumerate(sheet.iter_rows()):
                if r_idx >= 10: break # Limit to first 10 rows for sample
                row_data = []
                for c_idx, cell in enumerate(row):
                    if c_idx >= 10: break # Limit to first 10 columns for sample
                    if cell.data_type == 'f':
                        row_data.append(f"={cell.value}")
                    else:
                        row_data.append(cell.value)
                sheet_data.append(row_data)
            analysis_report["sheets"][sheet_name] = {
                "sample_data": sheet_data,
                "max_row": sheet.max_row,
                "max_column": sheet.max_column
            }

        # --- VBA Macro Analysis ---
        if filename.lower().endswith('.xlsm'):
            excel = None
            wb = None
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                wb = excel.Workbooks.Open(filepath, ReadOnly=True)

                vba_code = {}
                for component in wb.VBProject.VBComponents:
                    if component.Type in [1, 2, 3]: # Standard, Class, Object modules
                        if component.CodeModule.CountOfLines > 0:
                            vba_code[component.Name] = wb.VBProject.VBComponents(component.Name).CodeModule.Lines(1, wb.VBProject.VBComponents(component.Name).CodeModule.CountOfLines)
                analysis_report["vba_modules"] = vba_code
            except Exception as e:
                analysis_report["vba_modules_error"] = f"Error reading VBA macros: {e}"
            finally:
                if wb:
                    wb.Close(False)
                if excel:
                    excel.Quit()

            # --- VBA-Sheet Interaction Analysis ---
            if analysis_report["vba_modules"]:
                target_sheet_names = list(analysis_report["sheets"].keys())
                interactions = {name: {"reads": [], "writes": [], "other": []} for name in target_sheet_names}

                for module_name, code in analysis_report["vba_modules"].items():
                    for sheet_name in target_sheet_names:
                        # The `code` variable contains the VBA code as a string.
                        # If the sheet name is "INPUT", the code might contain Worksheets("INPUT")
                        # If the sheet name is "규정체크", the code might contain Worksheets("규정체크")
                        # or Worksheets("\\uaddc\\uc815\\uccb4\\ud06c") depending on how win32com returns it.

                        # Let's assume win32com returns the actual sheet name for English,
                        # and Unicode escaped for Korean.

                        # So, we need to construct a regex that matches either the direct sheet name
                        # or its Unicode escaped form.

                        # The VBA code string from read_macros has escaped quotes: Worksheets(\"SheetName\")
                        # So the regex needs to match Worksheets\\(\\\"[sheet_name_representation]\\\"\\)

                        # Let's try to match the literal sheet_name directly, and also its unicode escaped form.
                        # The `code` string itself contains `\"` for quotes.

                        # So, the regex should be:
                        # `Worksheets\\(\\\"` + re.escape(sheet_name) + `\\\"\)`  (for direct names like INPUT)
                        # OR
                        # `Worksheets\\(\\\"` + unicode_escape_korean(sheet_name) + `\\\"\)` (for unicode escaped names like 규정체크)

                        # Let's combine these two possibilities into one regex.

                        sheet_name_pattern_escaped = re.escape(sheet_name)
                        sheet_name_pattern_unicode_escaped = unicode_escape_korean(sheet_name)

                        # Create a combined pattern that matches either the direct escaped name or the unicode escaped name
                        combined_sheet_name_pattern = f"(?:{sheet_name_pattern_escaped}|{sheet_name_pattern_unicode_escaped})"

                        # Now, construct the full regex patterns using this combined pattern
                        # We need to match Worksheets("...") or Worksheets('...')
                        # The VBA code from read_macros uses \" for quotes.

                        read_pattern = re.compile(r'Worksheets\\(\\\"(?:' + combined_sheet_name_pattern + r')\\\"\\)\\.(Range|Cells)\\(.*?\\)', re.IGNORECASE)
                        write_pattern = re.compile(r'Worksheets\\(\\\"(?:' + combined_sheet_name_pattern + r')\\\"\\)\\.(Range|Cells)\\(.*?\\)\\s*=\\', re.IGNORECASE)
                        select_pattern = re.compile(r'Worksheets\\(\\\"(?:' + combined_sheet_name_pattern + r')\\\"\\)\\.Select', re.IGNORECASE)
                        activate_pattern = re.compile(r'Worksheets\\(\\\"(?:' + combined_sheet_name_pattern + r')\\\"\\)\\.Activate', re.IGNORECASE)

                        for line in code.splitlines():
                            if write_pattern.search(line):
                                interactions[sheet_name]["writes"].append(f"Module '{module_name}': {line.strip()}")
                            elif read_pattern.search(line):
                                interactions[sheet_name]["reads"].append(f"Module '{module_name}': {line.strip()}")
                            elif select_pattern.search(line) or activate_pattern.search(line):
                                interactions[sheet_name]["other"].append(f"Module '{module_name}': {line.strip()}")
                analysis_report["vba_sheet_interactions"] = interactions

    except Exception as e:
        return jsonify({"error": f"Error analyzing Excel file: {e}"}), 500

    return jsonify(analysis_report), 200

@app.route('/read_macros/<filename>', methods=['GET'])
def read_macros(filename):
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    if not filename.lower().endswith('.xlsm'):
        return jsonify({"error": "File is not a macro-enabled Excel workbook (.xlsm)"}), 400

    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False # Keep Excel hidden
        # Ensure the workbook is not already open in Excel
        # This might require closing it if it's open, or handling the error
        workbook = excel.Workbooks.Open(filepath, ReadOnly=True) # Open in read-only mode

        vba_code = {}
        for component in workbook.VBProject.VBComponents:
            # Type 1: Standard module, Type 2: Class module, Type 3: Object module (Sheet, ThisWorkbook)
            if component.Type in [1, 2, 3]:
                # Only extract if there's actual code
                if component.CodeModule.CountOfLines > 0:
                    vba_code[component.Name] = workbook.VBProject.VBComponents(component.Name).CodeModule.Lines(1, workbook.VBProject.VBComponents(component.Name).CodeModule.CountOfLines)

        return jsonify({"filename": filename, "vba_modules": vba_code}), 200
    except Exception as e:
        return jsonify({"error": f"Error reading macros from \'{filename}\': {e}. Ensure Excel is installed and the file is not open."}), 500
    finally:
        if workbook:
            workbook.Close(False) # Don't save changes
        if excel: # Only quit if excel was successfully dispatched
            excel.Quit()

@app.route('/execute_macro', methods=['POST']) # Modified to use query parameters
def execute_macro():
    filename = request.args.get('filename')
    macro_name = request.args.get('macro_name')

    if not filename or not macro_name:
        return jsonify({"error": "Filename and macro_name are required query parameters."}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    if not filename.lower().endswith('.xlsm'):
        return jsonify({"error": "File is not a macro-enabled Excel workbook (.xlsm)"}), 400

    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False # Keep Excel hidden
        workbook = excel.Workbooks.Open(filepath, ReadOnly=False) # Open in read-write mode to execute macro

        # Execute the macro
        excel.Application.Run(macro_name)

        workbook.Save() # Save changes made by macro
        return jsonify({"message": f"Macro '{macro_name}' executed successfully in '{filename}'"}), 200
    except Exception as e:
        return jsonify({"error": f"Error executing macro '{macro_name}' in '{filename}': {e}. Ensure Excel is installed and the file is not open."}), 500
    finally:
        if workbook:
            workbook.Close(True) # Save and close
        if excel:
            excel.Quit()

@app.route('/write_macro', methods=['POST']) # Modified to use query parameters
def write_macro():
    filename = request.args.get('filename')
    module_name = request.args.get('module_name')
    macro_name = request.args.get('macro_name')

    if not filename or not module_name or not macro_name:
        return jsonify({"error": "Filename, module_name, and macro_name are required query parameters."}), 400

    new_code = request.json.get('code')
    if not new_code:
        return jsonify({"error": "Macro code is required in the request body."}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    if not filename.lower().endswith('.xlsm'):
        return jsonify({"error": "File is not a macro-enabled Excel workbook (.xlsm)"}), 400

    excel = None
    workbook = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        workbook = excel.Workbooks.Open(filepath, ReadOnly=False)

        vba_project = workbook.VBProject
        vb_components = vba_project.VBComponents # Corrected line

        if module_name not in [comp.Name for comp in vb_components]:
            return jsonify({"error": f"Module '{module_name}' not found in '{filename}'"}), 404

        module = vb_components(module_name)

        # Find the existing macro and replace its code
        # This is a simplified approach. A more robust solution would parse VBA.
        # For now, we assume we are replacing the entire content of a macro.

        # Get current code
        current_code = module.CodeModule.Lines(1, module.CodeModule.CountOfLines)

        # Find start and end of the target macro
        start_line = -1
        end_line = -1
        lines = current_code.splitlines()
        for i, line in enumerate(lines):
            if f"Sub {macro_name}" in line or f"Function {macro_name}" in line:
                start_line = i + 1 # VBA lines are 1-indexed
            if start_line != -1 and ("End Sub" in line or "End Function" in line):
                end_line = i + 1
                break

        if start_line == -1:
            return jsonify({"error": f"Macro '{macro_name}' not found in module '{module_name}'"}), 404

        # Delete old code and insert new code
        module.CodeModule.DeleteLines(start_line, end_line - start_line + 1)
        module.CodeModule.InsertLines(start_line, new_code)

        workbook.Save()
        return jsonify({"message": f"Macro '{macro_name}' in module '{module_name}' of '{filename}' updated successfully"}), 200
    except Exception as e:
        return jsonify({"error": f"Error writing macro '{macro_name}' in '{filename}': {e}. Ensure Excel is installed, file is not open, and 'Trust access to the VBA project object model' is enabled in Excel's Trust Center."}), 500
    finally:
        if workbook:
            workbook.Close(True) # Save and close
        if excel:
            excel.Quit()


@app.route('/add_external_data_connection', methods=['POST'])
def add_external_data_connection():
    filename = request.json.get('filename')
    sheet_name = request.json.get('sheet_name')
    source_path = request.json.get('source_path')
    destination_cell = request.json.get('destination_cell')
    delimiter = request.json.get('delimiter', ',') # Default to comma
    has_headers = request.json.get('has_headers', True) # Default to True

    if not all([filename, sheet_name, source_path, destination_cell]):
        return jsonify({"error": "Filename, sheet_name, source_path, and destination_cell are required."}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(filepath, ReadOnly=False)

        if sheet_name not in [s.Name for s in wb.Sheets]:
            return jsonify({"error": f"Sheet '{sheet_name}' not found in '{filename}'"}), 404

        sheet = wb.Sheets(sheet_name)

        # Add QueryTable for CSV/Text file
        # TextFilePlatform = 850 (Western European DOS) or 1252 (Windows ANSI)
        # TextFileStartRow = 1
        # TextFileParseType = 1 (xlDelimited)
        # TextFileTextQualifier = 1 (xlTextQualifierDoubleQuote)
        # TextFileConsecutiveDelimiter = False
        # TextFileTabDelimiter = False
        # TextFileCommaDelimiter = True (if delimiter is comma)
        # TextFileSemicolonDelimiter = False
        # TextFileSpaceDelimiter = False
        # TextFileOtherDelimiter = delimiter (if not comma, tab, etc.)

        # Determine delimiter type
        if delimiter == ',':
            TextFileCommaDelimiter = True
            TextFileOtherDelimiter = False
        elif delimiter == '\t':
            TextFileTabDelimiter = True
            TextFileOtherDelimiter = False
        else:
            TextFileCommaDelimiter = False
            TextFileTabDelimiter = False
            TextFileOtherDelimiter = True

        QueryTable = sheet.QueryTables.Add(
            Connection=f"TEXT;{source_path}",
            Destination=sheet.Range(destination_cell)
        )
        QueryTable.Name = f"CSV_Import_{sheet_name}"
        QueryTable.FieldNames = has_headers
        QueryTable.RowNumbers = False
        QueryTable.FillAdjacentFormulas = False
        QueryTable.PreserveFormatting = True
        QueryTable.RefreshOnFileOpen = False
        QueryTable.RefreshStyle = 1 # xlInsertDeleteCells
        QueryTable.SavePassword = False
        QueryTable.SaveData = True
        QueryTable.AdjustColumnWidth = True
        QueryTable.RefreshPeriod = 0
        QueryTable.TextFilePromptOnRefresh = False
        QueryTable.TextFilePlatform = 65001 # 65001 for UTF-8
        QueryTable.TextFileStartRow = 1
        QueryTable.TextFileParseType = 1 # xlDelimited
        QueryTable.TextFileTextQualifier = 1 # xlTextQualifierDoubleQuote
        QueryTable.TextFileConsecutiveDelimiter = False
        QueryTable.TextFileTabDelimiter = (delimiter == '\t')
        QueryTable.TextFileSemicolonDelimiter = False
        QueryTable.TextFileCommaDelimiter = (delimiter == ',')
        QueryTable.TextFileSpaceDelimiter = False
        QueryTable.TextFileOtherDelimiter = (delimiter not in [',', '\t'])
        if QueryTable.TextFileOtherDelimiter:
            QueryTable.TextFileOtherDelimiter = delimiter

        QueryTable.Refresh(BackgroundQuery=False)

        wb.Save()
        return jsonify({"message": f"External data connection added to '{filename}' successfully", "filepath": filepath}), 200
    except Exception as e:
        return jsonify({"error": f"Error adding external data connection: {e}. Ensure Excel is installed and the file is not open."}), 500
    finally:
        if wb:
            wb.Close(True) # Save and close
        if excel:
            excel.Quit()

@app.route('/add_data_validation_to_existing_file', methods=['POST'])
def add_data_validation_to_existing_file():
    filename = request.json.get('filename')
    data_validations = request.json.get('data_validations', [])

    if not filename:
        return jsonify({"error": "Filename is required"}), 400
    if not data_validations:
        return jsonify({"error": "Data validations are required"}), 400

    filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(filename))
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    excel = None
    wb = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(filepath, ReadOnly=False)

        for dv_config in data_validations:
            sheet_name = dv_config.get('sheet_name')
            cell_range = dv_config.get('cell_range')
            dv_type = dv_config.get('type')
            operator = dv_config.get('operator')
            formula1 = dv_config.get('formula1')
            value1 = dv_config.get('value1')
            value2 = dv_config.get('value2')

            if not all([sheet_name, cell_range, dv_type]):
                return jsonify({"error": "Invalid data validation configuration. 'sheet_name', 'cell_range', and 'type' are required."}), 400

            if sheet_name not in [s.Name for s in wb.Sheets]:
                return jsonify({"error": f"Sheet '{sheet_name}' not found in '{filename}'"}), 404

            sheet = wb.Sheets(sheet_name)

            try:
                # Clear existing data validation for the range
                sheet.Range(cell_range).Validation.Delete()

                if dv_type == 'list':
                    if not formula1:
                        return jsonify({"error": "Invalid data validation configuration for list type. 'formula1' is required."}), 400
                    sheet.Range(cell_range).Validation.Add(
                        Type=3, # xlValidateList
                        AlertStyle=1, # xlValidAlertStop
                        Formula1=formula1
                    )
                else:
                    if not all([operator, value1]):
                        return jsonify({"error": "Invalid data validation configuration. 'operator' and 'value1' are required for non-list types."}), 400

                    # Map operators to win32com constants
                    xl_operator = None
                    if operator == 'between': xl_operator = constants.xlBetween
                    elif operator == 'notBetween': xl_operator = constants.xlNotBetween
                    elif operator == 'equal': xl_operator = constants.xlEqual
                    elif operator == 'notEqual': xl_operator = constants.xlNotEqual
                    elif operator == 'greaterThan': xl_operator = constants.xlGreaterThan
                    elif operator == 'lessThan': xl_operator = constants.xlLessThan
                    elif operator == 'greaterThanOrEqual': xl_operator = constants.xlGreaterEqual
                    elif operator == 'lessThanOrEqual': xl_operator = constants.xlLessEqual

                    if xl_operator is None:
                        return jsonify({"error": f"Unsupported operator for data validation: {operator}"}), 400

                    # Map dv_type to win32com constants
                    xl_validate_type = None
                    if dv_type == 'whole': xl_validate_type = constants.xlValidateWholeNumber
                    elif dv_type == 'decimal': xl_validate_type = constants.xlValidateDecimal

                    if xl_validate_type is None:
                        return jsonify({"error": f"Unsupported data validation type: {dv_type}"}), 400

                    sheet.Range(cell_range).Validation.Add(
                        Type=xl_validate_type,
                        AlertStyle=constants.xlValidAlertStop,
                        Operator=xl_operator,
                        Formula1=value1,
                        Formula2=value2
                    )
                sheet.Range(cell_range).Validation.IgnoreBlank = True
                sheet.Range(cell_range).Validation.InCellDropdown = True

            except Exception as e:
                return jsonify({"error": f"Error applying data validation to range '{cell_range}': {e}"}), 500

        wb.Save()
        return jsonify({"message": f"Data validation added to '{filename}' successfully", "filepath": filepath}), 200
    except Exception as e:
        return jsonify({"error": f"Error processing data validation: {e}. Ensure Excel is installed and the file is not open."}), 500
    finally:
        if wb:
            wb.Close(True) # Save and close
        if excel:
            excel.Quit()

@app.route('/routes')
def list_routes():
    output = []
    for rule in app.url_map.iter_rules():
        methods = ','.join(rule.methods)
        output.append(f"{rule.endpoint}: {rule.rule} ({methods})")
    return jsonify({"routes": output})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)