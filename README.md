<<<<<<< HEAD
# Excel MCP Server

This project provides a Flask-based server for managing and controlling Excel files, offering advanced functionalities such as creating files from templates, defining named ranges, inserting formulas, and handling VBA code.

## Features

- **File Creation (`/create_excel`):** Create new `.xlsx` or `.xlsm` files. Supports template-based creation, sheets with initial data, named ranges, and formulas.
  - For `.xlsx` files, it supports data validation.
  - For `.xlsm` files with VBA code, it supports sheet creation, data population, named ranges, formulas, and VBA code insertion. **Note: Data validation is currently skipped for these files due to `win32com.client` limitations.**
- **External Data Connection (`/add_external_data_connection`):** Add CSV/Text file external data connections to existing Excel files.
- **File Upload (`/upload`):** Upload Excel files to the server.
- **Sheet Management (`/sheets`, `/read_sheet`, `/write_sheet`):** List, read, and write data to Excel sheets.
- **Macro Management (`/read_macros`, `/execute_macro`, `/write_macro`):** Read, execute, and write VBA macros in `.xlsm` files.
- **Excel Analysis (`/analyze_excel`):** Analyze Excel files for sheet data and VBA macro interactions.

## Installation

1.  **Clone the repository:**
    ```bash
    git clone <repository_url>
    cd InsuranceProject
    ```
2.  **Create a virtual environment (recommended):**
    ```bash
    python -m venv venv
    ```
3.  **Activate the virtual environment:**
    -   **Windows:**
        ```bash
        .\venv\Scripts\activate
        ```
    -   **macOS/Linux:**
        ```bash
        source venv/bin/activate
        ```
4.  **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

## Running the Server

To start the Flask server, run:

```bash
python app.py
```

The server will typically run on `http://127.0.0.1:5001`.

## Usage Examples

Refer to the API endpoints and their expected JSON payloads for detailed usage. You can use `curl` or any HTTP client to interact with the server.

### Example: Creating an Excel file with sheets and formulas

```bash
curl -X POST -H "Content-Type: application/json" -d 
'{ 
  "filename": "my_new_file.xlsx",
  "sheets": [
    {
      "sheet_name": "Sheet1",
      "data": [
        ["Header1", "Header2"],
        [10, 20]
      ]
    }
  ],
  "formulas": [
    {
      "sheet_name": "Sheet1",
      "cell": "C2",
      "formula": "=A2+B2"
    }
  ]
}' http://127.0.0.1:5001/create_excel
```

### Example: Creating an XLSM file with VBA code

```bash
curl -X POST -H "Content-Type: application/json" -d 
'{ 
  "filename": "my_macro_file.xlsm",
  "sheets": [
    {
      "sheet_name": "Data",
      "data": [
        ["Value", 100]
      ]
    }
  ],
  "vba_code": [
    {
      "module_name": "Module1",
      "code": "Sub ShowMessage()\n  MsgBox \"Hello from VBA!\"
End Sub"
    }
  ]
}' http://127.0.0.1:5001/create_excel
```

## Limitations

-   **Data Validation in `.xlsm` files with VBA:** Due to persistent issues with `win32com.client` interaction with Excel's COM interface, data validation cannot be reliably applied to `.xlsm` files that contain VBA code during creation. This feature is currently skipped for such files.
-   **External Data Connection Types:** Currently supports only CSV/Text file connections. Other types (e.g., databases, web queries) are not yet implemented.

## Trust Access to the VBA project object model

For VBA macro functionality to work correctly, you might need to enable "Trust access to the VBA project object model" in Excel's Trust Center settings. This is typically found under `File > Options > Trust Center > Trust Center Settings... > Macro Settings`.
=======
# mcp-excel

This project provides a Flask-based server for managing andcontrolling Excel files, offering advanced functionalities such  as creating files from templates, defining named ranges,inserting formulas, and handling VBA code.
>>>>>>> 1c8ded511a3b41c4fe36cd9cd22da36eeb15052a
