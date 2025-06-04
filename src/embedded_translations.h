#ifndef EMBEDDED_TRANSLATIONS_H
#define EMBEDDED_TRANSLATIONS_H

namespace embedded_translations {

// Embedded English translation JSON
const char* EN_JSON = R"json(
{
  "log": {
    "error": {
      "no_excel_path": "No Excel file path has been set. Please use 'open_excel_and_list_sheets' first.",
      "failed_open_excel": "Failed to open Excel file: {0}",
      "failed_open_or_list": "Failed to open Excel file or list sheets: {0}",
      "missing_params": {
        "get_range": "Missing required parameters for get_sheet_range_content.",
        "create_xlsx": "Missing 'file_path' parameter for create_xlsx_file.",
        "set_range": "Missing required parameters for set_sheet_range_content."
      },
      "failed_select_sheet": "Failed to select sheet: {0}",
      "values_not_2d_array": "'values' parameter must be a 2D array for set_sheet_range_content.",
      "values_row_not_array": "Each row in 'values' must be an array for set_sheet_range_content.",
      "unsupported_cell_type": {
         "set_range": "Unsupported cell value type in 'values' array for set_sheet_range_content."
      },
      "failed_create_excel": "Failed to create Excel file: {0}",
      "failed_set_range": "Failed to set sheet range content for sheet: {0}"
    },
    "warn": {
       "unsupported_cell_type": {
          "get_range": "Unsupported cell type encountered at row {0}, col {1}: {2}",
          "standard": "Unsupported cell type encountered during standard processing: {0}"
       }
    },
    "info": {
      "opened_excel": "Successfully opened Excel file: {0}",
      "retrieved_range": "Successfully retrieved sheet range content from sheet: {0}",
      "created_excel": "Successfully created Excel file: {0}",
      "set_range": "Successfully set sheet range content for sheet: {0}",
      "server_start": "Starting MCP server at localhost:{0}",
      "server_stop_prompt": "Press Ctrl+C to stop the server"
    }
  },
  "exception": {
    "error": {
      "no_excel_path": "No Excel file path has been set. Please use 'open_excel_and_list_sheets' first.",
      "failed_open_excel": "Failed to open Excel file: {0}",
      "missing_param": {
         "file_path": "Missing 'file_path' parameter",
         "file_path_create": "Missing 'file_path' parameter"
      },
      "failed_open_or_list": "Failed to open Excel file or list sheets: {0}",
      "missing_params": {
         "get_range": "Missing required parameters for sheet range content.",
         "set_range": "Missing required parameters for setting sheet range content."
      },
      "failed_select_sheet": "Failed to select sheet: {0}",
      "failed_create_excel": "Failed to create Excel file: {0}",
      "values_not_2d_array": "'values' parameter must be a 2D array.",
      "values_row_not_array": "Each row in 'values' must be an array.",
      "unsupported_cell_type": {
         "set_range": "Unsupported cell value type in 'values' array."
      },
      "failed_set_range": "Failed to set sheet range content."
    }
  },
  "tool": {
    "open_excel": {
      "description": "Open an Excel file and list all sheet names. This tool will also set the current Excel file path for subsequent operations. RECOMMENDED TO RUN THIS TOOL FIRST BEFORE ANY OPERATION OR IF WANT TO CHANGE THE FILE TO MODIFY.",
      "param": {
        "file_path": "The full path to the Excel file"
      }
    },
    "get_range": {
      "description": "Get and output table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.",
      "param": {
        "sheet_name": "The name of the sheet to read from",
        "first_row": "The starting row number (1-indexed)",
        "first_column": "The starting column number (1-indexed)",
        "last_row": "The ending row number (1-indexed)",
        "last_column": "The ending column number (1-indexed)",
        "cell_with_coord": "Output non-empty cells with their respective coordinates, suitable for situations where the output area contains a large number of empty cells"
      }
    },
    "set_range": {
      "description": "Set table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.",
      "param": {
        "sheet_name": "The name of the sheet to write to",
        "first_row": "The starting row number (1-indexed)",
        "first_column": "The starting column number (1-indexed)",
        "values": "The 2D array of values to write to the range"
      }
    },
    "create_xlsx": {
      "description": "Create a new xlsx file with the given path. Automatically closes the Excel file after creation.",
      "param": {
        "file_path": "The ABSOLUTE path with which the file should create to"
      }
    }
  },
  "result": {
    "created_excel": "Excel file created successfully: {0}",
    "set_range": "Successfully set sheet range content.",
    "unsupported_type": "[Unsupported Type]",
    "invalid_address": "InvalidAddress"
  }
}
)json";

} // namespace embedded_translations

#endif // EMBEDDED_TRANSLATIONS_H