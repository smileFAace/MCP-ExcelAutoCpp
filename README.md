Language：[中文](README_zh-CN.md) | [English](README.md)
# C++ Excel Automation MCP Server

This is a C++ based Excel Automation MCP (Model Context Protocol) server project. It utilizes the OpenXLSX library for Excel file operations and aims to provide an intelligent set of Excel manipulation tools through the MCP protocol, leveraging LLM capabilities.

## Table of Contents

*   [Features](#features)
*   [Building the Project](#building-the-project)
*   [Usage](#usage)


## Features

*   **Simple and Easy-to-Use MCP Interface**: Provides standardized MCP tools, making it convenient for clients (like AI assistants) to invoke Excel automation functions.
*   **Single-File Deployment**: Compiles into a single executable file for easy deployment and execution.
*   **Customizable Multi-language Support**: Easily add or modify interface languages via JSON files.

## Server Capabilities

This MCP server provides the following tools, which can be invoked by an LLM to interact with xlsx spreadsheet files:

*   **`open_excel_and_list_sheets`**:
    *   Description: Open an Excel file and list all sheet names. This tool will also set the current Excel file path for subsequent operations. RECOMMENDED TO RUN THIS TOOL FIRST BEFORE ANY OPERATION OR IF WANT TO CHANGE THE FILE TO MODIFY.
    *   Parameters:
        *   `file_path` (string): The absolute path to the Excel file.
*   **`get_sheet_range_content`**:
    *   Description: Get and output table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.
    *   Parameters:
        *   `sheet_name` (string): The name of the sheet to read from.
        *   `first_row` (number): The starting row number (1-indexed).
        *   `first_column` (number): The starting column number (1-indexed).
        *   `last_row` (number): The ending row number (1-indexed).
        *   `last_column` (number): The ending column number (1-indexed).
        *   `cell_with_coord` (boolean, optional): Output non-empty cells with their respective coordinates, suitable for situations where the output area contains a large number of empty cells.
*   **`set_sheet_range_content`**:
    *   Description: Set table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.
    *   Parameters:
        *   `sheet_name` (string): The name of the sheet to write to.
        *   `first_row` (number): The starting row number (1-indexed).
        *   `first_column` (number): The starting column number (1-indexed).
        *   `values` (array[array]): The 2D array of values to write to the range (supports null, boolean, number, string types).
*   **`create_xlsx_file_by_absolute_path`**: (Note: The tool name is defined as this in the code, but the key in JSON is `create_xlsx`)
    *   Description: Create a new xlsx file with the given path. Automatically closes the Excel file after creation.
    *   Parameters:
        *   `file_path` (string): The ABSOLUTE path where the file should be created.

*(Note: The automatic open/close behavior mentioned in the tool descriptions is an internal implementation detail and does not require user attention.)*

## Building the Project

This project uses CMake for building. Installing Ninja is recommended for faster compilation speeds.

**Build Steps:**

1.  **Prepare Environment:**
    *   Ensure you have a C++17 compiler, CMake (>= 3.15), and Ninja installed.
    *   Clone the repository: `git clone https://github.com/smileFAace/MCP-ExcelAutoCpp.git`
2.  **Compile:**
    ```bash
    # Change to the project directory (if not already in the root)
    # cd MCP-ExcelAutoCpp

    # Generate build files (Ninja recommended)
    cmake -G "Ninja" -S . -B build

    # Perform compilation
    cmake --build build
    ```
    The compiled executable is typically located in the `bin/` directory.

## Usage

After successful compilation, find the executable named `ExcelAutoCpp` (Linux/macOS) or `ExcelAutoCpp.exe` (Windows) in the `./bin/` directory.

**Start the Server**

Simply run the executable to start the MCP server.

```bash
# Example (adjust path as needed)
./bin/ExcelAutoCpp
```

When you see output similar to the following, the server has started successfully and is listening on `localhost:8888`:

```

░█▀▀░█░█░█▀▀░█▀▀░█░░░█▀█░█░█░▀█▀░█▀█
░█▀▀░▄▀▄░█░░░█▀▀░█░░░█▀█░█░█░░█░░█░█
░▀▀▀░▀░▀░▀▀▀░▀▀▀░▀▀▀░▀░▀░▀▀▀░░▀░░▀▀▀
v0.0.2                 By smileFAace

I(17:35:48) i18n: Successfully loaded language string for language code 'en'
I(17:35:48) i18n: Set current language to 'en'
W(17:35:48) lang.json not found at 'E:\prj\MCP-ExcelAutoCpp\bin\lang.json'. Using default language.
I(17:35:48) i18n: Set current language to 'en'
I(17:35:48) Current language set to: en
I(17:35:48) Starting MCP server at localhost:8888
I(17:35:48) Press Ctrl+C to stop the server
```

**Connecting and Using:**

This server adheres to the MCP protocol. You can use any MCP-compatible client (such as Roo, cline, claude, cherry studio, etc.) to connect to the server's SSE endpoint (default: `http://localhost:8888/sse`) to invoke the provided Excel automation tools.

For example, to configure this MCP server in cline, simply add the following content to the mcp configuration JSON file provided by the plugin after the server is running successfully, and then refresh:
```json
{
  "mcpServers": {
    "excel-auto-cpp": {
      "url": "http://localhost:8888/sse"
    }
  }
}
```

**Changing and Customizing Server Language:**

The server defaults to English (`en`) for its interface language. You can change the language by creating a custom language file:

1.  Locate the directory containing the server executable (usually `bin/`).
2.  Create a text file named `lang.json` in that directory.
3.  Copy the **complete key-value pairs** for your desired language into this `lang.json` file.
    *   You can refer to the `*.json` files in the `lang/` directory of the project source code as templates (e.g., [`lang/zh-cn/lang.json`](lang/zh-cn/lang.json:1) contains the Chinese translations).
    *   The `lang.json` file needs to contain all translation entries required for the program interface and logs.

4.  Save the `lang.json` file and restart the server.
    *   The server will automatically detect and load this file upon startup. If loaded successfully, it will use the language defined in `lang.json`.
    *   If the `lang.json` file does not exist, is incorrectly formatted, or is missing necessary translation entries, the server will fall back to the default English (`en`) language.