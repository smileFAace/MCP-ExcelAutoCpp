# C++ Excel Automation MCP Server

This is a C++-based Excel automation MCP (Model Context Protocol) server project. It leverages the OpenXLSX library for Excel file operations and uses spdlog for logging. The project aims to provide a set of programmable Excel manipulation tools via the MCP protocol.

## Table of Contents

*   [Building the Project](#building-the-project)
    *   [Prerequisites](#prerequisites)
    *   [Getting Submodules](#getting-submodules)
    *   [Compilation Steps](#compilation-steps)
*   [Usage](#usage)
    *   [Running the Server](#running-the-server)
    *   [MCP Server Capabilities](#mcp-server-capabilities)
    *   [Provided Tools](#provided-tools)

## Building the Project

This project uses CMake as its build system, and it is recommended to use Ninja as the generator for compilation.

### Prerequisites

Before compiling, please ensure the following tools are installed on your system:

*   **C++ Compiler**: A C++17 compliant compiler (e.g., GCC, Clang, MSVC).
*   **CMake**: Version 3.15 or higher.
*   **Ninja**: A fast build system.

### Getting Submodules

This project uses Git submodules for its external libraries. After cloning the repository, you need to initialize and update these submodules:

```bash
git submodule update --init --recursive
```

### Compilation Steps

1.  **Generate Build Files(Ninja should be installed first for fast compile)**:
    ```bash
    cmake -G "Ninja" -S . -B build
    ```
    This command tells CMake to generate Ninja build files in a `build` directory. If you are on Windows using Visual Studio, you might consider using `cmake -G "Visual Studio 17 2022" -S . -B build` (or your installed Visual Studio version).

2.  **Compile the Project**:
    ```bash
    cmake --build build
    ```
    This command will compile the project using the generated build files in the `build` directory. The executable is typically located in the `build/bin/` directory.

## Usage

After compilation, this project generates an executable that runs as an MCP server.

### Running the Server

In the `build/bin/` directory, you will find the executable. The name and extension of the executable will vary depending on your operating system (e.g., `ExcelAutoCpp.exe` on Windows, `ExcelAutoCpp` on Linux/macOS). Run this file to start the MCP server:

```bash
./bin/ExcelAutoCpp
```

And when this shows up, it means the server is ready to provide the funciton:
```bash

 ░█▀▀░█░█░█▀▀░█▀▀░█░░░█▀█░█░█░▀█▀░█▀█
 ░█▀▀░▄▀▄░█░░░█▀▀░█░░░█▀█░█░█░░█░░█░█
 ░▀▀▀░▀░▀░▀▀▀░▀▀▀░▀▀▀░▀░▀░▀▀▀░░▀░░▀▀▀
 v0.0.1                  

I(12:17:38) Starting MCP server at localhost:8888       
I(12:17:38) Press Ctrl+C to stop the server
```

### MCP Server Capabilities

This project functions as an MCP server. Its capabilities are defined by its configuration. For a detailed understanding of the tools it provides and how to interact with them, refer to the following server configuration:

```json
{
  "mcpServers": {
    "excel-auto-cpp": {
      "url": "http://localhost:8888/sse",
      "disabled": false,
      "timeout": 15
    }
  }
}
```

You can connect to this server using any MCP-compatible client (e.g., Roo, cline, Claude, or other custom applications/integrated environments) and invoke these tools to automate Excel operations.

### Provided Tools

This MCP server provides the following tools for Excel automation:

#### `open_excel_and_list_sheets`

*   **Description**: Open an Excel file and list all sheet names. This tool will also set the current Excel file path for subsequent operations. It is recommended to run this tool first before any operation or if you want to change the file to modify.
*   **Parameters**:
    *   `file_path` (string): The full path to the Excel file.

#### `get_sheet_range_content`

*   **Description**: Get and output table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.
*   **Parameters**:
    *   `sheet_name` (string): The name of the sheet to read from.
    *   `first_row` (number): The starting row number (1-indexed).
    *   `first_column` (number): The starting column number (1-indexed).
    *   `last_row` (number): The ending row number (1-indexed).
    *   `last_column` (number): The ending column number (1-indexed).

#### `set_sheet_range_content`

*   **Description**: Set table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.
*   **Parameters**:
    *   `sheet_name` (string): The name of the sheet to write to.
    *   `first_row` (number): The starting row number (1-indexed).
    *   `first_column` (number): The starting column number (1-indexed).
    *   `values` (array of array of object): The 2D array of values to write to the range.

#### `create_xlsx_file`

*   **Description**: Create a new xlsx file with the given path. Automatically closes the Excel file after creation.
*   **Parameters**:
    *   `file_path` (string): The path with which the file should be created.
