#include "main.h"
 
#include <string>
#include <algorithm> // for std::reverse

// Helper function to convert column number to Excel column letter (e.g., 1 -> A, 27 -> AA)
static std::string s_colNumberToLetters(uint32_t col_num) {
   std::string col_letters = "";
   while (col_num > 0) {
       int rem = col_num % 26;
       if (rem == 0) {
           col_letters += 'Z';
           col_num = (col_num / 26) - 1;
       } else {
           col_letters += (rem - 1) + 'A';
           col_num = col_num / 26;
       }
   }
   // If the original number was 0 or negative, return empty string or handle error
   if (col_letters.empty() && col_num <= 0) {
        // Handle error or return a default, e.g., empty string or throw exception
        // For simplicity, returning empty for column 0 or less. Adjust as needed.
        return "";
   }
   std::reverse(col_letters.begin(), col_letters.end());
   return col_letters;
}

// Function to get cell address string (e.g., "A1")
static std::string s_getCellAddress(uint32_t row, uint32_t col) {
   if (row == 0 || col == 0) {
       // Handle invalid row/column index, Excel is 1-based
       return "InvalidAddress"; // Or throw an exception
   }
   return s_colNumberToLetters(col) + std::to_string(row);
}

 using ExcelWrapper::ExcelOperator;
 static const int SERVER_PORT = 8888; 
 static const char ASCII_ART[] = "\n\
 ░█▀▀░█░█░█▀▀░█▀▀░█░░░█▀█░█░█░▀█▀░█▀█\n\
 ░█▀▀░▄▀▄░█░░░█▀▀░█░░░█▀█░█░█░░█░░█░█\n\
 ░▀▀▀░▀░▀░▀▀▀░▀▀▀░▀▀▀░▀░▀░▀▀▀░░▀░░▀▀▀\n\
 v0.0.2                 By smileFAace\n";
 
ExcelOperator g_excel_operator;
std::string g_current_excel_file_path;

void ensure_excel_open() {
    if (g_current_excel_file_path.empty()) {
        spdlog::error("No Excel file path has been set. Please use 'open_excel_and_list_sheets' first.");
        throw mcp::mcp_exception(mcp::error_code::internal_error, "No Excel file path has been set. Please use 'open_excel_and_list_sheets' first.");
    }
    std::vector<std::string> dummy_sheet_names; // Only for open function signature
    if (!g_excel_operator.open(g_current_excel_file_path, dummy_sheet_names)) {
        spdlog::error("Failed to open Excel file: {}", g_current_excel_file_path);
        throw mcp::mcp_exception(mcp::error_code::internal_error, "Failed to open Excel file: " + g_current_excel_file_path);
    }
}

mcp::json open_excel_and_list_sheets_handler(const mcp::json& params, const std::string& /* session_id */) {
    if (!params.contains("file_path")) {
        spdlog::error("Missing 'file_path' parameter for open_excel_and_list_sheets.");
        throw mcp::mcp_exception(mcp::error_code::invalid_params, "Missing 'file_path' parameter");
    }

    std::string file_path = params["file_path"].get<std::string>();
    std::vector<std::string> sheet_names;

    if (g_excel_operator.open(file_path, sheet_names)) {
        g_current_excel_file_path = file_path; // for global file path
        mcp::json result_sheets = mcp::json::array();
        for (const auto& name : sheet_names) {
            result_sheets.push_back(name);
        }
        mcp::json result = {
            {
                {"type", "text"},
                {"text", result_sheets.dump()}
            }
        };
        g_excel_operator.close();
        spdlog::info("Successfully opened Excel file: {}", file_path);
        return result;
    } else {
        spdlog::error("Failed to open Excel file or list sheets: {}", file_path);
        throw mcp::mcp_exception(mcp::error_code::internal_error, "Failed to open Excel file or list sheets: " + file_path);
    }
}

mcp::json get_sheet_range_content_handler(const mcp::json& params, const std::string& /* session_id */) {
    ensure_excel_open();

    if (!params.contains("sheet_name") || !params.contains("first_row") || !params.contains("first_column") ||
        !params.contains("last_row") || !params.contains("last_column")) {
        g_excel_operator.close();
        spdlog::error("Missing required parameters for get_sheet_range_content.");
        throw mcp::mcp_exception(mcp::error_code::invalid_params, "Missing required parameters for sheet range content.");
    }

    bool seperate_cell = false;

    if (params.contains("seperate_cell")){
        seperate_cell = params["seperate_cell"].get<bool>();
    }

    std::string sheet_name = params["sheet_name"].get<std::string>();
    uint32_t first_row = params["first_row"].get<uint32_t>();
    uint32_t first_column = params["first_column"].get<uint32_t>();
    uint32_t last_row = params["last_row"].get<uint32_t>();
    uint32_t last_column = params["last_column"].get<uint32_t>();

    if (!g_excel_operator.selectSheet(sheet_name)) {
        g_excel_operator.close();
        spdlog::error("Failed to select sheet: {}", sheet_name);
        throw mcp::mcp_exception(mcp::error_code::internal_error, "Failed to select sheet: " + sheet_name);
    }

    std::vector<std::vector<OpenXLSX::XLCellValue>> range_values =
        g_excel_operator.getRangeValues(first_row, first_column, last_row, last_column);

    mcp::json result_array = mcp::json::array();
    if (seperate_cell)
    {
        uint32_t current_row = first_row;
        for (const auto& row : range_values)
        {
            uint32_t current_col = first_column;
            for (const auto& cell_value : row)
            {
                if (cell_value.type() != OpenXLSX::XLValueType::Empty)
                {
                    std::string cell_content_str;
                    if (cell_value.type() == OpenXLSX::XLValueType::Boolean)
                    {
                        cell_content_str = cell_value.get<bool>() ? "TRUE" : "FALSE";
                    }
                    else if (cell_value.type() == OpenXLSX::XLValueType::Integer)
                    {
                        cell_content_str = std::to_string(cell_value.get<int64_t>());
                    }
                    else if (cell_value.type() == OpenXLSX::XLValueType::Float)
                    {
                        // Use std::ostringstream for better float formatting if needed
                        cell_content_str = std::to_string(cell_value.get<double>());
                    }
                    else if (cell_value.type() == OpenXLSX::XLValueType::String)
                    {
                        cell_content_str = cell_value.get<std::string>();
                    }
                    else
                    { // Consider other types as string for simplicity
                        try {
                           cell_content_str = cell_value.get<std::string>();
                        } catch (const OpenXLSX::XLValueTypeError& e) {
                           // Handle cases where conversion to string might fail for unexpected types
                           cell_content_str = "[Unsupported Type]";
                           spdlog::warn("Unsupported cell type encountered at row {}, col {}: {}", current_row, current_col, e.what());
                        }
                    }
                    std::string cell_address = s_getCellAddress(current_row, current_col);
                    result_array.push_back(cell_content_str + "@" + cell_address);
                }
                current_col++;
            }
            current_row++;
        }
    }
    else
    {
        for (const auto& row : range_values)
        {
            mcp::json row_array = mcp::json::array();
            for (const auto& cell_value : row)
            {
                if (cell_value.type() == OpenXLSX::XLValueType::Empty)
                {
                    row_array.push_back(nullptr);
                }
                else if (cell_value.type() == OpenXLSX::XLValueType::Boolean)
                {
                    row_array.push_back(cell_value.get<bool>());
                }
                else if (cell_value.type() == OpenXLSX::XLValueType::Integer)
                {
                    row_array.push_back(cell_value.get<int64_t>());
                }
                else if (cell_value.type() == OpenXLSX::XLValueType::Float)
                {
                    row_array.push_back(cell_value.get<double>());
                }
                else
                { // Treat String and others similarly
                     try {
                        row_array.push_back(cell_value.get<std::string>());
                     } catch (const OpenXLSX::XLValueTypeError& e) {
                        row_array.push_back("[Unsupported Type]");
                         // Optionally log the error with row/col if needed, though harder without tracking here
                        spdlog::warn("Unsupported cell type encountered during standard processing: {}", e.what());
                     }
                }
            }
            result_array.push_back(row_array);
        }
    }

    mcp::json result = {
        {
            {"type", "text"},
            {"text", result_array.dump()}
        }
    };
    g_excel_operator.close();
    spdlog::info("Successfully retrieved sheet range content from sheet: {}", sheet_name);
    return result;
}

mcp::json create_xlsx_file_handler(const mcp::json& params, const std::string& /* session_id */) {
    if (!params.contains("file_path")) {
        spdlog::error("Missing 'file_path' parameter for create_xlsx_file.");
        throw mcp::mcp_exception(mcp::error_code::invalid_params, "Missing 'file_path' parameter");
    }

    std::string file_path = params["file_path"].get<std::string>();

    if (g_excel_operator.create(file_path)) {
        g_current_excel_file_path = file_path;
        mcp::json result = {
            {
                {"type", "text"},
                {"text", "Excel file created successfully: " + file_path}
            }
        };
        g_excel_operator.close();
        spdlog::info("Successfully created Excel file: {}", file_path);
        return result;
    } else {
        spdlog::error("Failed to create Excel file: {}", file_path);
        throw mcp::mcp_exception(mcp::error_code::internal_error, "Failed to create Excel file: " + file_path);
    }
}

mcp::json set_sheet_range_content_handler(const mcp::json& params, const std::string& /* session_id */) {
    ensure_excel_open();

    if (!params.contains("sheet_name") || !params.contains("first_row") || !params.contains("first_column") ||
        !params.contains("values")) {
        g_excel_operator.close();
        spdlog::error("Missing required parameters for set_sheet_range_content.");
        throw mcp::mcp_exception(mcp::error_code::invalid_params, "Missing required parameters for setting sheet range content.");
    }

    std::string sheet_name = params["sheet_name"].get<std::string>();
    uint32_t first_row = params["first_row"].get<uint32_t>();
    uint32_t first_column = params["first_column"].get<uint32_t>();
    mcp::json json_values = params["values"];

    if (!json_values.is_array()) {
        g_excel_operator.close();
        spdlog::error("'values' parameter must be a 2D array for set_sheet_range_content.");
        throw mcp::mcp_exception(mcp::error_code::invalid_params, "'values' parameter must be a 2D array.");
    }

    std::vector<std::vector<OpenXLSX::XLCellValue>> values_to_set;
    for (const auto& row_json : json_values) {
        if (!row_json.is_array()) {
            g_excel_operator.close();
            spdlog::error("Each row in 'values' must be an array for set_sheet_range_content.");
            throw mcp::mcp_exception(mcp::error_code::invalid_params, "Each row in 'values' must be an array.");
        }
        std::vector<OpenXLSX::XLCellValue> row_values;
        for (const auto& cell_json : row_json) {
            if (cell_json.is_boolean()) {
                row_values.push_back(OpenXLSX::XLCellValue(cell_json.get<bool>()));
            } else if (cell_json.is_number_integer()) {
                row_values.push_back(OpenXLSX::XLCellValue(cell_json.get<int64_t>()));
            } else if (cell_json.is_number_float()) {
                row_values.push_back(OpenXLSX::XLCellValue(cell_json.get<double>()));
            } else if (cell_json.is_string()) {
                row_values.push_back(OpenXLSX::XLCellValue(cell_json.get<std::string>()));
            } else if (cell_json.is_null()) {
                row_values.push_back(OpenXLSX::XLCellValue());
            } else {
                g_excel_operator.close();
                spdlog::error("Unsupported cell value type in 'values' array for set_sheet_range_content.");
                throw mcp::mcp_exception(mcp::error_code::invalid_params, "Unsupported cell value type in 'values' array.");
            }
        }
        values_to_set.push_back(row_values);
    }

    if (!g_excel_operator.selectSheet(sheet_name)) {
        g_excel_operator.close();
        spdlog::error("Failed to select sheet: {}", sheet_name);
        throw mcp::mcp_exception(mcp::error_code::internal_error, "Failed to select sheet: " + sheet_name);
    }

    if (g_excel_operator.setRangeValues(first_row, first_column, values_to_set)) {
        mcp::json result = {
            {
                {"type", "text"},
                {"text", "Successfully set sheet range content."}
            }
        };
        g_excel_operator.close();
        spdlog::info("Successfully set sheet range content for sheet: {}", sheet_name);
        return result;
    } else {
        g_excel_operator.close();
        spdlog::error("Failed to set sheet range content for sheet: {}", sheet_name);
        throw mcp::mcp_exception(mcp::error_code::internal_error, "Failed to set sheet range content.");
    }
}

static void s_spdlog_init() {

    spdlog::set_pattern("%^%L%$(%H:%M:%S) %v");

    bool console_sink_exists = false;
    for (const auto& sink : spdlog::default_logger()->sinks()) {
        if (std::dynamic_pointer_cast<spdlog::sinks::stdout_color_sink_mt>(sink)) {
            console_sink_exists = true;
            break;
        }
    }

    if (!console_sink_exists) {
        auto console_sink = std::make_shared<spdlog::sinks::stdout_color_sink_mt>();
        spdlog::default_logger()->sinks().push_back(console_sink);
    }
}



static void s_mcpServer_init(mcp::server& server, bool blocking_mode) {
    server.set_server_info("ExcelAutoCpp", "1.0.0");
    
    mcp::json capabilities = {
        {"tools", mcp::json::object()}
    };
    server.set_capabilities(capabilities);

    mcp::tool open_excel_tool = mcp::tool_builder("open_excel_and_list_sheets")
        .with_description("Open an Excel file and list all sheet names. This tool will also set the current Excel file path for subsequent operations.RECOMMANDED TO RUN THIS TOOL FIRST BEFORE ANY OPERATION OR IF WANT TO CHANGE THE FILE TO MODIFY.")
        .with_string_param("file_path", "The full path to the Excel file")
        .build();
    server.register_tool(open_excel_tool, open_excel_and_list_sheets_handler);

    mcp::tool get_range_tool = mcp::tool_builder("get_sheet_range_content")
        .with_description("Get and output table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.")
        .with_string_param("sheet_name", "The name of the sheet to read from")
        .with_number_param("first_row", "The starting row number (1-indexed)")
        .with_number_param("first_column", "The starting column number (1-indexed)")
        .with_number_param("last_row", "The ending row number (1-indexed)")
        .with_number_param("last_column", "The ending column number (1-indexed)")
        .with_boolean_param("seperate_cell", "Output the none-null cells seperately, suitable for the sheet with many null cells")
        .build();
    server.register_tool(get_range_tool, get_sheet_range_content_handler);

    mcp::tool set_range_tool = mcp::tool_builder("set_sheet_range_content")
        .with_description("Set table content within a specified range in a specific sheet. Automatically opens and closes the Excel file.")
        .with_string_param("sheet_name", "The name of the sheet to write to")
        .with_number_param("first_row", "The starting row number (1-indexed)")
        .with_number_param("first_column", "The starting column number (1-indexed)")
        .with_array_param("values", "The 2D array of values to write to the range", "object")
        .build();
    server.register_tool(set_range_tool, set_sheet_range_content_handler);

    mcp::tool create_xlsx_tool = mcp::tool_builder("create_xlsx_file_by_absolute_path")
       .with_description("Create a new xlsx file with the given path. Automatically closes the Excel file after creation.")
       .with_string_param("file_path", "The ABSOLUTE path with which the file should create to")
       .build();
    server.register_tool(create_xlsx_tool, create_xlsx_file_handler);

    spdlog::info("Starting MCP server at localhost:{}", SERVER_PORT);
    spdlog::info("Press Ctrl+C to stop the server");
    

    server.start(blocking_mode);
}

int main() {
#ifdef _WIN32
    #include <windows.h>
    SetConsoleOutputCP(CP_UTF8);
#endif

    std::cout << ASCII_ART << std::endl;
    spdlog::set_level(spdlog::level::info);
    s_spdlog_init();

    mcp::server server("localhost", SERVER_PORT);
    mcp::set_log_level(mcp::log_level::error);
    s_mcpServer_init(server, true);
    
    return 0;
}
