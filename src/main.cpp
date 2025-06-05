// Include project headers first
#include "i18n.h" // Include the i18n header
#include "embedded_translations.h" // Include the embedded translations
// Include the precompiled header last among project headers
#include "main.h"

#include <string>
#include <filesystem> // Required for path operations
#include <algorithm> // for std::reverse

using ExcelWrapper::ExcelOperator;

static const char DEFAULT_LANG[] = "zh-CN";
static const int SERVER_PORT = 8888; 

static const char ASCII_ART[] = "\n\
░█▀▀░█░█░█▀▀░█▀▀░█░░░█▀█░█░█░▀█▀░█▀█\n\
░█▀▀░▄▀▄░█░░░█▀▀░█░░░█▀█░█░█░░█░░█░█\n\
░▀▀▀░▀░▀░▀▀▀░▀▀▀░▀▀▀░▀░▀░▀▀▀░░▀░░▀▀▀\n\
v0.0.2                 By smileFAace\n";
 
ExcelOperator g_excel_operator;
std::string g_current_excel_file_path;


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
       return i18n::t("result.invalid_address"); // Or throw an exception
   }
   return s_colNumberToLetters(col) + std::to_string(row);
}

void ensure_excel_open() {
    if (g_current_excel_file_path.empty()) {
        spdlog::error(i18n::t("log.error.no_excel_path"));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.no_excel_path"));
    }
    std::vector<std::string> dummy_sheet_names; // Only for open function signature
    if (!g_excel_operator.open(g_current_excel_file_path, dummy_sheet_names)) {
        spdlog::error(i18n::t("log.error.failed_open_excel", g_current_excel_file_path));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.failed_open_excel", g_current_excel_file_path));
    }
}

mcp::json open_excel_and_list_sheets_handler(const mcp::json& params, const std::string& /* session_id */) {
    if (!params.contains("file_path")) {
        spdlog::error(i18n::t("log.error.missing_params.create_xlsx")); // Reusing create_xlsx key as it's just file_path
        throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.missing_param.file_path"));
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
        spdlog::info(i18n::t("log.info.opened_excel", file_path));
        return result;
    } else {
        spdlog::error(i18n::t("log.error.failed_open_or_list", file_path));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.failed_open_or_list", file_path));
    }
}

mcp::json get_sheet_range_content_handler(const mcp::json& params, const std::string& /* session_id */) {
    ensure_excel_open();

    if (!params.contains("sheet_name") || !params.contains("first_row") || !params.contains("first_column") ||
        !params.contains("last_row") || !params.contains("last_column")) {
        g_excel_operator.close();
        spdlog::error(i18n::t("log.error.missing_params.get_range"));
        throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.missing_params.get_range"));
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
        spdlog::error(i18n::t("log.error.failed_select_sheet", sheet_name));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.failed_select_sheet", sheet_name));
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
                           cell_content_str = i18n::t("result.unsupported_type");
                           spdlog::warn(i18n::t("log.warn.unsupported_cell_type.get_range", current_row, current_col, e.what()));
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
                        row_array.push_back(i18n::t("result.unsupported_type"));
                         // Optionally log the error with row/col if needed, though harder without tracking here
                        spdlog::warn(i18n::t("log.warn.unsupported_cell_type.standard", e.what()));
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
    spdlog::info(i18n::t("log.info.retrieved_range", sheet_name));
    return result;
}

mcp::json create_xlsx_file_handler(const mcp::json& params, const std::string& /* session_id */) {
    if (!params.contains("file_path")) {
        spdlog::error(i18n::t("log.error.missing_params.create_xlsx"));
        throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.missing_param.file_path_create"));
    }

    std::string file_path = params["file_path"].get<std::string>();

    if (g_excel_operator.create(file_path)) {
        g_current_excel_file_path = file_path;
        mcp::json result = {
            {
                {"type", "text"},
                {"text", i18n::t("result.created_excel", file_path)}
            }
        };
        g_excel_operator.close();
        spdlog::info(i18n::t("log.info.created_excel", file_path));
        return result;
    } else {
        spdlog::error(i18n::t("log.error.failed_create_excel", file_path));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.failed_create_excel", file_path));
    }
}

mcp::json set_sheet_range_content_handler(const mcp::json& params, const std::string& /* session_id */) {
    ensure_excel_open();

    if (!params.contains("sheet_name") || !params.contains("first_row") || !params.contains("first_column") ||
        !params.contains("values")) {
        g_excel_operator.close();
        spdlog::error(i18n::t("log.error.missing_params.set_range"));
        throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.missing_params.set_range"));
    }

    std::string sheet_name = params["sheet_name"].get<std::string>();
    uint32_t first_row = params["first_row"].get<uint32_t>();
    uint32_t first_column = params["first_column"].get<uint32_t>();
    mcp::json json_values = params["values"];

    if (!json_values.is_array()) {
        g_excel_operator.close();
        spdlog::error(i18n::t("log.error.values_not_2d_array"));
        throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.values_not_2d_array"));
    }

    std::vector<std::vector<OpenXLSX::XLCellValue>> values_to_set;
    for (const auto& row_json : json_values) {
        if (!row_json.is_array()) {
            g_excel_operator.close();
            spdlog::error(i18n::t("log.error.values_row_not_array"));
            throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.values_row_not_array"));
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
                spdlog::error(i18n::t("log.error.unsupported_cell_type.set_range"));
                throw mcp::mcp_exception(mcp::error_code::invalid_params, i18n::t("exception.error.unsupported_cell_type.set_range"));
            }
        }
        values_to_set.push_back(row_values);
    }

    if (!g_excel_operator.selectSheet(sheet_name)) {
        g_excel_operator.close();
        spdlog::error(i18n::t("log.error.failed_select_sheet", sheet_name));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.failed_select_sheet", sheet_name));
    }

    if (g_excel_operator.setRangeValues(first_row, first_column, values_to_set)) {
        mcp::json result = {
            {
                {"type", "text"},
                {"text", i18n::t("result.set_range")}
            }
        };
        g_excel_operator.close();
        spdlog::info(i18n::t("log.info.set_range", sheet_name));
        return result;
    } else {
        g_excel_operator.close();
        spdlog::error(i18n::t("log.error.failed_set_range", sheet_name));
        throw mcp::mcp_exception(mcp::error_code::internal_error, i18n::t("exception.error.failed_set_range"));
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
    server.set_server_info("ExcelAutoCpp", "1.0.0"); // Server name/version likely not translated

    mcp::json capabilities = {
        {"tools", mcp::json::object()}
    };
    server.set_capabilities(capabilities);

    mcp::tool open_excel_tool = mcp::tool_builder("open_excel_and_list_sheets")
        .with_description(i18n::t("tool.open_excel.description"))
        .with_string_param("file_path", i18n::t("tool.open_excel.param.file_path"))
        .build();
    server.register_tool(open_excel_tool, open_excel_and_list_sheets_handler);

    mcp::tool get_range_tool = mcp::tool_builder("get_sheet_range_content")
        .with_description(i18n::t("tool.get_range.description"))
        .with_string_param("sheet_name", i18n::t("tool.get_range.param.sheet_name"))
        .with_number_param("first_row", i18n::t("tool.get_range.param.first_row"))
        .with_number_param("first_column", i18n::t("tool.get_range.param.first_column"))
        .with_number_param("last_row", i18n::t("tool.get_range.param.last_row"))
        .with_number_param("last_column", i18n::t("tool.get_range.param.last_column"))
        .with_boolean_param("cell_with_coord", i18n::t("tool.get_range.param.cell_with_coord")) // Note: Key was 'seperate_cell' in code, 'cell_with_coord' in JSON
        .build();
    server.register_tool(get_range_tool, get_sheet_range_content_handler);

    mcp::tool set_range_tool = mcp::tool_builder("set_sheet_range_content")
        .with_description(i18n::t("tool.set_range.description"))
        .with_string_param("sheet_name", i18n::t("tool.set_range.param.sheet_name"))
        .with_number_param("first_row", i18n::t("tool.set_range.param.first_row"))
        .with_number_param("first_column", i18n::t("tool.set_range.param.first_column"))
        .with_array_param("values", i18n::t("tool.set_range.param.values"), "object") // Schema type "object" likely remains untranslated
        .build();
    server.register_tool(set_range_tool, set_sheet_range_content_handler);

    mcp::tool create_xlsx_tool = mcp::tool_builder("create_xlsx_file_by_absolute_path")
       .with_description(i18n::t("tool.create_xlsx.description"))
       .with_string_param("file_path", i18n::t("tool.create_xlsx.param.file_path"))
       .build();
    server.register_tool(create_xlsx_tool, create_xlsx_file_handler);

    spdlog::info(i18n::t("log.info.server_start", SERVER_PORT));
    spdlog::info(i18n::t("log.info.server_stop_prompt"));


    server.start(blocking_mode);
}

#ifdef _WIN32
#include <windows.h> // Moved include here as it's needed by s_i18n_init
#endif

static void s_i18n_init() {
    auto& i18n = i18n::I18nManager::getInstance();

    // Load English from embedded string in the dedicated header
    bool en_loaded = i18n.loadLanguageFromString("en", embedded_translations::EN_JSON);
    if (!en_loaded) {
        spdlog::error("Failed to load embedded English language string from header.");
    }

    // Determine the path to lang.json relative to the executable
    std::filesystem::path exe_path;
#ifdef _WIN32
    char path_buf[MAX_PATH];
    GetModuleFileNameA(NULL, path_buf, MAX_PATH);
    exe_path = path_buf;
#else
    // For non-Windows, assume current working directory or add platform-specific logic
    exe_path = std::filesystem::current_path();
#endif
    std::filesystem::path lang_json_path = exe_path.parent_path() / "lang.json";

    // Load language from lang.json
    bool lang_json_loaded = false;
    if (std::filesystem::exists(lang_json_path)) {
        lang_json_loaded = i18n.loadLanguage("custom", lang_json_path.string());
        if (!lang_json_loaded) {
            spdlog::error("Failed to load language file from '{}'.", lang_json_path.string());
        }
    } else {
        spdlog::warn("lang.json not found at '{}'. Using default language.", lang_json_path.string());
    }

    // Set default language
    if (lang_json_loaded) {
        if (!i18n.setLanguage("custom")) {
            spdlog::error("Failed to set 'custom' language from lang.json. Falling back to English.");
            i18n.setLanguage("en"); // Fallback to English
        }
    } else if (!i18n.setLanguage("en")) { // Fallback to English if lang.json not loaded or failed
        spdlog::critical("Failed to load ANY language data. Application might not function correctly.");
    }

    spdlog::info("Current language set to: {}", i18n.getCurrentLanguage());
}


int main() {
#ifdef _WIN32
    SetConsoleOutputCP(CP_UTF8);
#endif

    std::cout << ASCII_ART << std::endl;

    spdlog::set_level(spdlog::level::info);
    s_spdlog_init();

    s_i18n_init(); 


    mcp::server server("localhost", SERVER_PORT);
    mcp::set_log_level(mcp::log_level::error); // Keep MCP library logs concise
    s_mcpServer_init(server, true);

    return 0;
}
