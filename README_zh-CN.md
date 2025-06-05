Language：[中文](README_zh-CN.md) | [English](README.md)
# C++ Excel 自动化 MCP 服务器

这是一个基于 C++ 的 Excel 自动化 MCP (模型上下文协议) 服务器项目。本项目使用 OpenXLSX 库进行 Excel 文件操作，旨在通过 MCP 协议配合 LLM 能力提供一套智能化的 Excel 操作工具。

## 目录

*   [项目特点](#项目特点)
*   [构建项目](#构建项目)
*   [使用方法](#使用方法)


## 项目特点

*   **简洁易用的 MCP 接口**: 提供标准化的 MCP 工具，方便客户端（如 AI 助手）调用 Excel 自动化功能。
*   **单文件部署**: 编译后生成单个可执行文件，方便部署和运行。
*   **可定制多语言**: 支持通过 JSON 文件轻松添加或修改界面语言。

## 服务器能力

本 MCP 服务器提供以下工具，可通过 LLM 调用从而实现和 xlsx 表格文件的交互：

*   **`open_excel_and_list_sheets`**:
    *   描述: 打开一个 Excel 文件并列出所有工作表名称。此工具还将设置当前 Excel 文件路径以供后续操作使用。建议在进行任何操作之前或想要更改要修改的文件时首先运行此工具。
    *   参数:
        *   `file_path` (string): Excel 文件的绝对路径。
*   **`get_sheet_range_content`**:
    *   描述: 获取并输出指定工作表中指定范围内的表格内容。自动打开和关闭 Excel 文件。
    *   参数:
        *   `sheet_name` (string): 要读取的工作表名称。
        *   `first_row` (number): 起始行号（从 1 开始）。
        *   `first_column` (number): 起始列号（从 1 开始）。
        *   `last_row` (number): 结束行号（从 1 开始）。
        *   `last_column` (number): 结束列号（从 1 开始）。
        *   `cell_with_coord` (boolean, 可选): 输出非空单元格及其各自的坐标，适用于输出区域包含大量空单元格的情况。
*   **`set_sheet_range_content`**:
    *   描述: 设置指定工作表中指定范围内的表格内容。自动打开和关闭 Excel 文件。
    *   参数:
        *   `sheet_name` (string): 要写入的工作表名称。
        *   `first_row` (number): 起始行号（从 1 开始）。
        *   `first_column` (number): 起始列号（从 1 开始）。
        *   `values` (array[array]): 要写入范围的二维数组值 (支持 null, boolean, number, string 类型)。
*   **`create_xlsx_file_by_absolute_path`**: (注意：工具名称在代码中定义为此，但 JSON 中键为 `create_xlsx`)
    *   描述: 使用给定路径创建一个新的 xlsx 文件。创建后自动关闭 Excel 文件。
    *   参数:
        *   `file_path` (string): 文件应创建到的绝对路径。

*(注意：工具描述中提到的自动打开/关闭行为是内部实现细节，用户无需关心。)*

## 构建项目

本项目使用 CMake 构建。推荐安装 Ninja 以获得更快的编译速度。

**构建步骤:**

1.  **准备环境:**
    *   确保已安装 C++17 编译器、CMake (>= 3.15) 和 Ninja。
    *   克隆仓库: `git clone https://github.com/smileFAace/MCP-ExcelAutoCpp.git`
2.  **编译:**
    ```bash
    # 进入项目目录 (如果不在项目根目录)
    # cd MCP-ExcelAutoCpp
    
    # 生成构建文件 (推荐使用 Ninja)
    cmake -G "Ninja" -S . -B build
    
    # 执行编译
    cmake --build build
    ```
    编译后的可执行文件通常位于 `bin/` 目录下。

## 使用方法

编译成功后，在 `./bin/` 目录下找到名为 `ExcelAutoCpp` (Linux/macOS) 或 `ExcelAutoCpp.exe` (Windows) 的可执行文件。

**启动服务器**

直接运行该可执行文件即可启动 MCP 服务器。

```bash
# 示例 (请根据实际路径调整)
./bin/ExcelAutoCpp
```

当看到类似以下的输出时，表示服务器已成功启动并监听 `localhost:8888`：

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

**连接与使用:**

本服务器遵循 MCP 协议。您可以使用任何兼容 MCP 的客户端（如 Roo、cline、claude、cherry studio 等）连接到服务器的 SSE 端点（默认为 `http://localhost:8888/sse`）来调用其提供的 Excel 自动化工具。

例如，在 cline 中配置该 MCP 服务器，只需在服务运行成功后于该插件所提供的 mcp 配置 json 文件中添加下述内容并刷新：
```json
{
  "mcpServers": {
    "excel-auto-cpp": {
      "url": "http://localhost:8888/sse"
    }
  }
}
```

**更改和自定义服务器语言:**

服务器默认使用英文 (`en`) 作为界面语言。您可以通过创建自定义语言文件来更改语言：

1.  找到服务器可执行文件所在的目录（通常是 `bin/`）。
2.  在该目录下创建一个名为 `lang.json` 的文本文件。
3.  将您希望使用的语言的**完整翻译键值对**复制到这个 `lang.json` 文件中。
    *   您可以参考项目源码中 `lang/` 目录下的 `*.json` 文件作为模板（例如，[`lang/zh-cn/lang.json`](lang/zh-cn/lang.json) 包含中文翻译）。
    *   `lang.json` 文件需要包含所有程序界面和日志所需的翻译条目。

4.  保存 `lang.json` 文件并重新启动服务器。
    *   服务器启动时会自动检测并加载该文件。如果加载成功，将使用 `lang.json` 中定义的语言。
    *   如果 `lang.json` 文件不存在、格式错误或缺少必要的翻译条目，服务器将回退到默认的英文 (`en`) 语言。
    ```bash

    ░█▀▀░█░█░█▀▀░█▀▀░█░░░█▀█░█░█░▀█▀░█▀█
    ░█▀▀░▄▀▄░█░░░█▀▀░█░░░█▀█░█░█░░█░░█░█
    ░▀▀▀░▀░▀░▀▀▀░▀▀▀░▀▀▀░▀░▀░▀▀▀░░▀░░▀▀▀
    v0.0.2                 By smileFAace

    I(20:27:30) i18n: Successfully loaded language string for language code 'en'
    I(20:27:30) i18n: Set current language to 'en'
    I(20:27:30) i18n: Successfully loaded language file 'F:\Prj\MCP-ExcelAutoCpp\bin\lang.json' for language code 'custom'
    I(20:27:30) i18n: Set current language to 'custom'
    I(20:27:30) Current language set to: custom
    I(20:27:30) 在 localhost:8888启动 MCP 服务器
    I(20:27:30) 按 Ctrl+C 停止服务器
    ```


