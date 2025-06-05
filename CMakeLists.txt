cmake_minimum_required(VERSION 3.10)

# -------- Compile options --------
option(BUILD_RELEASE "Build in release mode with O3 optimization" ON)
option(PROJECT_STATIC "Build project as static library/executable" ON)
option(THIRD_LIB_STATIC "Build third-party libraries as static libraries" ON)


# -------- Project overall compile setting --------
if(BUILD_RELEASE)
    set(CMAKE_BUILD_TYPE Release)
    add_compile_options(-O3)
    message(STATUS "Compiling in Release mode, with O3 optimization enabled.")
else()
    set(CMAKE_BUILD_TYPE Debug)
    message(STATUS "Compiling in Debug mode.")
endif()

project(ExcelAutoCpp CXX)

# Ensures static linking of C/C++ runtime libraries for MinGW.
# This helps create a self-contained executable that doesn't rely on external runtime DLLs.
if(MINGW AND PROJECT_STATIC)
    set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -static")
    message(STATUS "Static linking of C/C++ runtime libraries for MinGW enabled.")
else()
    message(STATUS "Static linking of C/C++ runtime libraries for MinGW disabled.Executable file might need DLLs to run.")
endif()

 # Ensures that third-party libraries are built as static libraries (.lib or .a) instead of dynamic libraries (DLLs or .so).
if(THIRD_LIB_STATIC)
    set(BUILD_SHARED_LIBS OFF)
    message(STATUS "Third-party libraries are built as static libraries.")
else()
    set(BUILD_SHARED_LIBS ON)
    message(STATUS "Third-party libraries are built as dynamic libraries.")
endif()

set(CMAKE_EXPORT_COMPILE_COMMANDS ON) # for compile_commands.json for clangd

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED True)

# Set compiler paths (can be overridden by user)
if(NOT CMAKE_CXX_COMPILER)
    set(CMAKE_CXX_COMPILER "") # Placeholder for C++ compiler path
endif()
if(NOT CMAKE_C_COMPILER)
    set(CMAKE_C_COMPILER "") # Placeholder for C compiler path
endif()
message(STATUS "C compiler: ${CMAKE_C_COMPILER}")
message(STATUS "C++ compiler: ${CMAKE_CXX_COMPILER}")

# Tell the compiler that source files are UTF-8 encoded
if(MSVC)
    add_compile_options(/utf-8)
    message(STATUS "MSVC compiler: Adding /utf-8 compile option.")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    add_compile_options(-finput-charset=UTF-8 -fexec-charset=UTF-8)
    message(STATUS "GCC/Clang compiler: Adding -finput-charset=UTF-8 -fexec-charset=UTF-8 compile options.")
endif()

# Specify the path of exec file
set(EXECUTABLE_OUTPUT_PATH ${PROJECT_SOURCE_DIR}/bin)
set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${PROJECT_SOURCE_DIR}/bin)

# -------- Application Core --------
# include_directories(src) # Removed global include
add_executable(${PROJECT_NAME}
src/main.cpp
src/ExcelOperator.cpp
src/i18n.cpp
)
target_precompile_headers(${PROJECT_NAME} PRIVATE src/main.h) # Set precompiled header

# Add include directories specifically for the target
target_include_directories(${PROJECT_NAME} PUBLIC src) # Add src directory

# -------- Third-party Libraries --------
# cpp-mcp
add_subdirectory(extlib/cpp-mcp)
target_include_directories(${PROJECT_NAME} PRIVATE extlib/cpp-mcp/include extlib/cpp-mcp/common) # Add cpp-mcp includes
if(WIN32)
    target_link_libraries(${PROJECT_NAME} PRIVATE mcp ws2_32 iphlpapi)
else()
    target_link_libraries(${PROJECT_NAME} PRIVATE mcp)
endif()
# OpenXLSX
add_subdirectory(extlib/OpenXLSX)
target_include_directories(${PROJECT_NAME} PRIVATE extlib/OpenXLSX/OpenXLSX/headers) # Add OpenXLSX includes
target_link_libraries(${PROJECT_NAME} PRIVATE OpenXLSX::OpenXLSX)
# spdlog
add_subdirectory(extlib/spdlog)
target_include_directories(${PROJECT_NAME} PRIVATE extlib/spdlog/include) # Add spdlog includes
target_link_libraries(${PROJECT_NAME} PRIVATE spdlog)

# -------- Copy Resources --------
# Copy the lang directory to the executable output directory after build
set(LANGUAGE_NAME zh-cn)
add_custom_command(TARGET ${PROJECT_NAME} POST_BUILD
    COMMAND ${CMAKE_COMMAND} -E copy
        ${PROJECT_SOURCE_DIR}/lang/${LANGUAGE_NAME}/lang.json $<TARGET_FILE_DIR:${PROJECT_NAME}>/lang.json
    COMMENT "Copying lang.json to output directory")
