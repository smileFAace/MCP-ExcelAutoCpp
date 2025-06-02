#ifndef EXCEL_OPERATOR_H
#define EXCEL_OPERATOR_H

#include <string>
#include <vector>
#include <cstdint>

#include <OpenXLSX.hpp>

using OpenXLSX::XLCellValue;

namespace ExcelWrapper {

class ExcelOperator {
public:
    ExcelOperator();
    ~ExcelOperator();

    bool open(const std::string& filePath, std::vector<std::string>& sheetNames);
    bool create(const std::string& filePath);
    bool save();
    bool saveAs(const std::string& filePath);
    bool close();

    bool selectSheet(const std::string& sheetName);
    bool selectSheet(uint32_t sheetIndex);
    bool addSheet(const std::string& sheetName);
    bool deleteSheet(const std::string& sheetName);
    bool renameSheet(const std::string& oldName, const std::string& newName);
    uint32_t sheetCount() const;
    std::string currentSheetName() const;

    template<typename T>
    void setCellValue(const std::string& cellReference, const T& value);

    template<typename T>
    T getCellValue(const std::string& cellReference);

    bool clearCell(uint32_t row, uint32_t column);
    bool mergeCells(uint32_t firstRow, uint32_t firstColumn, uint32_t lastRow, uint32_t lastColumn);
    bool unmergeCells(uint32_t firstRow, uint32_t firstColumn, uint32_t lastRow, uint32_t lastColumn);

    bool setCellFontColor(uint32_t row, uint32_t column, uint8_t red, uint8_t green, uint8_t blue, uint8_t alpha = 255);
    bool setCellBackgroundColor(uint32_t row, uint32_t column, uint8_t red, uint8_t green, uint8_t blue, uint8_t alpha = 255);
    bool setCellFontSize(uint32_t row, uint32_t column, uint16_t size);
    bool setCellFontBold(uint32_t row, uint32_t column, bool bold);
    bool setCellFontItalic(uint32_t row, uint32_t column, bool italic);
    bool setCellFontUnderline(uint32_t row, uint32_t column, bool underline);
    bool setCellAlignment(uint32_t row, uint32_t column, const std::string& horizontal, const std::string& vertical);

    bool setColumnWidth(uint32_t column, double width);
    bool setRowHeight(uint32_t row, double height);
    uint32_t columnCount() const;
    uint32_t rowCount() const;

    template<typename T>
    bool setRowData(uint32_t rowNumber, const std::vector<T>& data);

    template<typename T>
    std::vector<T> getRowData(uint32_t rowNumber);

    template<typename T>
    void setColumnData(uint16_t columnNumber, const std::vector<T>& data);

    template<typename T>
    std::vector<T> getColumnData(uint16_t columnNumber);

    std::vector<std::vector<OpenXLSX::XLCellValue>> getRangeValues(uint32_t firstRow, uint32_t firstColumn, uint32_t lastRow, uint32_t lastColumn);

    bool setRangeValues(uint32_t firstRow, uint32_t firstColumn, const std::vector<std::vector<XLCellValue>>& values);

private:
    OpenXLSX::XLDocument m_document;
    OpenXLSX::XLWorkbook m_workbook;
    OpenXLSX::XLWorksheet m_currentSheet;
    bool m_isOpen;
};

} // namespace ExcelWrapper

template<typename T>
void ExcelWrapper::ExcelOperator::setCellValue(const std::string& cellReference, const T& value) {
    if (!m_isOpen) {
        return;
    }
    m_currentSheet.cell(cellReference).value() = value;
}

template<typename T>
T ExcelWrapper::ExcelOperator::getCellValue(const std::string& cellReference) {
    if (!m_isOpen) {
        return T();
    }
    return m_currentSheet.cell(cellReference).value().get<T>();
}

template<typename T>
bool ExcelWrapper::ExcelOperator::setRowData(uint32_t rowNumber, const std::vector<T>& data) {
    if (!m_isOpen) {
        return false;
    }
    for (size_t i = 0; i < data.size(); ++i) {
        m_currentSheet.cell(rowNumber, static_cast<uint16_t>(i + 1)).value() = data[i];
    }
    return true;
}

template<typename T>
std::vector<T> ExcelWrapper::ExcelOperator::getRowData(uint32_t rowNumber) {
    std::vector<T> rowData;
    if (!m_isOpen) {
        return rowData;
    }
    for (uint16_t col = 1; col <= 100; ++col) {
        OpenXLSX::XLCell cell = m_currentSheet.cell(rowNumber, col);
        if (!cell) {
            break;
        }
        if (cell.value().type() == OpenXLSX::XLValueType::Empty) {
            break;
        }
        rowData.push_back(cell.value().get<T>());
    }
    return rowData;
}

template<typename T>
void ExcelWrapper::ExcelOperator::setColumnData(uint16_t columnNumber, const std::vector<T>& data) {
    if (!m_isOpen) {
        return;
    }
    for (size_t i = 0; i < data.size(); ++i) {
        m_currentSheet.cell(static_cast<uint32_t>(i + 1), columnNumber).value() = data[i];
    }
}

template<typename T>
std::vector<T> ExcelWrapper::ExcelOperator::getColumnData(uint16_t columnNumber) {
    std::vector<T> columnData;
    if (!m_isOpen) {
        return columnData;
    }
    for (uint32_t row = 1; row <= 100; ++row) {
        OpenXLSX::XLCell cell = m_currentSheet.findCell(row, columnNumber);
        if (!cell) {
            break;
        }
        if (cell.value().type() == OpenXLSX::XLValueType::Empty) {
            break;
        }
        columnData.push_back(cell.value().get<T>());
    }
    return columnData;
}

#endif // EXCEL_OPERATOR_H