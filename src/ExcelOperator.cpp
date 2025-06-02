#include "ExcelOperator.h"

namespace ExcelWrapper {

ExcelOperator::ExcelOperator() : m_isOpen(false) {
}

ExcelOperator::~ExcelOperator() {
    if (m_isOpen) {
        close();
    }
}

bool ExcelOperator::open(const std::string& filePath, std::vector<std::string>& sheetNames) {
    if (m_isOpen) {
        close();
    }
    m_document.open(filePath);
    m_workbook = m_document.workbook();

    for (uint32_t i = 1; i <= m_workbook.sheetCount(); ++i) {
        sheetNames.emplace_back(m_workbook.sheet(i).name());
    }

    m_currentSheet = m_workbook.worksheet(1);
    m_isOpen = true;
    return true;
}

bool ExcelOperator::create(const std::string& filePath) {
    if (m_isOpen) {
        close();
    }
    m_document.create(filePath, OpenXLSX::XLForceOverwrite);
    m_workbook = m_document.workbook();
    m_currentSheet = m_workbook.worksheet("Sheet1");
    m_isOpen = true;
    return true;
}

bool ExcelOperator::save() {
    if (!m_isOpen) {
        return false;
    }
    m_document.save();
    return true;
}

bool ExcelOperator::saveAs(const std::string& filePath) {
    if (!m_isOpen) {
        return false;
    }
    m_document.saveAs(filePath, OpenXLSX::XLForceOverwrite);
    return true;
}

bool ExcelOperator::close() {
    if (m_isOpen) {
        try {
            m_document.close();
            m_isOpen = false;
            return true;
        } catch (const std::exception& e) {
            return false;
        }
    }
    return true;
}

bool ExcelOperator::selectSheet(const std::string& sheetName) {
    if (!m_isOpen) {
        return false;
    }
    m_currentSheet = m_workbook.worksheet(sheetName);
    return true;
}

bool ExcelOperator::selectSheet(uint32_t sheetIndex) {
    if (!m_isOpen) {
        return false;
    }
    m_currentSheet = m_workbook.worksheet(sheetIndex);
    return true;
}

bool ExcelOperator::addSheet(const std::string& sheetName) {
    if (!m_isOpen) {
        return false;
    }
    m_workbook.addWorksheet(sheetName);
    return true;
}

bool ExcelOperator::deleteSheet(const std::string& sheetName) {
    if (!m_isOpen) {
        return false;
    }
    m_workbook.deleteSheet(sheetName);
    return true;
}

bool ExcelOperator::renameSheet(const std::string& oldName, const std::string& newName) {
    if (!m_isOpen) {
        return false;
    }
    m_workbook.worksheet(oldName).setName(newName);
    return true;
}

uint32_t ExcelOperator::sheetCount() const {
    if (!m_isOpen) {
        return 0;
    }
    return m_workbook.sheetCount();
}

std::string ExcelOperator::currentSheetName() const {
    if (!m_isOpen) {
        return "";
    }
    return m_currentSheet.name();
}

bool ExcelOperator::clearCell(uint32_t row, uint32_t column) {
    if (!m_isOpen) {
        return false;
    }
    m_currentSheet.cell(row, column).clear(0);
    return true;
}

bool ExcelOperator::mergeCells(uint32_t firstRow, uint32_t firstColumn, uint32_t lastRow, uint32_t lastColumn) {
    if (!m_isOpen) {
        return false;
    }
    OpenXLSX::XLCellReference topLeft(firstColumn, firstRow);
    OpenXLSX::XLCellReference bottomRight(lastColumn, lastRow);
    m_currentSheet.mergeCells(m_currentSheet.range(topLeft, bottomRight));
    return true;
}

bool ExcelOperator::unmergeCells(uint32_t firstRow, uint32_t firstColumn, uint32_t lastRow, uint32_t lastColumn) {
    if (!m_isOpen) {
        return false;
    }
    OpenXLSX::XLCellReference topLeft(firstColumn, firstRow);
    OpenXLSX::XLCellReference bottomRight(lastColumn, lastRow);
    m_currentSheet.unmergeCells(m_currentSheet.range(topLeft, bottomRight));
    return true;
}

bool ExcelOperator::setCellFontColor(uint32_t row, uint32_t column, uint8_t red, uint8_t green, uint8_t blue, uint8_t alpha) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        
        auto font = styles.fonts()[currentFormat.fontIndex()];
        auto newFont = font;
        newFont.setFontColor(OpenXLSX::XLColor(red, green, blue, alpha));
        auto newFontIndex = styles.fonts().create(newFont);
        
        auto newFormat = currentFormat;
        newFormat.setFontIndex(newFontIndex);
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        
        cell.setCellFormat(newFormatIndex);
        
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setCellBackgroundColor(uint32_t row, uint32_t column, uint8_t red, uint8_t green, uint8_t blue, uint8_t alpha) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        
        auto fill = styles.fills()[currentFormat.fillIndex()];
        auto newFill = fill;
        newFill.setBackgroundColor(OpenXLSX::XLColor(red, green, blue, alpha));
        auto newFillIndex = styles.fills().create(newFill);
        
        auto newFormat = currentFormat;
        newFormat.setFillIndex(newFillIndex);
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        
        cell.setCellFormat(newFormatIndex);
        
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setCellFontSize(uint32_t row, uint32_t column, uint16_t size) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        
        auto font = styles.fonts()[currentFormat.fontIndex()];
        auto newFont = font;
        newFont.setFontSize(size);
        auto newFontIndex = styles.fonts().create(newFont);
        
        auto newFormat = currentFormat;
        newFormat.setFontIndex(newFontIndex);
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        
        cell.setCellFormat(newFormatIndex);
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setCellFontBold(uint32_t row, uint32_t column, bool bold) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        
        auto font = styles.fonts()[currentFormat.fontIndex()];
        auto newFont = font;
        newFont.setBold(bold);
        auto newFontIndex = styles.fonts().create(newFont);
        
        auto newFormat = currentFormat;
        newFormat.setFontIndex(newFontIndex);
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        
        cell.setCellFormat(newFormatIndex);
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setCellFontItalic(uint32_t row, uint32_t column, bool italic) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        
        auto font = styles.fonts()[currentFormat.fontIndex()];
        auto newFont = font;
        newFont.setItalic(italic);
        auto newFontIndex = styles.fonts().create(newFont);
        
        auto newFormat = currentFormat;
        newFormat.setFontIndex(newFontIndex);
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        
        cell.setCellFormat(newFormatIndex);
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setCellFontUnderline(uint32_t row, uint32_t column, bool underline) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        
        auto font = styles.fonts()[currentFormat.fontIndex()];
        auto newFont = font;
        newFont.setUnderline(underline ? OpenXLSX::XLUnderlineSingle : OpenXLSX::XLUnderlineNone);
        auto newFontIndex = styles.fonts().create(newFont);
        
        auto newFormat = currentFormat;
        newFormat.setFontIndex(newFontIndex);
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        
        cell.setCellFormat(newFormatIndex);
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setCellAlignment(uint32_t row, uint32_t column, const std::string& horizontal, const std::string& vertical) {
    if (!m_isOpen || row < 1 || column < 1) {
        return false;
    }
    
    try {
        auto& styles = m_document.styles();
        auto cell = m_currentSheet.cell(row, column);
        
        auto currentFormatIndex = cell.cellFormat();
        auto currentFormat = styles.cellFormats()[currentFormatIndex];
        auto newFormat = currentFormat;
        
        auto alignment = newFormat.alignment();
        if (horizontal == "left") {
            alignment.setHorizontal(OpenXLSX::XLAlignmentStyle::XLAlignLeft);
        } else if (horizontal == "center") {
            alignment.setHorizontal(OpenXLSX::XLAlignmentStyle::XLAlignCenter);
        } else if (horizontal == "right") {
            alignment.setHorizontal(OpenXLSX::XLAlignmentStyle::XLAlignRight);
        }

        if (vertical == "top") {
            alignment.setVertical(OpenXLSX::XLAlignmentStyle::XLAlignTop);
        } else if (vertical == "bottom") {
            alignment.setVertical(OpenXLSX::XLAlignmentStyle::XLAlignBottom);
        }

        newFormat.setApplyAlignment(true);
        
        auto newFormatIndex = styles.cellFormats().create(newFormat);
        cell.setCellFormat(newFormatIndex);
        return true;
    } catch (const std::exception& e) {
        return false;
    }
}

bool ExcelOperator::setColumnWidth(uint32_t column, double width) {
    if (!m_isOpen) {
        return false;
    }
    m_currentSheet.column(column).setWidth(width);
    return true;
}

bool ExcelOperator::setRowHeight(uint32_t row, double height) {
    if (!m_isOpen) {
        return false;
    }
    m_currentSheet.row(row).setHeight(height);
    return true;
}

uint32_t ExcelOperator::columnCount() const {
    if (!m_isOpen) {
        return 0;
    }
    return m_currentSheet.columnCount();
}

uint32_t ExcelOperator::rowCount() const {
    if (!m_isOpen) {
        return 0;
    }
    return m_currentSheet.rowCount();
}

std::vector<std::vector<OpenXLSX::XLCellValue>> ExcelOperator::getRangeValues(uint32_t firstRow, uint32_t firstColumn, uint32_t lastRow, uint32_t lastColumn) {
    std::vector<std::vector<OpenXLSX::XLCellValue>> rangeData;
    if (!m_isOpen || firstRow > lastRow || firstColumn > lastColumn) {
        return rangeData;
    }

    for (uint32_t r = firstRow; r <= lastRow; ++r) {
        std::vector<OpenXLSX::XLCellValue> rowData;
        for (uint32_t c = firstColumn; c <= lastColumn; ++c) {
            OpenXLSX::XLCell cell = m_currentSheet.cell(r, c);
            if (cell) {
                rowData.push_back(cell.value());
            } else {
                rowData.push_back(OpenXLSX::XLCellValue());
            }
        }
        rangeData.push_back(rowData);
    }
    return rangeData;
}

bool ExcelOperator::setRangeValues(uint32_t firstRow, uint32_t firstColumn, const std::vector<std::vector<XLCellValue>>& values) {
    if (!m_isOpen) {
        return false;
    }

    for (size_t r = 0; r < values.size(); ++r) {
        for (size_t c = 0; c < values[r].size(); ++c) {
            OpenXLSX::XLCellReference cellRef(firstRow + r, firstColumn + c);
            setCellValue(cellRef.address(), values[r][c]);
        }
    }
    this->save();
    return true;
}

} // namespace ExcelWrapper
