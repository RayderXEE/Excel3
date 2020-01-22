package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;

import java.util.Date;

public class C {

    String stringValue;
    boolean booleanValue;
    Date dateValue;
    double numericValue;
    RichTextString richStringValue;
    String formula;

    int type;
    CellStyle style;
    CellAddress address;
    int columnIndex;
    int rowIndex;

    void copyTo(int rowIndex, int columnIndex, Sheet sheet) {
        Cell cell = sheet.getRow(rowIndex).createCell(columnIndex);

        cell.setCellType(this.type);
        cell.setCellStyle(this.style);

        switch (this.type) {
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue(this.stringValue);
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue(this.booleanValue);
                break;
            case Cell.CELL_TYPE_NUMERIC:
                cell.setCellValue(this.numericValue);
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellFormula(this.formula);
                break;
            case Cell.CELL_TYPE_BLANK:
                break;
        }
    }

    C copyFrom(Cell cell) {
        this.type = cell.getCellType();
        this.style = cell.getCellStyle();
        this.address = cell.getAddress();
        this.columnIndex = cell.getColumnIndex();
        this.rowIndex = cell.getRowIndex();

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                this.stringValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                this.booleanValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                this.numericValue = cell.getNumericCellValue();
                break;
            case Cell.CELL_TYPE_FORMULA:
                this.formula = cell.getCellFormula();
                break;
        }

        return this;
    }

    String getValue() {
        switch (this.type) {
            case Cell.CELL_TYPE_STRING:
                return this.stringValue;

            case Cell.CELL_TYPE_BOOLEAN:
                return String.valueOf(this.booleanValue);

            case Cell.CELL_TYPE_NUMERIC:
                return String.valueOf(this.numericValue);

            case Cell.CELL_TYPE_FORMULA:
                return this.formula;

        }
        return null;
    }

}
