package com.example.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
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

    void copyTo(Cell cell) {
        cell.setCellType(this.type);
        cell.setCellStyle(this.style);

        switch (this.type) {
            case Cell.CELL_TYPE_STRING:
                cell.setCellValue(this.stringValue);
                //System.out.println(stringValue);
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                cell.setCellValue(this.booleanValue);
                break;
            case Cell.CELL_TYPE_NUMERIC:
                cell.setCellValue(this.numericValue);
                //System.out.println(numericValue);
                break;
            case Cell.CELL_TYPE_FORMULA:
                cell.setCellValue(this.formula);
                break;
        }
    }

    void copyFrom(Cell cell) {
        this.type = cell.getCellType();
        this.style = cell.getCellStyle();
        this.address = cell.getAddress();
        this.columnIndex = cell.getColumnIndex();
        this.rowIndex = cell.getRowIndex();

        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                this.stringValue = cell.getStringCellValue();
                //System.out.println(stringValue);
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                this.booleanValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                this.numericValue = cell.getNumericCellValue();
                //System.out.println(numericValue);
                break;
            case Cell.CELL_TYPE_FORMULA:
                this.formula = cell.getCellFormula();
                break;
        }
    }

}
