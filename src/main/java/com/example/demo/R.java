package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

public class R {

    static void copy(int from, int to, Sheet sheet) {
        Row fromr = sheet.getRow(from);
        Row tor = sheet.getRow(to);

        int k = to-sheet.getLastRowNum();
        if(k>0) {
            for (int i=sheet.getLastRowNum()+1;i<=to;i++) {
                Row row = sheet.createRow(i);
                for (int i2=0;i2<=50;i2++) {
                    Cell cell = row.createCell(i2);
                    System.out.println(cell);
                }
            }
        }

        tor.setHeight(fromr.getHeight());
        tor.setRowStyle(fromr.getRowStyle());

        for (Cell fromc :
                fromr) {
            C c = new C();
            c.copyFrom(fromc);
            //Cell toc = tor.getCell(fromc.getColumnIndex());
            try {
                c.copyTo(to, fromc.getColumnIndex(), sheet);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    static void shift(int from, int to, int shiftValue, Sheet sheet) {
        CellRangeAddressUList cellRangeAddressUList = new CellRangeAddressUList(sheet);
        cellRangeAddressUList.shift(from, to, shiftValue);
        cellRangeAddressUList.copyTo(sheet);

        for (int i=to;i>=from;i--) {
            copy(i,i+shiftValue,sheet);
        }

    }

    HashMap<Integer,C> cs = new HashMap<>();

    R copyFrom(int rowIndex, Sheet sheet) {
        Row row = sheet.getRow(rowIndex);

        for (Cell cell :
                row) {
            C c = new C().copyFrom(cell);
            cs.put(c.columnIndex, c);
        }

        return this;
    }

    R copyTo(int rowIndex, Sheet sheet, HashMap<Integer,Integer> links, CellStyle style) {
        Row row = sheet.createRow(rowIndex);

        for (C c :
                cs.values()) {
            c.style = style;
            if (links != null) {
                if (links.get(c.columnIndex) != null) {
                    c.copyTo(rowIndex, links.get(c.columnIndex), sheet);
                }
            } else {
                c.copyTo(rowIndex, c.columnIndex, sheet);
            }
        }

        return this;
    }

    R copyTo(int rowIndex, Sheet sheet) {
        return copyTo(rowIndex, sheet, null, null);
    }

}
