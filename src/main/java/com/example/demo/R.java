package com.example.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

public class R {

    static void copyMergedRegions(int from, int to, Sheet sheet) {
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress cellRangeAddress:
                mergedRegions) {
            int index = sheet.getMergedRegions().indexOf(cellRangeAddress);

            int fr = cellRangeAddress.getFirstRow();
            int lr = cellRangeAddress.getLastRow();
            int fc = cellRangeAddress.getFirstColumn();
            int lc = cellRangeAddress.getLastColumn();

            if (fr == to) {
                sheet.removeMergedRegion(index);
            }
        }

        mergedRegions = sheet.getMergedRegions();
        for (CellRangeAddress cellRangeAddress:
                mergedRegions) {
            int index = sheet.getMergedRegions().indexOf(cellRangeAddress);

            int fr = cellRangeAddress.getFirstRow();
            int lr = cellRangeAddress.getLastRow();
            int fc = cellRangeAddress.getFirstColumn();
            int lc = cellRangeAddress.getLastColumn();

            if (fr == from) {
                sheet.addMergedRegion(new CellRangeAddress(to,to,fc,lc));
            }
        }
    }

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
            Cell toc = tor.getCell(fromc.getColumnIndex());
            try {
                c.copyTo(toc);
            } catch (Exception e) {

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

}
