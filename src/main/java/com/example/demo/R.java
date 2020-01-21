package com.example.demo;

import org.apache.poi.ss.usermodel.Cell;
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
        tor.setRowStyle(fromr.getRowStyle());

        copyMergedRegions(from,to,sheet);

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

}
