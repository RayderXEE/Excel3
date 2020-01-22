package com.example.demo;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class CellRangeAddressUList {

    static void removeAllMergedRegions(Sheet sheet) {
        while (sheet.getMergedRegions().size()>0) {
            sheet.removeMergedRegion(0);
        }
    }

    public ArrayList<CellRangeAddressU> cellRangeAddressUArrayList = new ArrayList<>();

    public CellRangeAddressUList(Sheet sheet) {
        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions();
        for (CellRangeAddress cellRangeAddress :
                cellRangeAddressList) {
            cellRangeAddressUArrayList.add(new CellRangeAddressU().copyFrom(cellRangeAddress));
        }

    }

    void copyTo(Sheet sheet) {
        removeAllMergedRegions(sheet);

        for (CellRangeAddressU cellRangeAddressU :
                cellRangeAddressUArrayList) {
            int fr = cellRangeAddressU.fr;
            int lr = cellRangeAddressU.lr;
            int fc = cellRangeAddressU.fc;
            int lc = cellRangeAddressU.lc;

            sheet.addMergedRegion(new CellRangeAddress(fr,lr,fc,lc));
        }
    }

    void shift(int from, int to, int shiftValue) {
        ArrayList<CellRangeAddressU> newCellRangeAddressUArrayList = new ArrayList<>();

        for (int i=0; i<cellRangeAddressUArrayList.size(); i++) {
            CellRangeAddressU cellRangeAddressU = cellRangeAddressUArrayList.get(i);

            if (cellRangeAddressU.fr > to && cellRangeAddressU.fr <= to+shiftValue) {
                cellRangeAddressUArrayList.remove(cellRangeAddressU);
                i--;
            }

            if (cellRangeAddressU.fr >= from && cellRangeAddressU.fr <= to) {
                int fr = cellRangeAddressU.fr;
                int lr = cellRangeAddressU.lr;
                int fc = cellRangeAddressU.fc;
                int lc = cellRangeAddressU.lc;
                if (cellRangeAddressU.lr < from+shiftValue) {
                    newCellRangeAddressUArrayList.add(new CellRangeAddressU().setPosition(fr,lr,fc,lc));
                }

                cellRangeAddressU.fr += shiftValue;
                cellRangeAddressU.lr += shiftValue;
            }
        }

        for (CellRangeAddressU cellRangeAddressU :
                newCellRangeAddressUArrayList) {
            cellRangeAddressUArrayList.add(cellRangeAddressU);
        }

    }

}
