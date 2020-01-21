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
        //ArrayList<CellRangeAddressU> cellRangeAddressUArrayList = new ArrayList<>();
        for (CellRangeAddress cellRangeAddress :
                cellRangeAddressList) {
            cellRangeAddressUArrayList.add(new CellRangeAddressU().copyFrom(cellRangeAddress));
        }

//        for (CellRangeAddressU cellRangeAddressU :
//                cellRangeAddressUArrayList) {
//            System.out.println(cellRangeAddressU);
//        }
    }

    void copyTo(Sheet sheet) {
        //List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions();
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
        for (CellRangeAddressU cellRangeAddressU :
                cellRangeAddressUArrayList) {
            if (cellRangeAddressU.fr >= from && cellRangeAddressU.fr <= to) {
                cellRangeAddressU.fr += shiftValue;
                cellRangeAddressU.lr += shiftValue;
            }
        }
    }

}
