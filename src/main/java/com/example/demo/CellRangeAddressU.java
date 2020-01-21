package com.example.demo;

import org.apache.poi.ss.util.CellRangeAddress;

public class CellRangeAddressU {

    int fr;
    int lr;
    int fc;
    int lc;

    CellRangeAddressU copyFrom(CellRangeAddress cellRangeAddress) {
        this.fr = cellRangeAddress.getFirstRow();
        this.lr = cellRangeAddress.getLastRow();
        this.fc = cellRangeAddress.getFirstColumn();
        this.lc = cellRangeAddress.getLastColumn();
        return this;
    }

    public String toString() {
        return "fr " + fr + "; lr " + lr + "; fc " + fc + "; lc " + lc;
    }

}
