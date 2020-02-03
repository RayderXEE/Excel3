package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.decimal4j.util.DoubleRounder;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

@SpringBootApplication
public class DemoApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		Workbook workbookTemplate = new XSSFWorkbook(new FileInputStream("Template.xlsx"));
		Sheet sheetTemplate = workbookTemplate.getSheetAt(0);

		Workbook workbookOrder = new XSSFWorkbook(new FileInputStream("Order.xlsx"));
		Sheet sheetOrder = workbookOrder.getSheetAt(0);

		Workbook workbookPo = new XSSFWorkbook(new FileInputStream("Po.xlsx"));
		Sheet sheetPo = workbookPo.getSheet("Annex");

		ArrayList<R> rs = new ArrayList<>();
		for (int i=20;;i++) {
			Row row = sheetOrder.getRow(i);
			Cell cell = row.getCell(0);
			if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				break;
			}
			rs.add(new R().copyFrom(i,sheetOrder));
		}

        ArrayList<R> pors = new ArrayList<>();
        HashMap<String,String> pohm = new HashMap<>();
        for (int i=10;;i++) {
            Row row = sheetPo.getRow(i);
            Cell cell = row.getCell(2);
            //System.out.println(cell);
            if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
                break;
            }
            R r = new R().copyFrom(i,sheetPo);
			//System.out.println(r.cs.get(2).stringValue + " " + r.cs.get(4).stringValue);
			pohm.put(r.cs.get(2).stringValue, r.cs.get(4).stringValue);
			pors.add(r);
        }

        String purchaseOrderNo = sheetOrder.getRow(6).getCell(1).getStringCellValue();
        String expectedArrivalDate = sheetOrder.getRow(8).getCell(7).getStringCellValue();

		//System.out.println(expectedArrivalDate);

		R.shift(33,53,rs.size()-1, sheetTemplate);

		for (int i=0; i<rs.size()-1; i++) {
			R.shift(32+i,32+i,1, sheetTemplate);
		}

		HashMap<Integer,Integer> links = new HashMap<>();
		links.put(6, 36);
		links.put(0, 16);
		links.put(7, 39);
		links.put(8, 42);
		//links.put(8, 42);

		for (R r :
				rs) {
			C c = r.cs.get(6);
			c.type = Cell.CELL_TYPE_NUMERIC;
		}

        CellStyle style = workbookTemplate.createCellStyle();
		style.setWrapText(true);

		for (int i=0;i<rs.size();i++) {
			rs.get(i).copyTo(32+i,sheetTemplate, links, null);
			Row row = sheetTemplate.getRow(32+i);
            row.setHeight((short)750);
			row.setRowStyle(style);
			String name = pohm.get(rs.get(i).cs.get(0).stringValue);
            //System.out.println(name);
            row.createCell(0).setCellValue(i+1);
            row.createCell(3).setCellValue(name);
            row.getCell(3).setCellStyle(style);
            row.createCell(19).setCellValue("шт");
            row.createCell(22).setCellValue(796);
			row.createCell(45).setCellValue("20%");
			double priceWithoutVAT = row.getCell(42).getNumericCellValue();
			double vat = priceWithoutVAT/100*20;
			double priceWithVAT = priceWithoutVAT+vat;
			row.createCell(48).setCellValue(DoubleRounder.round( vat,2));
			row.createCell(52).setCellValue(DoubleRounder.round( priceWithVAT,2));
            sheetTemplate.getRow(21).createCell(7).setCellValue("Поставка товаров согласно контракту на поставку " +
                    "оборудования № "+purchaseOrderNo+" от "+expectedArrivalDate);
            sheetTemplate.getRow(21).setHeight((short)-1);
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            String currentDate = dateFormat.format(new Date());
            sheetTemplate.getRow(26).createCell(28).setCellValue(currentDate);
            sheetTemplate.getRow(59).createCell(7).setCellValue(currentDate);
		}

        //  Formulas
        Row rowTotal = sheetTemplate.getRow(32+rs.size());
		rowTotal.createCell(36).setCellFormula("SUM(AK33:AK"+(32+rs.size())+")");
        rowTotal.createCell(42).setCellFormula("SUM(AQ33:AQ"+(32+rs.size())+")");
        rowTotal.createCell(48).setCellFormula("SUM(AW33:AW"+(32+rs.size())+")");
        rowTotal.createCell(52).setCellFormula("SUM(BA33:BA"+(32+rs.size())+")");

        String printArea = workbookTemplate.getPrintArea(0);
        //System.out.println(printArea);
        workbookTemplate.setPrintArea(0,0,55,0,52+rs.size());
        //String[] printAreaSplit = printArea.split("$");
        //System.out.println(printAreaSplit[3]);

		workbookTemplate.write(new FileOutputStream("Template Output.xlsx"));
	}

}
