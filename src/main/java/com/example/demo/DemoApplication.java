package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.decimal4j.util.DoubleRounder;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileInputStream;
import java.io.FileOutputStream;
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

		for (int i=0;i<rs.size();i++) {
			rs.get(i).copyTo(32+i,sheetTemplate, links, null);
			Row row = sheetTemplate.getRow(32+i);
			String name = pohm.get(rs.get(i).cs.get(0).stringValue);
            //System.out.println(name);
            row.createCell(3).setCellValue(name);
			row.createCell(45).setCellValue("20%");
			double priceWithoutVAT = row.getCell(42).getNumericCellValue();
			double vat = priceWithoutVAT/100*20;
			double priceWithVAT = priceWithoutVAT+vat;
			row.createCell(48).setCellValue(DoubleRounder.round( vat,2));
			row.createCell(52).setCellValue(DoubleRounder.round( priceWithVAT,2));
		}

		workbookTemplate.write(new FileOutputStream("Template Output.xlsx"));
	}

}
