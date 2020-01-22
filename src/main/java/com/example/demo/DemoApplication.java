package com.example.demo;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

		ArrayList<R> rs = new ArrayList<>();

		for (int i=20;;i++) {
			Row row = sheetOrder.getRow(i);
			Cell cell = row.getCell(0);
			if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
				break;
			}
			rs.add(new R().copyFrom(i,sheetOrder));
		}

		for (R r :
				rs) {
			String s = r.cs.get(1).getValue();
			//System.out.println(s);
		}

		R.shift(33,53,7, sheetTemplate);



		for (int i=0; i<5; i++) {
			R.shift(32+i,32+i,1, sheetTemplate);
		}

		HashMap<Integer, Integer> links = new HashMap<>();
		links.put(6, 36);
		links.put(0, 16);

		rs.get(0).copyToUsingLinks(33,sheetTemplate, links);

		workbookTemplate.write(new FileOutputStream("Template Output.xlsx"));
	}

}
