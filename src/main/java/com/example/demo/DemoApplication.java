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
import java.util.List;

@SpringBootApplication
public class DemoApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {
		Workbook workbook = new XSSFWorkbook(new FileInputStream("Template.xlsx"));
		Sheet sheet = workbook.getSheetAt(0);

		//sheet.getRow(65).getCell(15).setCellValue("test");
		Row row = sheet.getRow(65);


		//R.shift(33,54,10,sheet);
//		CellRangeAddressUList cellRangeAddressUList = new CellRangeAddressUList(sheet);
//		cellRangeAddressUList.shift(1,55,2);
//		cellRangeAddressUList.copyTo(sheet);
		//sheet.shiftRows(1,55,2);
		workbook.setPrintArea(0, 0,70,0,70);

		sheet.setAutobreaks(true);

		sheet.setRowBreak(70);
		sheet.setColumnBreak(70);

		row.setRowStyle(sheet.getRow(32).getRowStyle());
		Cell cell = row.createCell(3);
		System.out.println(cell.getAddress());
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("test");
		System.out.println(cell.getStringCellValue());

		//R.shift(1,55,15,sheet);

		workbook.write(new FileOutputStream("Template Output.xlsx"));
	}

}
