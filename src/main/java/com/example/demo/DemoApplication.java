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

		//R.copy(32,33,sheet);

//		for (int i=54;i>=33;i--) {
//			R.copy(i,i+1,sheet);
//		}

		CellRangeAddressUList cellRangeAddressUList = new CellRangeAddressUList(sheet);
		CellRangeAddressUList.removeAllMergedRegions(sheet);

		cellRangeAddressUList.copyTo(sheet);

		workbook.write(new FileOutputStream("Template Output.xlsx"));
	}

}
