package com.example.demo;

import com.ibm.icu.text.RuleBasedNumberFormat;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.decimal4j.util.DoubleRounder;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.w3c.dom.*;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URL;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;

@SpringBootApplication
public class DemoApplication implements CommandLineRunner {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

    @Override
    public void run(String... strings) throws Exception {
        File directories = new File("directories");
        for (File file :
                directories.listFiles()) {
            String dirS = file.getPath();
            Dir dir = new Dir(dirS + "\\");
            dir.run();
        }
        //Dir dir = new Dir("");
        //dir.run();
    }

    //@Override
//	public void run(String... args) throws Exception {
//		Workbook workbookTemplate = new XSSFWorkbook(new FileInputStream("Template.xlsx"));
//		Sheet sheetTemplate = workbookTemplate.getSheetAt(0);
//
//		Workbook workbookOrder = new XSSFWorkbook(new FileInputStream("Order.xlsx"));
//		Sheet sheetOrder = workbookOrder.getSheetAt(0);
//
//		Workbook workbookPo = new XSSFWorkbook(new FileInputStream("Po.xlsx"));
//        Sheet sheetPo = workbookPo.getSheet("Annex");
//
//        Workbook workbookInterface = new XSSFWorkbook(new FileInputStream("Interface.xlsx"));
//        Sheet sheetInterface = workbookInterface.getSheet("Sheet1");
//
//		ArrayList<R> rs = new ArrayList<>();
//		for (int i=20;;i++) {
//			Row row = sheetOrder.getRow(i);
//			Cell cell = row.getCell(0);
//			if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
//				break;
//			}
//			rs.add(new R().copyFrom(i,sheetOrder));
//		}
//
//        ArrayList<R> pors = new ArrayList<>();
//        HashMap<String,String> pohm = new HashMap<>();
//        for (int i=10;;i++) {
//            Row row = sheetPo.getRow(i);
//            Cell cell = row.getCell(2);
//            //System.out.println(cell);
//            if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
//                break;
//            }
//            R r = new R().copyFrom(i,sheetPo);
//			//System.out.println(r.cs.get(2).stringValue + " " + r.cs.get(4).stringValue);
//			pohm.put(r.cs.get(2).stringValue, r.cs.get(4).stringValue);
//			pors.add(r);
//        }
//
//        String purchaseOrderNo = sheetOrder.getRow(6).getCell(1).getStringCellValue();
//        String expectedArrivalDate = sheetOrder.getRow(8).getCell(7).getStringCellValue();
//
//		//System.out.println(expectedArrivalDate);
//
//		R.shift(33,53,rs.size()-1, sheetTemplate);
//
//		for (int i=0; i<rs.size()-1; i++) {
//			R.shift(32+i,32+i,1, sheetTemplate);
//		}
//
//		HashMap<Integer,Integer> links = new HashMap<>();
//		links.put(6, 36);
//		links.put(0, 16);
//		links.put(7, 39);
//		links.put(8, 42);
//		//links.put(8, 42);
//
//		for (R r :
//				rs) {
//			C c = r.cs.get(6);
//			c.type = Cell.CELL_TYPE_NUMERIC;
//		}
//
//        CellStyle style = workbookTemplate.createCellStyle();
//		style.setWrapText(true);
//
//		DollarRate dollarRate = new DollarRate(sheetInterface);
//        double dollarValue = dollarRate.getDollarRate();
//
//        System.out.println(dollarValue);
//
//        double totalCount = 0;
//        double totalSumWithoutVAT = 0;
//        double totalVAT = 0;
//        double totalSum = 0;
//
//        CellStyle orderStyle = workbookTemplate.createCellStyle();
//        orderStyle.setBorderLeft(CellStyle.BORDER_THIN);
//        orderStyle.setBorderTop(CellStyle.BORDER_THIN);
//        orderStyle.setBorderRight(CellStyle.BORDER_THIN);
//        orderStyle.setBorderBottom(CellStyle.BORDER_THIN);
//        orderStyle.setWrapText(true);
//
//		for (int i=0;i<rs.size();i++) {
//			rs.get(i).copyTo(32+i,sheetTemplate, links, null);
//			Row row = sheetTemplate.getRow(32+i);
//            row.setHeight((short)750);
//			row.setRowStyle(style);
//			String name = pohm.get(rs.get(i).cs.get(0).stringValue);
//            //System.out.println(name);
//
//            Cell numberCell = row.createCell(0);
//            numberCell.setCellValue(i+1);
//            //numberCell.setCellStyle(orderStyle);
//
//            row.createCell(3).setCellValue(name);
//            row.getCell(3).setCellStyle(style);
//            row.createCell(19).setCellValue("шт");
//            row.createCell(22).setCellValue(796);
//
//            double count = Double.valueOf(rs.get(i).cs.get(6).stringValue);
//
//            double priceDollar = Double.valueOf(rs.get(i).cs.get(7).stringValue);
//            double priceRub = priceDollar * dollarValue;
//            priceRub = DoubleRounder.round(priceRub,2);
//            //System.out.println(priceDollar);
//            //System.out.println(priceRub);
//            row.createCell(39).setCellValue(priceRub);
//
//            double sum = priceRub * count;
//            row.createCell(42).setCellValue(sum);
//
//			row.createCell(45).setCellValue("20%");
//			double priceWithoutVAT = row.getCell(42).getNumericCellValue();
//			double vat = priceWithoutVAT/100*20;
//			double priceWithVAT = priceWithoutVAT+vat;
//
//			totalCount += count;
//			totalSumWithoutVAT += priceWithoutVAT;
//			totalVAT += vat;
//			totalSum += priceWithVAT;
//
//			row.createCell(48).setCellValue(DoubleRounder.round( vat,2));
//			row.createCell(52).setCellValue(DoubleRounder.round( priceWithVAT,2));
//
//			//row.setRowStyle(orderStyle);
////            for (Cell cell :
////                    row) {
////                cell.setCellStyle(orderStyle);
////            }
//
//            for (int c=0;c<56;c++) {
//                Cell cell = row.getCell(c);
//                if (cell == null) {
//                    cell = row.createCell(c);
//                }
//                cell.setCellStyle(orderStyle);
//            }
//
//        }
//
//        sheetTemplate.getRow(21).createCell(7).setCellValue("Поставка товаров согласно контракту на поставку " +
//                "оборудования № "+purchaseOrderNo+" от "+expectedArrivalDate);
//        sheetTemplate.getRow(21).setHeight((short)-1);
//        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
//        String currentDate = dateFormat.format(new Date());
//        sheetTemplate.getRow(26).createCell(28).setCellValue(currentDate);
//        sheetTemplate.getRow(59).createCell(7).setCellValue(currentDate);
//
//        double nDoc = sheetInterface.getRow(6).getCell(6).getNumericCellValue();
//        sheetTemplate.getRow(26).createCell(22).setCellValue(nDoc);
//
//        //  Formulas
//        Row rowTotal = sheetTemplate.getRow(32+rs.size());
////		rowTotal.createCell(36).setCellFormula("SUM(AK33:AK"+(32+rs.size())+")");
////        rowTotal.createCell(42).setCellFormula("SUM(AQ33:AQ"+(32+rs.size())+")");
////        rowTotal.createCell(48).setCellFormula("SUM(AW33:AW"+(32+rs.size())+")");
////        rowTotal.createCell(52).setCellFormula("SUM(BA33:BA"+(32+rs.size())+")");
//
//        totalCount = DoubleRounder.round(totalCount, 2);
//        totalSumWithoutVAT = DoubleRounder.round(totalSumWithoutVAT, 2);
//        totalVAT = DoubleRounder.round(totalVAT, 2);
//        totalSum = DoubleRounder.round(totalSum, 2);
//
//        rowTotal.createCell(36).setCellValue(totalCount);
//        rowTotal.createCell(42).setCellValue(totalSumWithoutVAT);
//        rowTotal.createCell(48).setCellValue(totalVAT);
//        rowTotal.createCell(52).setCellValue(totalSum);
//
//        for (int c=31;c<56;c++) {
//            Cell cell = rowTotal.getCell(c);
//            if (cell == null) {
//                cell = rowTotal.createCell(c);
//            }
//            cell.setCellStyle(orderStyle);
//        }
//
//        Row rowTotalSum = sheetTemplate.getRow(44+rs.size());
//        rowTotalSum.setHeight((short) 600);
//        Cell cellTotalSum = rowTotalSum.createCell(0);
//
//        RuleBasedNumberFormat nf = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"),
//                RuleBasedNumberFormat.SPELLOUT);
//        //System.out.println(nf.format(1234567));
//        //String totalSunCuirsive = nf.format(totalSum);
//        String totalSumCuirsive = new Spellout().format(totalSum);
//
//        cellTotalSum.setCellValue(totalSumCuirsive);
//        CellStyle cellStyleTotalSum = workbookTemplate.createCellStyle();
//        Font font = workbookTemplate.createFont();
//        font.setItalic(true);
//        cellStyleTotalSum.setAlignment(CellStyle.ALIGN_CENTER);
//        cellStyleTotalSum.setFont(font);
//        cellStyleTotalSum.setWrapText(true);
//        cellTotalSum.setCellStyle(cellStyleTotalSum);
//
//        sheetTemplate.getRow(43+rs.size()).setHeight((short)-1);
//
//        String printArea = workbookTemplate.getPrintArea(0);
//        //System.out.println(printArea);
//        workbookTemplate.setPrintArea(0,0,55,0,52+rs.size());
//        //String[] printAreaSplit = printArea.split("$");
//        //System.out.println(printAreaSplit[3]);
//
//		workbookTemplate.write(new FileOutputStream("Template Output.xlsx"));
//	}

}
