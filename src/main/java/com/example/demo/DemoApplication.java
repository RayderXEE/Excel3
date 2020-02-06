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
	public void run(String... args) throws Exception {
		Workbook workbookTemplate = new XSSFWorkbook(new FileInputStream("Template.xlsx"));
		Sheet sheetTemplate = workbookTemplate.getSheetAt(0);

		Workbook workbookOrder = new XSSFWorkbook(new FileInputStream("Order.xlsx"));
		Sheet sheetOrder = workbookOrder.getSheetAt(0);

		Workbook workbookPo = new XSSFWorkbook(new FileInputStream("Po.xlsx"));
        Sheet sheetPo = workbookPo.getSheet("Annex");

        Workbook workbookInterface = new XSSFWorkbook(new FileInputStream("Interface.xlsx"));
        Sheet sheetInterface = workbookInterface.getSheet("Sheet1");

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

        // Get Date for Valute
        Date dateValute = sheetInterface.getRow(16).getCell(6).getDateCellValue();
        DateFormat dateFormatValute = new SimpleDateFormat("dd/MM/yyyy");
        String dateValuteS = dateFormatValute.format(dateValute);
        //System.out.println(dateValuteS);

        System.setProperty("http.maxRedirects", "500");
        DocumentBuilderFactory dbf = DocumentBuilderFactory.newInstance();
        DocumentBuilder db = dbf.newDocumentBuilder();
        Document doc = db.parse(new URL("http://www.cbr.ru/scripts/XML_daily.asp?date_req="+dateValuteS).openStream());

        double dollarValue = 0;
        NodeList valutes = doc.getElementsByTagName("Valute");
        //System.out.println(valutes.getLength());
        for (int i=0;i<valutes.getLength();i++) {
            Node valute = valutes.item(i);
            NamedNodeMap valuteAttributes = valute.getAttributes();
            String valuteID = valuteAttributes.item(0).getTextContent();
            if (valuteID.equals("R01235")) {
                String sDollarValue = valute.getLastChild().getTextContent();

                NumberFormat format = NumberFormat.getInstance(Locale.FRANCE);
                Number number = format.parse(sDollarValue);
                dollarValue = number.doubleValue();
            }
        }

        System.out.println(dollarValue);

        double totalSum = 0;

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

            double count = Double.valueOf(rs.get(i).cs.get(6).stringValue);

            double priceDollar = Double.valueOf(rs.get(i).cs.get(7).stringValue);
            double priceRub = priceDollar * dollarValue;
            priceRub = DoubleRounder.round(priceRub,2);
            //System.out.println(priceDollar);
            //System.out.println(priceRub);
            row.createCell(39).setCellValue(priceRub);

            double sum = priceRub * count;
            row.createCell(42).setCellValue(sum);

			row.createCell(45).setCellValue("20%");
			double priceWithoutVAT = row.getCell(42).getNumericCellValue();
			double vat = priceWithoutVAT/100*20;
			double priceWithVAT = priceWithoutVAT+vat;
			totalSum += priceWithVAT;
			row.createCell(48).setCellValue(DoubleRounder.round( vat,2));
			row.createCell(52).setCellValue(DoubleRounder.round( priceWithVAT,2));
		}

        sheetTemplate.getRow(21).createCell(7).setCellValue("Поставка товаров согласно контракту на поставку " +
                "оборудования № "+purchaseOrderNo+" от "+expectedArrivalDate);
        sheetTemplate.getRow(21).setHeight((short)-1);
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        String currentDate = dateFormat.format(new Date());
        sheetTemplate.getRow(26).createCell(28).setCellValue(currentDate);
        sheetTemplate.getRow(59).createCell(7).setCellValue(currentDate);

        double nDoc = sheetInterface.getRow(6).getCell(6).getNumericCellValue();
        sheetTemplate.getRow(26).createCell(22).setCellValue(nDoc);

        //  Formulas
        Row rowTotal = sheetTemplate.getRow(32+rs.size());
		rowTotal.createCell(36).setCellFormula("SUM(AK33:AK"+(32+rs.size())+")");
        rowTotal.createCell(42).setCellFormula("SUM(AQ33:AQ"+(32+rs.size())+")");
        rowTotal.createCell(48).setCellFormula("SUM(AW33:AW"+(32+rs.size())+")");
        rowTotal.createCell(52).setCellFormula("SUM(BA33:BA"+(32+rs.size())+")");

        Row rowTotalSum = sheetTemplate.getRow(44+rs.size());
        rowTotalSum.setHeight((short) 600);
        Cell cellTotalSum = rowTotalSum.createCell(0);

        RuleBasedNumberFormat nf = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"),
                RuleBasedNumberFormat.SPELLOUT);
        //System.out.println(nf.format(1234567));
        //String totalSunCuirsive = nf.format(totalSum);
        String totalSumCuirsive = new Spellout().format(totalSum);

        cellTotalSum.setCellValue(totalSumCuirsive);
        CellStyle cellStyleTotalSum = workbookTemplate.createCellStyle();
        Font font = workbookTemplate.createFont();
        font.setItalic(true);
        cellStyleTotalSum.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleTotalSum.setFont(font);
        cellStyleTotalSum.setWrapText(true);
        cellTotalSum.setCellStyle(cellStyleTotalSum);

        sheetTemplate.getRow(43+rs.size()).setHeight((short)-1);

        String printArea = workbookTemplate.getPrintArea(0);
        //System.out.println(printArea);
        workbookTemplate.setPrintArea(0,0,55,0,52+rs.size());
        //String[] printAreaSplit = printArea.split("$");
        //System.out.println(printAreaSplit[3]);

		workbookTemplate.write(new FileOutputStream("Template Output.xlsx"));
	}

}
