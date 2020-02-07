package com.example.demo;

import org.apache.poi.ss.usermodel.Sheet;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.net.MalformedURLException;
import java.net.URL;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

/**
 * Created by Artem on 07.02.2020.
 */
public class DollarRate {

    private Sheet sheetInterface;

    public DollarRate(Sheet sheetInterface) throws ParseException {
        this.sheetInterface = sheetInterface;
    }

    double getDollarRate() throws ParserConfigurationException, SAXException, ParseException, IOException {
        double malualDollarRate = getManualDollarRate();
        if (malualDollarRate != 0) {
            return malualDollarRate;
        } else {
            return getDollarRateFromTSBRF();
        }
    }

    private String getDateValuteS() {
        Date dateValute = sheetInterface.getRow(16).getCell(6).getDateCellValue();
        DateFormat dateFormatValute = new SimpleDateFormat("dd/MM/yyyy");
        String dateValuteS = dateFormatValute.format(dateValute);
        return dateValuteS;
    }

    private double getDollarRateFromTSBRF() throws ParserConfigurationException, IOException, SAXException, ParseException {
        String dateValuteS = getDateValuteS();

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
        return dollarValue;
    }

    private double getManualDollarRate() {
        double dollarRateFromCell = sheetInterface.getRow(18).getCell(6).getNumericCellValue();
        return dollarRateFromCell;
    }

}
