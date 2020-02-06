package com.example.demo;

import com.ibm.icu.text.RuleBasedNumberFormat;
import org.decimal4j.util.DoubleRounder;

import java.util.Locale;

/**
 * Created by Artem on 06.02.2020.
 */
public class Spellout {

    String format(double d) {
        d = DoubleRounder.round(d,2);
        System.out.println(d);
        int d1 = (int) d;
        int d2 = (int) ((d-d1)*100);
        RuleBasedNumberFormat nf = new RuleBasedNumberFormat(Locale.forLanguageTag("ru"),
                RuleBasedNumberFormat.SPELLOUT);
        return nf.format(d1) + " " + getFormR(d1) + " " + d2 + " " + getFormK(d2);
    }

    int getForm(int i) {
        int r=2;
        int l = i%10;
        switch (l) {
            case 1:
                r=0;
                break;
            case 2:
            case 3:
            case 4:
                r=1;
                break;
        }
        return r;
    }

    String getFormR(int i) {
        int f = getForm(i);
        switch (f) {
            case 0:
                return "рубль";
            case 1:
                return "рубля";
            case 2:
                return "рублей";
        }
        return "";
    }

    String getFormK(int i) {
        int f = getForm(i);
        switch (f) {
            case 0:
                return "копейка";
            case 1:
                return "копейки";
            case 2:
                return "копеек";
        }
        return "";
    }

}
