package exercise;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exercise {

    public static void main(String[] args) {
        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja = libro.createSheet("Excercise");

        Map<String, Object[]> dato = new TreeMap<>();
        dato.put("1", new Object[]{"NAME", "LASTNAME", "EMAIL", "PASSWORD", "COMPANY", "ADDRESS", "CITY", "ZIP_CODE", "MOBILE_PHONE"});
        dato.put("2", new Object[]{"John", "Adwards", "JohnA@email.com", "JohnAwd&%05", "IT Enterprice", "Sn. Calle, S/N, Centro","Primary Town", "12345", "99 12 45 78 23"});

        Set<String> primaria = dato.keySet();
        int num = 0;
        for (String key : primaria) {
            Row fila = hoja.createRow(num++);
            Object[] objArr = dato.get(key);
            int celda = 0;
            for (Object obj : objArr) {

                Cell cell = fila.createCell(celda++);
                if (obj instanceof String){cell.setCellValue((String)obj);}
                else if (obj instanceof Integer){cell.setCellValue((Integer)obj);}
            }
        }
        colorRow(hoja);
        try {
            FileOutputStream salida = new FileOutputStream(new File("Excercise.xlsx"));
            libro.write(salida);
            salida.close();
            System.out.println("Documento creado con exito.");
        } catch (Exception e) {
        }
    }
    static void colorRow(Sheet sheet) {
    SheetConditionalFormatting colorRow = sheet.getSheetConditionalFormatting();
 
    ConditionalFormattingRule rule1 = colorRow.createConditionalFormattingRule("MOD(ROW(),2)");
    PatternFormatting fill1 = rule1.createPatternFormatting();
    fill1.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
    fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
 
    CellRangeAddress[] regions = {
            CellRangeAddress.valueOf("A1:Z100")
    };
 
    colorRow.addConditionalFormatting(regions, rule1);
}

}
