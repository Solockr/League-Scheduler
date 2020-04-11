/*
 * Creator:Solomon Lo
 * Version: 2.0
 * Read comments in PopulateListOfFields class as well!
 * Please run the ResetListOfFields class first! Also, run ResetListOfFields every time you finish running this class.
 * Make sure you run this as a Java program.
 * First-Time Setup for this class: Change the CSV arrays in the datatypes variable to reflect the available fields. I find
 * it easier to copy and paste existing arrays, changing the Field Name and Size only.
 * The Strings for Size can only be "Small", "Medium","Large", or "Full" Small has only 1 array, but medium has 2 to reflect
 * how multiple teams can play at once on it, and large has 3, etc.
 */

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ResetListOfFields {
	
    private static final String FILE_NAME = "MyFirstExcel.xlsx";
    public static void main(String[] args) {
    	
    	

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet Of Fields");
        Object[][] datatypes = {
                {"Field Name", "Size", "MW AM Slot 1", "MW AM Slot 2", "MW PM Slot 1", "MW PM Slot 2", "THR AM Slot 1", "THR AM Slot 1", "THR PM Slot 1", "THR PM Slot 2"},
                {"MacFarland 1", "Small", "", "", "", "", "", "", "", ""},
                {"Empire Oaks 1", "Medium", "", "", "", "", "", "", "", ""},
                {"Empire Oaks 2", "Medium", "", "", "", "", "", "", "", ""},
                {"Imaginary 1", "Medium", "", "", "", "", "", "", "", ""},
                {"Imaginary 2", "Medium", "", "", "", "", "", "", "", ""},
                {"Lembi 1", "Large", "", "", "", "", "", "", "", ""},
                {"Lembi 2", "Large", "", "", "", "", "", "", "", ""},
                {"Lembi 3", "Large", "", "", "", "", "", "", "", ""},
                {"Kemp 1", "Full", "", "", "", "", "", "", "", ""},
                {"Kemp 2", "Full", "", "", "", "", "", "", "", ""}, 
                {"Kemp 3", "Full", "", "", "", "", "", "", "", ""}, 
        };

        int rowNum = 0;
        System.out.println("Creating an unfilled Excel file of the fields.");

        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");
    }
}