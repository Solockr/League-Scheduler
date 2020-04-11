/*
 * Creator:Solomon Lo
 * Version: 2.0
 * Read comments in ResetListOfFields class as well!
 * First-time setup: Change the file directories(where indicated with comments) to the Excel file of your available
 * soccer fields and the Excel file of the teacher preferences(also where indicated with comments).
 * Please run the ResetListOfFields class first! Also, run ResetListOfFields every time you finish running this class.
 * Make sure you run both as a Java program.
 * 
 */

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class PopulateListOfFields {
	
	
	public static Workbook wb;
	public static Sheet coachList;
	public static Workbook listOfFields;
	public static Sheet availableFields;
	public static ArrayList<Integer> columnNumbers;
	public static List<Integer> correspondingRows;
	public static List<Integer> modifiedCorrespondingRows;

	public PopulateListOfFields() throws IOException, EncryptedDocumentException, InvalidFormatException {   
}		
	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException {
        FileInputStream fsIP= new FileInputStream(new File("[Insert file directory here]")); //change this to Excel file directory with coach preferences!
        FileInputStream listOfFields = new FileInputStream(new File("[Insert file directory here]"))); //change this to Excel file directory with available soccer fields!
        XSSFWorkbook wb = new XSSFWorkbook(fsIP); //Access the workbook
        XSSFWorkbook things = new XSSFWorkbook(listOfFields);
        coachList = wb.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        availableFields = things.getSheetAt(0); //Access the worksheet, so that we can update / modify it.
        start();  
        fsIP.close(); //Close the InputStream
        listOfFields.close(); 
        FileOutputStream output_file =new FileOutputStream(new File("C:\\Users\\Solomon Lo\\eclipse-workspace\\Rough Draft\\CoachInformationList.xlsx"));  //change this to Excel file directory with available soccer fields!
        FileOutputStream updateList =new FileOutputStream(new File("C:\\Users\\Solomon Lo\\eclipse-workspace\\Rough Draft\\MyFirstExcel.xlsx"));  //change this to Excel file directory with available soccer fields!
        wb.write(output_file); //write changes
        things.write(updateList);  
        output_file.close(); 
        updateList.close();//close the stream    		
		
	}
	public static void start() throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException {

		int i = 1;
		while(i < coachList.getRow(0).getPhysicalNumberOfCells()) {
			String preferredTime = coachList.getRow(2).getCell(i).getRichStringCellValue().toString();
			String teacher = coachList.getRow(0).getCell(i).getRichStringCellValue().toString();
		    Row row = coachList.getRow(4);
		    String dayOfWeek = coachList.getRow(3).getCell(i).getRichStringCellValue().toString();
		    Cell desiredField = row.getCell(i);		        
			correspondingRows = findRow(availableFields, desiredField);
			modifiedCorrespondingRows = modifiedFindRow(coachList);
	        //System.out.println(findRow(availableFields, desiredField));
	        //System.out.println(desiredField.getRichStringCellValue().toString()+ preferredTime + teacher + dayOfWeek);
	        assign(desiredField.getRichStringCellValue().toString(), preferredTime, teacher, dayOfWeek);
		    //System.out.println("Assigned" + i);
	        i++; 
		}
		//System.out.println(columnNumbers);
		//System.out.println("end");
		
	}

	private static List findRow(Sheet sheet, Cell cell2) {
		List columnNumbers = new ArrayList<Integer>();
	    for (int r = 0; r < sheet.getLastRowNum(); r++) {
        Cell cell = sheet.getRow(r).getCell(0);
        	//System.out.println(cell2.toString().substring(0, temp));
    	
            if (cell.getRichStringCellValue().getString().trim().substring(0, 4).equals(cell2.toString().substring(0, 4))){
            	//System.out.println(cell2.getRichStringCellValue().getString().trim());
                columnNumbers.add(r);
            }
        }
	    return columnNumbers;
	}
	private static List modifiedFindRow(Sheet sheet) {
		List columnNumbers = new ArrayList<Integer>();
	    for (int r = 1; r < sheet.getRow(1).getLastCellNum(); r++) {
        Cell cell = sheet.getRow(1).getCell(r);
        String tempString = cell.getRichStringCellValue().getString().trim();
        //System.out.println(tempString + "is what we're converting");
        String tempSubString = tempString.substring(tempString.length()-1);
        Integer result = Integer.valueOf(tempSubString);
        //System.out.println("Our modified division is" + result);
        	//System.out.println(cell2.toString().substring(0, temp));
    	
            if ((result >= 6 ) && (result <= 8)){
        	    for (int m = 0; m < sheet.getLastRowNum(); m++) {
        	        Cell tempCell = sheet.getRow(m).getCell(1);
        	        	//System.out.println(cell2.toString().substring(0, temp));
        	    	
        	            if (tempCell.getRichStringCellValue().getString().equals("Small")){
        	                columnNumbers.add(r);
        	            }
        	        }
            }
            if ((result >= 9) && (result <= 10)){
        	    for (int m = 0; m < sheet.getLastRowNum(); m++) {
        	        Cell tempCell = sheet.getRow(m).getCell(1);
        	        	//System.out.println(cell2.toString().substring(0, temp));
        	    	
        	            if (tempCell.getRichStringCellValue().getString().equals("Medium")){
        	                columnNumbers.add(r);
        	            }
        	        }
            }
            if ((result >= 11) && (result <= 14)){
        	    for (int m = 0; m < sheet.getLastRowNum(); m++) {
        	        Cell tempCell = sheet.getRow(m).getCell(1);
        	        	//System.out.println(cell2.toString().substring(0, temp));
        	    	
        	            if (tempCell.getRichStringCellValue().getString().equals("Large")){
        	                columnNumbers.add(r);
        	            }
        	        }
            }
            if ((result >= 15) && (result <= 18)){
        	    for (int m = 0; m < sheet.getLastRowNum(); m++) {
        	        Cell tempCell = sheet.getRow(m).getCell(1);
        	        	//System.out.println(cell2.toString().substring(0, temp));
        	    	
        	            if (tempCell.getRichStringCellValue().getString().equals("Full")){
        	                columnNumbers.add(r);
        	            }
        	        }
            }
        }
	    return columnNumbers;
	}
	public static void assign(String desiredField, String preferredTime, String coachToPlace, String dayOfWeek) {
		//Testing for all preferences.
		for(int n = 0; n < correspondingRows.size(); n++) {
			int tempRow = (int) correspondingRows.get(n);
			Row examine = availableFields.getRow(tempRow); //Gets the row of the field and looks for an empty box
			int slot = 0;
			if(dayOfWeek.equals("MW")) {
				slot = 2;
			}
			if(dayOfWeek.equals("THR")) {
				slot = 6;  //This should be 6
			}
			if(preferredTime.equals("PM")) {
				slot+=2;
			}
			for(int z = slot; z < slot + 2; z++) {
				if (examine.getCell(z) == null || examine.getCell(z).toString().equals("")) {					
					Cell cell = examine.getCell(z);
					//System.out.println("Got Cell");
					cell.setCellValue(coachToPlace);
					//System.out.println(coachToPlace);
					return;
				}
			}
			//Testing for only time and Day of Week, not caring about the field, but still placing them in correct size.
			for(int x = 0; x < modifiedCorrespondingRows.size(); x++) {
				int tempRow1 = (int) modifiedCorrespondingRows.get(x);
				Row examine1 = availableFields.getRow(tempRow1); //Gets the row of the field and looks for an empty box
				int slot1 = 0;
				if(dayOfWeek.equals("MW")) {
					slot1 = 2;
				}
				if(dayOfWeek.equals("THR")) {
					slot1 = 6;  //This should be 6
				}
				if(preferredTime.equals("PM")) {
					slot1+=2;
				}
				for(int z = slot1; z < slot1 + 2; z++) {
					if (examine1.getCell(z) == null || examine1.getCell(z).toString().equals("")) {					
						Cell cell = examine1.getCell(z);
						//System.out.println("Got MODIFIED Cell");
						//System.out.println(teacher);
						cell.setCellValue(coachToPlace);
						System.out.println(coachToPlace);
						return;
					}
				}
			}
			//For only time, not caring about day of week or field, but still placing them in correct size

			for(int x = 0; x < modifiedCorrespondingRows.size(); x++) {
				int tempRow1 = (int) modifiedCorrespondingRows.get(x);
				Row examine1 = availableFields.getRow(tempRow1); //Gets the row of the field and looks for an empty box
				if(preferredTime.equals("AM")) {
				for(int z = 2; z < 4; z++) {
					if (examine1.getCell(z) == null || examine1.getCell(z).toString().equals("")) {					
						Cell cell = examine1.getCell(z);
						//System.out.println("Got MODIFIED Cell");
						//System.out.println(teacher);
						cell.setCellValue(coachToPlace);
						System.out.println(coachToPlace);
						return;
					}
				}
				for(int z = 6; z < 8; z++) {
					if (examine1.getCell(z) == null || examine1.getCell(z).toString().equals("")) {					
						Cell cell = examine1.getCell(z);
						//System.out.println("Got MODIFIED Cell");
						//System.out.println(teacher);
						cell.setCellValue(coachToPlace);
						System.out.println(coachToPlace);
						return;
					}
				}
			}
	
			}
			if(preferredTime.equals("PM")) {
				for(int x = 0; x < modifiedCorrespondingRows.size(); x++) {
					int tempRow1 = (int) modifiedCorrespondingRows.get(x);
					Row examine1 = availableFields.getRow(tempRow1); //Gets the row of the field and looks for an empty box
					for(int z = 5; z < 7; z++) {
						if (examine1.getCell(z) == null || examine1.getCell(z).toString().equals("")) {					
							Cell cell = examine1.getCell(z);
							//System.out.println("Got MODIFIED P,Cell");
							//System.out.println(teacher);
							cell.setCellValue(coachToPlace);
							//System.out.println(coachToPlace);
							return;
						}
					}
					for(int z = 8; z < 10; z++) {
						if (examine1.getCell(z) == null || examine1.getCell(z).toString().equals("")) {					
							Cell cell = examine1.getCell(z);
							//System.out.println("Got MODIFIED PM Cell");
							//System.out.println(teacher);
							cell.setCellValue(coachToPlace);
							//System.out.println(coachToPlace);
							return;
						}
					}
				}
	
			}
		
		}
		System.out.println(coachToPlace + " can't be placed because of a scheduling error. Please adjust their preferences.");
}
}
