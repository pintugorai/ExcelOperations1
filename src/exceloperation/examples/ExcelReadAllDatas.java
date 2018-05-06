package exceloperation.examples;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * 1. Read all the values in a sheets
 * 2. Read all the value in specific sheet
 */

public class ExcelReadAllDatas {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		// Add a File class and create object. Load the excel sheet
				File myfile= new File("D:\\Knowledge based\\MyJava\\Excel Operation in Java\\ExcelRead.xlsx");
				
				// Load the excel sheet as a form of bytes.
				FileInputStream fis = new FileInputStream(myfile);
				
				// Read as X-SSF-Workbook. This class coming form poi lib added as external jars 
				//It will load complete workbook(Complete excel file is called workbook).
				// XSSFWorkbook class for .xlsx file
				// HSSFWorkbook class for .xls file
				XSSFWorkbook workbook = new XSSFWorkbook(fis);
				
				// loaded the sheet 0/first sheet
				XSSFSheet sheet0 = workbook.getSheetAt(0);
				
				// Check how many row it have
				
				int noOfRows = sheet0.getLastRowNum();
				System.out.println("number of rows = " + (noOfRows + 1)); // row start with 0.
				
				int noOfSheet = workbook.getNumberOfSheets();
				
				//1. Print all values in sheet0
				for (int i=0;i<=noOfRows;i++)
				{
					for(int j=0; j<2;j++)
					System.out.println("Cell(" + i + "," + j + ")=" + sheet0.getRow(i).getCell(j).getStringCellValue());
				}
				
				//Get sheetname
				String sheetName = sheet0.getSheetName().toString();
				System.out.println("Sheet name: " + sheetName);
				
				//2. Read all the value of sheet name is sheet3
				int found=0;
				XSSFSheet targetSheet = null;
				for (int i=0;i<=workbook.getNumberOfSheets();i++)
				{
					if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase("sheet3"))
						{
							targetSheet=workbook.getSheetAt(i);
							found=1;
							break;
						}
				}
				
				if(found == 0)
					System.out.println("Sheet name not exists");
				else {
					
					System.out.println("Target Sheet Name = " + targetSheet.getSheetName());
				// Print all values in target sheet
				for (int i=0;i<=noOfRows;i++)
				{
					for(int j=0; j<2;j++)
					System.out.println("Cell(" + i + "," + j + ")=" + targetSheet.getRow(i).getCell(j).getStringCellValue());
				}
				}
				
				workbook.close();
				
	}

}
