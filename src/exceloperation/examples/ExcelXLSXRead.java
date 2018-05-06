package exceloperation.examples;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelXLSXRead {

	public static void main(String[] args) throws Exception {
		
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
		
		// Get cell(0,0). Specify which row and which collum
		String cell00= sheet0.getRow(0).getCell(0).getStringCellValue();
		
		System.out.println("Cell00 = " + cell00);
		
		// Close the workbook. Otherwise some case it will leads to memory leak
		workbook.close();
		
		// Print all the records of sheet0(first sheet)
		
		
		
	}
}
