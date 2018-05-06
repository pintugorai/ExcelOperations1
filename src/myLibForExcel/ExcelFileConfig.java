package myLibForExcel;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileConfig {
	XSSFWorkbook workbook;
	XSSFSheet sheet;

	public ExcelFileConfig(String excelfilePath) {

		try {
			File myfile = new File(excelfilePath);
			FileInputStream fis = new FileInputStream(myfile);
			workbook = new XSSFWorkbook(fis);

		} catch (Exception e) {

			System.out.println(e.getMessage());
		}

	}

	public String getDataBySheetNumber(int sheetNumber, int rowNumber, int cellNumber)
	{
		sheet = workbook.getSheetAt(sheetNumber);
		String data = sheet.getRow(rowNumber).getCell(cellNumber).getStringCellValue();
		
		return data;
	}
	
	public String getDataBySheetName(String sheetName, int rowNumber, int cellNumber)
	{
		//sheet = workbook.getSheetAt(sheetName);
		
		int found = 0;
		for (int i=0;i<=workbook.getNumberOfSheets();i++)
		{
			if(workbook.getSheetAt(i).getSheetName().equalsIgnoreCase(sheetName))
				{
					sheet=workbook.getSheetAt(i);
					found=1;
					break;
				}
		}
		
		if(found == 0)
			return null;
		else {
			String data = sheet.getRow(rowNumber).getCell(cellNumber).getStringCellValue();
			return data;
			  
		}
	}
}
