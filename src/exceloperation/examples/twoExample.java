package exceloperation.examples;

import myLibForExcel.ExcelFileConfig;

public class twoExample {

	public static void main(String[] args) {
		ExcelFileConfig epc = new ExcelFileConfig("D:\\\\Knowledge based\\\\MyJava\\\\Excel Operation in Java\\\\ExcelRead.xlsx");
		
		System.out.println(epc.getDataBySheetNumber(0, 0, 0));
		
		System.out.println(epc.getDataBySheetName("Sheet3", 1, 1));
	}
}
