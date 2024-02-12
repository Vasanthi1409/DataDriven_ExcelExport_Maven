package data_Driver_Excel_Export.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Export {

	public static void main(String[] args) throws IOException {
		
		File Excel = new File("C:\\\\Users\\\\inbaraj\\\\eclipse-workspace\\\\Maven_TCF\\\\Testcases_Login.xlsx");
		
		FileInputStream Excel_Input = new FileInputStream(Excel);
		
		Workbook WB = new XSSFWorkbook(Excel_Input);
		
		Sheet sheet_index = WB.getSheetAt(0);
		
		Row row_no = sheet_index.getRow(0);
		
		Cell cell_no = row_no.getCell(0);
		
		CellType cell_Type = cell_no.getCellType();

		
		if(cell_Type.equals(CellType.NUMERIC)) {
			double numericCellValue = cell_no.getNumericCellValue();
			System.out.println(numericCellValue);
		}else if(cell_Type.equals(CellType.STRING)) {
			String stringCellValue = cell_no.getStringCellValue();
			System.out.println(stringCellValue);
		}
	}

}
