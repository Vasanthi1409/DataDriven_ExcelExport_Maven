package data_Driver_Excel_Export.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_Export_RowColumn {

	public static void main(String[] args) throws IOException {

		File Excel = new File("C:\\Users\\inbaraj\\eclipse-workspace\\Maven_TCF\\Testcases_Login.xlsx");

		FileInputStream Excel_Input = new FileInputStream(Excel);

		Workbook WB = new XSSFWorkbook(Excel_Input);

		Sheet sheet_name = WB.getSheetAt(0);

		int lastRow_Count = sheet_name.getLastRowNum();

		for (int i = 0; i <= lastRow_Count; i++) {
			Row row = sheet_name.getRow(i);
			short lastCell_Count = row.getLastCellNum();
			
			for (int j = 0; j <= lastCell_Count; j++) {
				Cell cell = row.getCell(j);
				CellType cell_Type;
				
				if(cell != null) {
					cell_Type = cell.getCellType();
					if (cell_Type.equals(CellType.NUMERIC)) {
						int numericCellValue = (int) cell.getNumericCellValue();
						System.out.println(numericCellValue);
					} else if (cell_Type.equals(CellType.STRING)) {
						String stringCellValue = cell.getStringCellValue();
						System.out.println(stringCellValue);
					}
				}
			}
		}
	}
}
