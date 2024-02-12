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

public class Excel_Export_Row {

	public static void main(String[] args) throws IOException {

		File Excel = new File("C:\\Users\\inbaraj\\eclipse-workspace\\Maven_TCF\\Testcases_Login.xlsx");

		FileInputStream Excel_Input = new FileInputStream(Excel);

		Workbook WB = new XSSFWorkbook(Excel_Input);

		Sheet sheet_name = WB.getSheet("Sheet1");

		Row row = sheet_name.getRow(0);

		short lastCell_count = row.getLastCellNum();

		for (int i = 0; i <= lastCell_count; i++) {

			Cell cell = row.getCell(i);
			CellType cell_Type;
			if (cell != null) {
				cell_Type = cell.getCellType();
				if (cell_Type.equals(CellType.NUMERIC)) {
					double numericCellValue = cell.getNumericCellValue();
					System.out.println(numericCellValue);
				} else if (cell_Type.equals(CellType.STRING)) {
					String stringCellValue = cell.getStringCellValue();
					System.out.println(stringCellValue);
				}
			}

		}
	}
}
