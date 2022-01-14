package ReadExcelFile.readexcelsheet;

import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelSheetMethods {
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;

	public ReadExcelSheetMethods(String excelPath, String sheetName) {
		try {
			workbook = new XSSFWorkbook(excelPath);
			sheet = workbook.getSheet(sheetName);
		} catch(Exception ex) {
			System.out.println(ex.getCause());
			System.out.println(ex.getLocalizedMessage());
			ex.printStackTrace();
		}
	}
	public static void getCellData(int rowNum, int colNum) throws IOException {
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
		System.out.println(value);
	}
	public static void getRowCount() {
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("Total No of Rows in Table : "+rowCount);
	} 
}


