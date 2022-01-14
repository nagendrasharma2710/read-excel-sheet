package ReadExcelFile.readexcelsheet;

import java.io.IOException;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class ReadExcelSheetApplication {
	
	public static void main(String []args) throws IOException {
		String excelPath = "C:\\Users\\praveen\\Downloads\\read-excel-sheet\\read-excel-sheet\\data\\EmployeeSheet.xlsx";
		String sheetName = "EmployeeSheet";
		ReadExcelSheetMethods excel = new ReadExcelSheetMethods(excelPath, sheetName);
		
		excel.getRowCount();
		System.out.println("------------------------");
		System.out.println("Details of First Row : ");
		System.out.println("------------------------");
		excel.getCellData(1, 0);
		excel.getCellData(1, 1);
		excel.getCellData(1, 2);
		excel.getCellData(1, 3);
		excel.getCellData(1, 4);
		System.out.println("------------------------");
		System.out.println("Details of Second Row : ");
		System.out.println("------------------------");
		excel.getCellData(2, 0);
		excel.getCellData(2, 1);
		excel.getCellData(2, 2);
		excel.getCellData(2, 3);
		excel.getCellData(2, 4);
		System.out.println("------------------------");
		System.out.println("Details of Third Row : ");
		System.out.println("------------------------");
		excel.getCellData(3, 0);
		excel.getCellData(3, 1);
		excel.getCellData(3, 2);
		excel.getCellData(3, 3);
		excel.getCellData(3, 4);
		System.out.println("------------------------");
		System.out.println("Details of Fourth Row : ");
		System.out.println("------------------------");
		excel.getCellData(4, 0);
		excel.getCellData(4, 1);
		excel.getCellData(4, 2);
		excel.getCellData(4, 3);
		excel.getCellData(4, 4);
		System.out.println("------------------------");
		System.out.println("Details of Fifth Row : ");
		System.out.println("------------------------");
		excel.getCellData(5, 0);
		excel.getCellData(5, 1);
		excel.getCellData(5, 2);
		excel.getCellData(5, 3);
		excel.getCellData(5, 4);
	}
}
