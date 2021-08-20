package Utility;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtility {
	//Create Variables
	static String projectPath;
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	//Create Constructor and Parametrize to get excelPath and sheetName
	//Contructor is a special method without any return type and name same as class name
	//Constructor is called every time a class object is created using the new() keyword
	public ExcelUtility(String excelPath, String sheetName) {
		try {
			String projectPath = System.getProperty("user.dir");
			XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
			XSSFSheet sheet = workbook.getSheet("sheetName");
		}catch(Exception e) {
			e.printStackTrace();
		}
	}
	public static void main(String[] args) {
		getRowCount();
		getCellDataString(0,0);
		getCellDataNumber(1,1);
	}
	//Create Row Count Function
	public static void getRowCount() {
		//surround try/catch for Exception Handling
		try {
			//For location of framework
			String projectPath = System.getProperty("user.dir");
			//Create references for Workbook and location of Workbook
			XSSFWorkbook workbook = new XSSFWorkbook(projectPath+"\\ExcelFile\\Data.xlsx");  
			//Create references for Worksheet
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			//Get the row count
			int rowCount = sheet.getPhysicalNumberOfRows();
			System.out.println("No of rows :"+rowCount);
		}catch(Exception exp) {
			System.out.println(exp.getMessage());//if any exception occurs
			System.out.println(exp.getCause());//if any exception occurs
			exp.printStackTrace();//if any exception occurs
		}	
	}
	//Create CellData Function
	public static void getCellDataString(int rowNum, int colNum) { //use Camel's case
		try {
			String projectPath = System.getProperty("user.dir");
			workbook = new XSSFWorkbook(projectPath+"\\ExcelFile\\Data.xlsx");
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			String cellData = sheet.getRow(rowNum).getCell(colNum).getStringCellValue();
			System.out.println(cellData);
		}catch(Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();
		}
	}
	//Create CellDataNumber Function
	public static void getCellDataNumber(int rowNum, int colNum) {	//use Camel's case
		try {
			String projectPath = System.getProperty("user.dir");
			XSSFWorkbook workbook = new XSSFWorkbook(projectPath+"\\ExcelFile\\Data.xlsx");
			XSSFSheet sheet = workbook.getSheet("Sheet1");
			double cellData = sheet.getRow(rowNum).getCell(colNum).getNumericCellValue();
			System.out.println(cellData);
		}catch(Exception exp) {
			System.out.println(exp.getMessage());
			System.out.println(exp.getCause());
			exp.printStackTrace();		
		}
	}
}

