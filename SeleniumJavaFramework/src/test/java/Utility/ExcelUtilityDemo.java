package Utility;

public class ExcelUtilityDemo {

	private static String sheetName;
	//Create new class for excel functions to validate and use this anywhere in framework
	public static void main(String[] args) {
		String projectPath = System.getProperty("user.dir");
		//Create Object
		ExcelUtility excel = new ExcelUtility(projectPath+"\\ExcelFile\\Data.xlsx", sheetName);
		excel.getRowCount();
		excel.getCellDataString(0, 0);
		excel.getCellDataNumber(1, 1);
	}
}
