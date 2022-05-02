package execute;

import utils.ExcelUtils;

public class ExcelUtilsExecution {

    public static void main(String[] args) {
        //https://www.youtube.com/watch?v=CV3SOorFydE&list=PLhW3qG5bs-L8oRay6qeS70vJYZ3SBQnFa&index=20

        String projectPath = System.getProperty("user.dir");
        ExcelUtils excel = new ExcelUtils(projectPath + "/excel/MyExcelFile.xlsx", "Sheet1");
        excel.getRowCount();
        excel.getCellDataNumber(1,2);
        excel.getCellDataString(0,0);

    }
}
