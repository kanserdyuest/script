import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelDataToSql {

    private static final String sheetName = "user";

    private static final String INSERT = "INSERT INTO ";

    // INSERT INTO aspdb.`user`
    // (id, name, pwd)
    // VALUES(0, '', '');

    public static void main(String[] args) throws Exception {
        // File file = new File("c:\\Users\\Jar\\Desktop\\test.xlsx");
        // File file = new File("c:\\Users\\Jar\\Desktop\\test.xlsx");
        FileInputStream file = new FileInputStream("c:\\Users\\Jar\\Desktop\\test.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(sheetName);
        String sqlString = INSERT + "'" + sheetName + "' (";
        XSSFRow firstRow = sheet.getRow(0);

        for (int i = 3; i < firstRow.getLastCellNum(); i++) {
            sqlString = sqlString + firstRow.getCell(i).getRichStringCellValue() + ",";
        }
        sqlString = sqlString.substring(0, sqlString.length() - 1) + ") VALUES( ";

        for (int r = 1; r < sheet.getPhysicalNumberOfRows(); r++) {
            XSSFRow row = sheet.getRow(r);
            String rowSqlString = sqlString;
            for (int c = 3; c < row.getLastCellNum(); c++) {
                XSSFCell cell = row.getCell(c);
                String stringCellValue;
                if (cell == null) {
                    stringCellValue = " NULL, ";
                } else {
                    cell.setCellType(CellType.STRING);
                    stringCellValue = "'" + cell.getStringCellValue() + "',";
                }
                rowSqlString = rowSqlString + stringCellValue;
            }
            rowSqlString = rowSqlString.substring(0, rowSqlString.length() - 1) + ");";
            row.createCell(row.getLastCellNum()).setCellValue(rowSqlString);
        }
        FileOutputStream outputStream = new FileOutputStream("c:\\Users\\Jar\\Desktop\\test.xlsx");

        workbook.write(outputStream);
        outputStream.close();
        file.close();
    }
}
