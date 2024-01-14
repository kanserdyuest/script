import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.hpsf.Array;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
    public static void main(String[] args) throws Exception {
        File file = new File("C:/Users/Jar/Desktop/Temp.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet("Test");

        List<String> list = new ArrayList<String>(Arrays.asList("22", "23", "2", "6"));
        XSSFSheet newSheet = workbook.createSheet("newSheet");
        for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            XSSFRow row = sheet.getRow(i);
            if (row.getCell(4) != null) {
                row.getCell(4).setCellType(CellType.STRING);
                for (String s : list) {
                    if (s.equals(row.getCell(4).getRichStringCellValue().getString())) {
                        System.out.println("发现一条");
                        newSheet.createRow(0).copyRowFrom(row, new CellCopyPolicy());
                    }
                }
            }

        }

        FileOutputStream os = new FileOutputStream(file);
        workbook.write(os);
        os.close();
    }
}
