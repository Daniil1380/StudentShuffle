import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.*;

public class Main {

    public static void main(String[] args) throws IOException {
        XSSFWorkbook book = new XSSFWorkbook();
        FileOutputStream fileOut = new FileOutputStream("students.xlsx");
        XSSFSheet sheet = book.createSheet("Список студентов");
        File file = new File("students.txt");
        List<String> list = FileUtils.readLines(file, StandardCharsets.UTF_8);
        Collections.shuffle(list);
        List<String> reverseList = new ArrayList<>(list);
        Collections.reverse(reverseList);
        showQueue(list, reverseList, sheet);
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        book.write(fileOut);
        fileOut.close();
    }



    public static void showQueue(List<String> list, List<String> reverseList, XSSFSheet sheet){
        for (int i = 0; i < list.size(); i++) {
            XSSFRow row = sheet.createRow((short)i);
            XSSFCell cell0 = row.createCell(0);
            XSSFCell cell1 = row.createCell(1);
            cell0.setCellType(CellType.STRING);
            cell0.setCellValue(list.get(i));
            cell1.setCellType(CellType.STRING);
            cell1.setCellValue(reverseList.get(i));
        }
    }
}
