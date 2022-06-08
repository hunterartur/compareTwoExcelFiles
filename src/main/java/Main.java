import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args) throws IOException {
        String pathToFilePrikrep = "D:\\files\\Prikrep.xlsx";
        String pathToFilePrikrepZ = "D:\\files\\PrikrepZ.xlsx";
        String testFile = "D:\\files\\PrikrepNoZ.xlsx";

//        Workbook prikrep = new XSSFWorkbook(new FileInputStream(pathToFilePrikrep));
//        Workbook prikrepZ = new XSSFWorkbook(new FileInputStream(pathToFilePrikrepZ));
        Workbook test = new XSSFWorkbook(new FileInputStream(testFile));

        Map<String, List<String>> dataFromExcel = getDataFromExcel(test);
//        Map<String, List<String>> dataFromExcel1 = dataFromExcel.entrySet().stream().sorted().collect(Collectors.toMap());
        dataFromExcel.forEach((s, strings) -> System.out.println(s + ": " + strings));
    }

    public static Map<String, List<String>> getDataFromExcel(Workbook workbook) {
        Map<String, List<String>> map = new HashMap<>();
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            List<String> value = new ArrayList<>();
            Iterator<Cell> iterator = row.iterator();
            iterator.next();
            String key = iterator.next().getStringCellValue();
            while (iterator.hasNext()) {
                value.add(iterator.next().getStringCellValue());
            }
            map.put(key, value);
        }
        return map;
    }
}
