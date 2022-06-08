import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args) throws IOException {
        String pathToFilePrikrep = "D:\\files\\Prikrep.xlsx";
        String pathToFilePrikrepZ = "D:\\files\\PrikrepZ.xlsx";
        String result = "D:\\files\\result.xlsx";

//        Workbook prikrep = new XSSFWorkbook(new FileInputStream(pathToFilePrikrep));
//        Workbook prikrepZ = new XSSFWorkbook(new FileInputStream(pathToFilePrikrepZ));

        Workbook workbook1 = new XSSFWorkbook(new FileInputStream(pathToFilePrikrep));
        Workbook workbook2 = new XSSFWorkbook(new FileInputStream(pathToFilePrikrepZ));

        Map<String, List<String>> dataFromExcel = sort(getDataFromExcel(workbook1));
        Map<String, List<String>> dataToExcel = sort(getDataFromExcel(workbook2));

        Map<String, List<String>> elems = findElems(dataFromExcel, dataToExcel);
        elems.forEach((s, strings) -> System.out.println(s + ": " + strings));

        Workbook wb = createWB(elems);
        wb.write(new FileOutputStream(result));
        wb.close();
    }

    public static Workbook createWB(Map<String, List<String>> map) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        int i = 0;
        for (Map.Entry<String, List<String>> elem : map.entrySet()) {
            Row row = sheet.createRow(i++);
            row.createCell(0).setCellValue(elem.getKey());
            int j = 1;
            for (String s : elem.getValue()) {
                row.createCell(j++).setCellValue(s);
            }
        }
        return workbook;
    }

    public static Map<String, List<String>> findElems(Map<String, List<String>> from, Map<String, List<String>> to) {
        Map<String, List<String>> newMap = new LinkedHashMap<>();
        for (Map.Entry<String, List<String>> elem : from.entrySet()) {
            if (!to.containsKey(elem.getKey())) {
                newMap.put(elem.getKey(), elem.getValue());
            }
        }
        return newMap;
    }

    public static Map<String, List<String>> getDataFromExcel(Workbook workbook) {
        Map<String, List<String>> map = new HashMap<>();
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            List<String> value = new ArrayList<>();
            Iterator<Cell> iterator = row.iterator();
//            iterator.next();
            String key = iterator.next().getStringCellValue();
            while (iterator.hasNext()) {
                value.add(iterator.next().getStringCellValue());
            }
            map.put(key, value);
        }
        return map;
    }

    public static Map<String, List<String>> sort(Map<String, List<String>> unsortedMap) {
        return unsortedMap
                .entrySet()
                .stream()
                .sorted(Comparator.comparing(o -> Long.valueOf(o.getKey())))
                .collect(Collectors.toMap(
                        Map.Entry::getKey,
                        Map.Entry::getValue,
                        (a, b) -> {
                            throw new AssertionError();
                        },
                        LinkedHashMap::new
                ));
    }
}
