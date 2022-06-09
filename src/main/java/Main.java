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
        String pathToFilePrikrep = "D:\\files\\prikrep.xlsx";
        String pathToFilePrikrepZ = "D:\\files\\p2c_7.xlsx";
        String result = "D:\\files\\result.xlsx";
        String result2 = "D:\\files\\result2.xlsx";

        Workbook workbook1 = new XSSFWorkbook(new FileInputStream(pathToFilePrikrep));
        Workbook workbook2 = new XSSFWorkbook(new FileInputStream(pathToFilePrikrepZ));

        Map<String, String> maps = getDataFromExcel1(workbook1);
        Map<String, String> maps2 = getDataFromExcel1(workbook2);

//        Map<String, List<String>> dataFromExcel = getDataFromExcel(workbook1);
//        Map<String, List<String>> dataToExcel = getDataFromExcel(workbook2);
//
        Map<String, String> elems = findElems1(maps2, maps);
        Map<String, String> elems2 = findElems1(maps, maps2);
//        elems.forEach((s, strings) -> System.out.println(s + ": " + strings));
//
        Workbook wb = createWB1(elems);
        wb.write(new FileOutputStream(result));
        wb.close();

        Workbook wb1 = createWB1(elems2);
        wb1.write(new FileOutputStream(result2));
        wb1.close();
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

    public static Workbook createWB1(Map<String, String> map) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        int i = 0;
        for (Map.Entry<String, String> elem : map.entrySet()) {
            Row row = sheet.createRow(i++);
            row.createCell(0).setCellValue(elem.getKey());
            row.createCell(1).setCellValue(elem.getValue());
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

    public static Map<String, String> findElems1(Map<String, String> from, Map<String, String> to) {
        Map<String, String> newMap = new LinkedHashMap<>();
        for (Map.Entry<String, String> elem : from.entrySet()) {
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

    public static Map<String, String> getDataFromExcel1(Workbook workbook) {
        Map<String, String> map = new HashMap<>();
        Sheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            StringBuilder key = new StringBuilder();
            Iterator<Cell> iterator = row.iterator();
            for (int i = 0; i < 4; i++) {
                if (iterator.hasNext()) {
                    key.append(iterator.next().getStringCellValue().trim().toLowerCase()).append(" ");
                }
            }
            if (row.getCell(4) == null) {
                map.put(key.toString(), "Null");
            } else {
                String value = row.getCell(4).getStringCellValue();
                map.put(key.toString(), value);
            }

        }
        return map;
    }
}
