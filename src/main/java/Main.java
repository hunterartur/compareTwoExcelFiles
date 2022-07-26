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
        String kt = "D:\\files\\KT.xlsx";
        String pathToFilePrikrepZ = "D:\\files\\p2c_7.xlsx";
        String result = "D:\\files\\result.xlsx";
        String resultSet = "D:\\files\\resultSet.xlsx";

        Workbook workbook1 = new XSSFWorkbook(new FileInputStream(kt));
//        Workbook workbook2 = new XSSFWorkbook(new FileInputStream(pathToFilePrikrepZ));

        Set<String> dataEx = getDataEx(workbook1.getSheetAt(0), 4);
        Set<String> dataEx1 = getDataEx(workbook1.getSheetAt(1), 2);
//        Map<String, String> maps2 = getDataFromExcel1(workbook2);

        Set<String> stringSet = symmetricDifference(dataEx, dataEx1);

//        System.out.println(maps2.size());
//        System.out.println(maps2.size() - maps.size());
//        System.out.println(33404 - maps2.size());
//        Map<String, List<String>> dataFromExcel = getDataFromExcel(workbook1);
//        Map<String, List<String>> dataToExcel = getDataFromExcel(workbook2);
//
//        Map<String, String> elems = findElems1(maps2, maps);
//        Map<String, String> elems2 = findElems1(maps, maps2);
//        elems.forEach((s, strings) -> System.out.println(s + ": " + strings));
//
//        Workbook wb = createWB1(elems);
//        wb.write(new FileOutputStream(result));
//        wb.close();

        Workbook wb1 = createSetToExcell(stringSet);
        wb1.write(new FileOutputStream(resultSet));
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


    public static Workbook createSetToExcell(Set<String> stringsSet) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        int i = 0;
        for (String str : stringsSet) {
            String[] s = str.split(" ");
            Row row = sheet.createRow(i++);
            for (int j = 0; j < s.length; j++) {
                row.createCell(j).setCellValue(s[j]);
            }
        }
        return workbook;
    }

    //Присылаем лист и количество значичих ячек начиная с 0
    public static Set<String> getDataEx(Sheet sheet, int n) {
        Set<String> strings = new HashSet<>();
        for (Row row : sheet) {
            StringBuilder stringBuilder = new StringBuilder();
            for (int i = 0; i < n; i++) {
                stringBuilder.append(row.getCell(i)).append(" ");
            }
            strings.add(stringBuilder.toString().trim().toLowerCase());
        }
        return strings;
    }

    public static Set<String> symmetricDifference(Set<String> set1, Set<String> set2) {
        //твой код здесь
        Set<String> sets1 = new HashSet<>(set1);
        Set<String> sets2 = new HashSet<>(set1);
        sets1.retainAll(set2);
        sets2.addAll(set2);
        sets2.removeAll(sets1);
        return sets2;
    }

//    public static <T> Set<T> symmetricDifference(Set<? extends T> set1, Set<? extends T> set2) {
//        //твой код здесь
//        Set<T> sets1 = new HashSet<>(set1);
//        Set<T> sets2 = new HashSet<>(set1);
//        sets1.retainAll(set2);
//        sets2.addAll(set2);
//        sets2.removeAll(sets1);
//        return sets2;
//    }
}
