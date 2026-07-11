package com.example.project.service.finance.cwgjx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.*;

/**
 * 财务工具箱 - 数据比较
 * 工人预支申请表 vs 预支明细表 对比（cw_bijiao_xzy.py）
 */
@Service
public class CwgjxBijiaoService {

    public CompareResult compare(MultipartFile file1, String sheet1Name,
                          MultipartFile file2, String sheet2Name,
                          String startDate, String endDate,
                          String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File temp1 = new File(dir, UUID.randomUUID() + "_" + file1.getOriginalFilename());
        File temp2 = new File(dir, UUID.randomUUID() + "_" + file2.getOriginalFilename());
        file1.transferTo(temp1);
        file2.transferTo(temp2);

        try {
            // 读取表1（工人预支申请表）
            Table1AggregateResult t1Result = loadTable1(temp1, sheet1Name, startDate, endDate);
            // 读取表2（预支明细表）
            List<Table2Row> table2 = loadTable2(temp2, sheet2Name, startDate, endDate, t1Result.multiIdNames);
            // 合并对比
            List<ResultRow> results = mergeCompare(t1Result.rows, table2, t1Result.multiIdNames);
            // 保存
            String outputPath = uploadDir + File.separator + "预支对比结果.xlsx";
            saveResults(results, outputPath);

            // 计算预览统计信息
            long trueCount = results.stream().filter(r -> "TRUE".equals(r.result)).count();
            long falseCount = results.size() - trueCount;
            double sumLeft = results.stream().mapToDouble(r -> r.amountLeft).sum();
            double sumDiff = results.stream().mapToDouble(r -> r.diff).sum();
            double sumRight = results.stream().mapToDouble(r -> r.amountRight).sum();

            ComparePreview preview = new ComparePreview();
            preview.totalCount = results.size();
            preview.trueCount = trueCount;
            preview.falseCount = falseCount;
            preview.sumLeft = sumLeft;
            preview.sumDiff = sumDiff;
            preview.sumRight = sumRight;

            CompareResult result = new CompareResult();
            result.fileName = new File(outputPath).getName();
            result.preview = preview;
            // 填充预览明细行（最多显示500行，避免响应过大）
            int maxRows = Math.min(results.size(), 500);
            for (int i = 0; i < maxRows; i++) {
                ResultRow r = results.get(i);
                Map<String, Object> row = new LinkedHashMap<>();
                row.put("name", r.name);
                row.put("amountLeft", r.amountLeft);
                row.put("idCard", r.idCard);
                row.put("result", r.result);
                row.put("diff", r.diff);
                row.put("nameRight", r.nameRight);
                row.put("amountRight", r.amountRight);
                row.put("remark", r.remark);
                row.put("unit", r.unit);
                result.previewRows.add(row);
            }
            return result;
        } finally {
            temp1.delete();
            temp2.delete();
        }
    }

    public static class CompareResult {
        public String fileName;
        public ComparePreview preview;
        public List<Map<String, Object>> previewRows = new ArrayList<>();
    }

    public static class ComparePreview {
        public int totalCount;
        public long trueCount;
        public long falseCount;
        public double sumLeft;
        public double sumDiff;
        public double sumRight;
    }

    private Table1AggregateResult loadTable1(File file, String sheetName, String startDate, String endDate) throws Exception {
        List<Table1Row> rows = new ArrayList<>();
        Date s = parseDate(startDate), e = parseDateEndOfDay(endDate);
        try (Workbook wb = WorkbookFactory.create(new FileInputStream(file))) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) throw new Exception("工作表不存在: " + sheetName);

            // 查找表头（前10行）
            Map<String, Integer> cols = findHeaders(sheet, Arrays.asList("姓名", "预支金额", "身份证号", "提交时间", "运营中心"));
            if (cols == null) cols = new HashMap<>();
            // 固定列兜底（Python 中的固定布局）
            if (!cols.containsKey("姓名")) cols.put("姓名", 8);
            if (!cols.containsKey("预支金额")) cols.put("预支金额", 5);
            if (!cols.containsKey("身份证号")) cols.put("身份证号", 18);
            if (!cols.containsKey("提交时间")) cols.put("提交时间", 13);
            if (!cols.containsKey("运营中心")) cols.put("运营中心", 3);

            int nameCol = cols.get("姓名"), amountCol = cols.get("预支金额");
            int idCol = cols.get("身份证号"), dateCol = cols.get("提交时间");
            Integer unitCol = cols.get("运营中心");

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = getStr(row, nameCol);
                Double amount = parseAmount(getStr(row, amountCol));
                String id = getStr(row, idCol);
                Date d = parseDateCell(row.getCell(dateCol));
                String unit = unitCol != null ? getStr(row, unitCol) : "";
                if (name == null || name.trim().isEmpty() || name.trim().equals("nan") || amount == null) continue;
                if (d == null || d.before(s) || d.after(e)) continue;
                rows.add(new Table1Row(name.trim(), amount, id.trim(), unit.trim()));
            }
        }
        return aggregateTable1(rows);
    }

    private Table1AggregateResult aggregateTable1(List<Table1Row> raw) {
        // 按姓名聚合，处理多身份证号场景
        Map<String, List<Table1Row>> byName = new LinkedHashMap<>();
        for (Table1Row r : raw) byName.computeIfAbsent(r.name, k -> new ArrayList<>()).add(r);

        Set<String> multiIdNames = new HashSet<>();
        List<Table1Row> result = new ArrayList<>();

        for (Map.Entry<String, List<Table1Row>> e : byName.entrySet()) {
            String name = e.getKey();
            List<Table1Row> group = e.getValue();

            Set<String> globalIds = new HashSet<>();
            for (Table1Row r : group) if (!r.idCard.isEmpty()) globalIds.add(r.idCard);

            if (globalIds.size() <= 1) {
                // 场景A/B：全局只有一个唯一非空身份证号，或全部空白
                double sum = group.stream().mapToDouble(r -> r.amount).sum();
                String id = globalIds.isEmpty() ? "" : globalIds.iterator().next();
                String unit = "";
                if (!id.isEmpty()) {
                    for (Table1Row r : group) {
                        if (r.idCard.equals(id) && !r.unit.isEmpty()) {
                            unit = r.unit;
                            break;
                        }
                    }
                }
                result.add(new Table1Row(name, sum, id, unit));
            } else {
                // 场景C：全局存在多个不同非空身份证号 -> 按运营中心子判断
                multiIdNames.add(name);
                Map<String, List<Table1Row>> byUnit = new LinkedHashMap<>();
                for (Table1Row r : group) byUnit.computeIfAbsent(r.unit, k -> new ArrayList<>()).add(r);
                for (Map.Entry<String, List<Table1Row>> ue : byUnit.entrySet()) {
                    List<Table1Row> unitGroup = ue.getValue();
                    Set<String> unitIds = new HashSet<>();
                    for (Table1Row r : unitGroup) if (!r.idCard.isEmpty()) unitIds.add(r.idCard);

                    if (unitIds.size() <= 1) {
                        // a/b：该运营中心下只有一个唯一非空身份证，或全部空白
                        double sum = unitGroup.stream().mapToDouble(r -> r.amount).sum();
                        String id = unitIds.isEmpty() ? "" : unitIds.iterator().next();
                        result.add(new Table1Row(name, sum, id, ue.getKey()));
                    } else {
                        // c：该运营中心下仍有多个不同非空身份证 -> 按身份证号拆分
                        Map<String, List<Table1Row>> byId = new LinkedHashMap<>();
                        for (Table1Row r : unitGroup) {
                            String key = r.idCard.isEmpty() ? "" : r.idCard;
                            byId.computeIfAbsent(key, k -> new ArrayList<>()).add(r);
                        }
                        for (Map.Entry<String, List<Table1Row>> ie : byId.entrySet()) {
                            double sum = ie.getValue().stream().mapToDouble(r -> r.amount).sum();
                            result.add(new Table1Row(name, sum, ie.getKey(), ue.getKey()));
                        }
                    }
                }
            }
        }
        return new Table1AggregateResult(result, multiIdNames);
    }

    private List<Table2Row> loadTable2(File file, String sheetName, String startDate, String endDate, Set<String> multiIdNames) throws Exception {
        List<Table2Row> rows = new ArrayList<>();
        Date s = parseDate(startDate), e = parseDateEndOfDay(endDate);
        try (Workbook wb = WorkbookFactory.create(new FileInputStream(file))) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) throw new Exception("工作表不存在: " + sheetName);

            // 固定第2行为表头（Python逻辑）
            Map<String, Integer> cols = findHeaders(sheet, Arrays.asList("预支日期", "姓名", "预支数额", "运营中心"));
            if (cols == null) cols = new HashMap<>();
            if (!cols.containsKey("预支日期")) cols.put("预支日期", 1);
            if (!cols.containsKey("姓名")) cols.put("姓名", 4);
            if (!cols.containsKey("预支数额")) cols.put("预支数额", 5);
            if (!cols.containsKey("运营中心")) cols.put("运营中心", 2);

            int dateCol = cols.get("预支日期"), nameCol = cols.get("姓名");
            int amountCol = cols.get("预支数额");
            Integer unitCol = cols.get("运营中心");

            for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = getStr(row, nameCol);
                Double amount = parseAmount(getStr(row, amountCol));
                Date d = parseDateCell(row.getCell(dateCol));
                String unit = unitCol != null ? getStr(row, unitCol) : "";
                if (name == null || name.trim().isEmpty() || name.trim().equals("nan") || amount == null) continue;
                if (d == null || d.before(s) || d.after(e)) continue;
                rows.add(new Table2Row(name.trim(), amount, unit.trim()));
            }
        }
        return aggregateTable2(rows, multiIdNames);
    }

    private List<Table2Row> aggregateTable2(List<Table2Row> raw, Set<String> multiIdNames) {
        if (multiIdNames == null || multiIdNames.isEmpty()) {
            Map<String, Double> byName = new LinkedHashMap<>();
            for (Table2Row r : raw) byName.merge(r.name, r.amount, Double::sum);
            List<Table2Row> result = new ArrayList<>();
            for (Map.Entry<String, Double> e : byName.entrySet()) result.add(new Table2Row(e.getKey(), e.getValue(), ""));
            return result;
        }

        List<Table2Row> multiRows = new ArrayList<>();
        List<Table2Row> normalRows = new ArrayList<>();
        for (Table2Row r : raw) {
            if (multiIdNames.contains(r.name)) multiRows.add(r);
            else normalRows.add(r);
        }

        List<Table2Row> result = new ArrayList<>();
        // 普通姓名按姓名汇总，运营中心留空
        Map<String, Double> normalByName = new LinkedHashMap<>();
        for (Table2Row r : normalRows) normalByName.merge(r.name, r.amount, Double::sum);
        for (Map.Entry<String, Double> e : normalByName.entrySet()) result.add(new Table2Row(e.getKey(), e.getValue(), ""));

        // multi_id 姓名按 [姓名, 运营中心] 汇总
        Map<String, Double> multiByKey = new LinkedHashMap<>();
        for (Table2Row r : multiRows) {
            String key = r.name + "\0" + r.unit;
            multiByKey.merge(key, r.amount, Double::sum);
        }
        for (Map.Entry<String, Double> e : multiByKey.entrySet()) {
            String[] parts = e.getKey().split("\0", -1);
            result.add(new Table2Row(parts[0], e.getValue(), parts[1]));
        }
        return result;
    }

    private List<ResultRow> mergeCompare(List<Table1Row> t1, List<Table2Row> t2, Set<String> multiIdNames) {
        List<ResultRow> results = new ArrayList<>();
        Map<String, Double> rightMapName = new HashMap<>();
        Map<String, Double> rightMapNameUnit = new HashMap<>();

        for (Table2Row r : t2) {
            if (multiIdNames != null && multiIdNames.contains(r.name)) {
                String key = r.name + "\0" + r.unit;
                rightMapNameUnit.merge(key, r.amount, Double::sum);
            } else {
                rightMapName.merge(r.name, r.amount, Double::sum);
            }
        }

        Set<String> leftNames = new HashSet<>();

        for (Table1Row r : t1) {
            leftNames.add(r.name);
            if (multiIdNames != null && multiIdNames.contains(r.name)) {
                String key = r.name + "\0" + r.unit;
                Double ra = rightMapNameUnit.get(key);
                if (ra != null) {
                    boolean match = Math.abs(r.amount - ra) < 0.01;
                    double diff = r.amount - ra;
                    results.add(new ResultRow(r.name, r.amount, r.idCard, match ? "TRUE" : "FALSE",
                            diff, r.name, ra, remarkByDiff(diff), r.unit));
                } else {
                    results.add(new ResultRow(r.name, r.amount, r.idCard, "FALSE",
                            r.amount, "", 0.0, "申请表多报了", r.unit));
                }
            } else {
                Double ra = rightMapName.get(r.name);
                if (ra != null) {
                    boolean match = Math.abs(r.amount - ra) < 0.01;
                    double diff = r.amount - ra;
                    results.add(new ResultRow(r.name, r.amount, r.idCard, match ? "TRUE" : "FALSE",
                            diff, r.name, ra, remarkByDiff(diff), ""));
                } else {
                    results.add(new ResultRow(r.name, r.amount, r.idCard, "FALSE",
                            r.amount, "", 0.0, "申请表多报了", ""));
                }
            }
        }

        for (Table2Row r : t2) {
            if (!leftNames.contains(r.name)) {
                results.add(new ResultRow(r.name, 0.0, "", "FALSE",
                        -r.amount, r.name, r.amount, "明细表有但没申请", r.unit));
            }
        }

        results.sort(Comparator.comparing((ResultRow r) -> r.name).thenComparing(r -> -r.amountLeft));
        return results;
    }

    private String remarkByDiff(double diff) {
        if (diff > 0) return "申请表多报了";
        if (diff < 0) return "明细表有但没申请";
        return "";
    }

    private void saveResults(List<ResultRow> results, String path) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("对比结果");
            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font hf = wb.createFont(); hf.setColor(IndexedColors.WHITE.getIndex()); hf.setBold(true);
            headerStyle.setFont(hf);

            CellStyle trueStyle = wb.createCellStyle();
            XSSFColor green = new XSSFColor(new byte[]{(byte)0xE2, (byte)0xEF, (byte)0xDA}, null);
            trueStyle.setFillForegroundColor(green);
            trueStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle falseStyle = wb.createCellStyle();
            XSSFColor red = new XSSFColor(new byte[]{(byte)0xFC, (byte)0xE4, (byte)0xEC}, null);
            falseStyle.setFillForegroundColor(red);
            falseStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            String[] headers = {"姓名", "预支金额", "身份证号", "比较结果", "差额", "姓名(右)", "预支数额(右)", "备注", "运营中心"};
            Row h = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell c = h.createCell(i); c.setCellValue(headers[i]); c.setCellStyle(headerStyle);
            }

            for (int i = 0; i < results.size(); i++) {
                ResultRow r = results.get(i);
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(r.name);
                row.createCell(1).setCellValue(r.amountLeft);
                row.createCell(2).setCellValue(r.idCard);
                row.createCell(3).setCellValue(r.result);
                row.createCell(4).setCellValue(r.diff);
                row.createCell(5).setCellValue(r.nameRight);
                row.createCell(6).setCellValue(r.amountRight);
                row.createCell(7).setCellValue(r.remark);
                row.createCell(8).setCellValue(r.unit);
                CellStyle style = "TRUE".equals(r.result) ? trueStyle : falseStyle;
                for (int c = 0; c < headers.length; c++) row.getCell(c).setCellStyle(style);
            }

            // 统计行
            int totalRow = results.size() + 2;
            double sumLeft = results.stream().mapToDouble(r -> r.amountLeft).sum();
            double sumDiff = results.stream().mapToDouble(r -> r.diff).sum();
            double sumRight = results.stream().mapToDouble(r -> r.amountRight).sum();
            Row tr = sheet.createRow(totalRow);
            tr.createCell(0).setCellValue("总计");
            tr.createCell(1).setCellValue(sumLeft);
            tr.createCell(4).setCellValue(sumDiff);
            tr.createCell(6).setCellValue(sumRight);

            long trueCount = results.stream().filter(r -> "TRUE".equals(r.result)).count();
            long falseCount = results.size() - trueCount;
            Row s1 = sheet.createRow(totalRow + 2);
            s1.createCell(0).setCellValue("相符(TRUE)"); s1.createCell(1).setCellValue(trueCount);
            Row s2 = sheet.createRow(totalRow + 3);
            s2.createCell(0).setCellValue("不符(FALSE)"); s2.createCell(1).setCellValue(falseCount);
            Row s3 = sheet.createRow(totalRow + 4);
            s3.createCell(0).setCellValue("总计"); s3.createCell(1).setCellValue(results.size());

            try (FileOutputStream fos = new FileOutputStream(path)) { wb.write(fos); }
        }
    }

    private Map<String, Integer> findHeaders(Sheet sheet, List<String> required) {
        for (int r = 0; r < Math.min(10, sheet.getLastRowNum() + 1); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Map<String, Integer> found = new HashMap<>();
            for (Cell cell : row) {
                if (cell == null) continue;
                String v = getStr(cell);
                for (String req : required) {
                    if (v.contains(req)) { found.put(req, cell.getColumnIndex()); break; }
                }
            }
            if (found.size() == required.size()) return found;
        }
        return null;
    }

    private Date parseDate(String s) {
        try {
            LocalDate d = LocalDate.parse(s, DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            return Date.from(d.atStartOfDay(ZoneId.systemDefault()).toInstant());
        } catch (DateTimeParseException e) {
            return new Date();
        }
    }

    private Date parseDateEndOfDay(String s) {
        try {
            LocalDate d = LocalDate.parse(s, DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            return Date.from(d.atTime(23, 59, 59, 999999999).atZone(ZoneId.systemDefault()).toInstant());
        } catch (DateTimeParseException e) {
            return new Date();
        }
    }

    private Date parseDateCell(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getDateCellValue();
            }
            String s = getStr(cell).trim();
            if (s.isEmpty() || s.equalsIgnoreCase("nan")) return null;
            for (String fmt : new String[]{"yyyy-MM-dd", "yyyy/MM/dd", "yyyy年MM月dd日"}) {
                try {
                    LocalDate d = LocalDate.parse(s, DateTimeFormatter.ofPattern(fmt));
                    return Date.from(d.atStartOfDay(ZoneId.systemDefault()).toInstant());
                } catch (DateTimeParseException ignored) {}
            }
            try {
                double d = Double.parseDouble(s);
                if (d > 1 && d < 100000) return DateUtil.getJavaDate(d);
            } catch (NumberFormatException ignored) {}
        } catch (Exception ignored) {}
        return null;
    }

    private Double parseAmount(String s) {
        if (s == null) return null;
        s = s.trim();
        if (s.isEmpty() || s.equalsIgnoreCase("nan") || s.equalsIgnoreCase("none")) return 0.0;
        String[] zeros = {"未预支", "无", "零", "没有"};
        for (String z : zeros) if (s.equals(z)) return 0.0;
        s = s.replaceAll("[元块毛分钱整]$", "").replace(",", "").replace("，", "");
        try { return Double.parseDouble(s); } catch (NumberFormatException e) { return null; }
    }

    private String getStr(Row row, int col) {
        Cell cell = row.getCell(col);
        return cell == null ? "" : getStr(cell);
    }

    private String getStr(Cell cell) {
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }

    private static class Table1AggregateResult {
        List<Table1Row> rows;
        Set<String> multiIdNames;
        Table1AggregateResult(List<Table1Row> rows, Set<String> multiIdNames) {
            this.rows = rows;
            this.multiIdNames = multiIdNames;
        }
    }

    private static class Table1Row {
        String name; double amount; String idCard; String unit;
        Table1Row(String n, double a, String i, String u) { name = n; amount = a; idCard = i; unit = u; }
    }
    private static class Table2Row {
        String name; double amount; String unit;
        Table2Row(String n, double a, String u) { name = n; amount = a; unit = u; }
    }
    private static class ResultRow {
        String name; double amountLeft; String idCard; String result;
        double diff; String nameRight; double amountRight; String remark; String unit;
        ResultRow(String n, double al, String id, String r, double d, String nr, double ar, String rm, String u) {
            name = n; amountLeft = al; idCard = id; result = r; diff = d;
            nameRight = nr; amountRight = ar; remark = rm; unit = u;
        }
    }
}
