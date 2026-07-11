package com.example.project.service.finance.cxm;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 数据核对器：对比待核对表与基准表的金额
 * 对应 Python: hs_checker.py -> DataChecker 类
 * <p>
 * 使用方式：
 * 1. 加载基准表（供应商汇总结果）
 * 2. 加载待核对文件（含月份、介绍人、金额列）
 * 3. 获取月份和介绍人列表供前端展示
 * 4. 执行核对，返回结果
 */
@Service




public class HsHuizongCxmChecker {

    private static final Logger log = LoggerFactory.getLogger(HsHuizongCxmChecker.class);

    // 月份提取正则（单月份，如"1月"）
    private static final Pattern MONTH_PATTERN = Pattern.compile("(?<!\\d)(\\d{1,2}月)(?!\\d)");
    // 范围月份（如"1-2月"）
    private static final Pattern RANGE_PATTERN = Pattern.compile("(?<!\\d)(\\d{1,2}-\\d{1,2}月)(?!\\d)");
    // 空白字符正则
    private static final Pattern BLANK_PATTERN = Pattern.compile("[\\s\\u200b\\u200c\\u200d\\u00a0\\u3000]+");

    /** 基准表数据：供应商名称（基础名称）-> 总计金额 */
    private Map<String, Double> referenceData = new HashMap<>();
    /** 基准表原始 DataFrame 等效（供应商名称 -> 总计） */
    private Map<String, Double> supplierTotals = new HashMap<>();

    // ========== 加载基准表 ==========

    /**
     * 从上传的Excel文件加载基准表（供应商名称 -> 总计）
     * 对应 Python: DataChecker.load_reference()
     */
    public void loadReference(MultipartFile file) throws Exception {
        Map<String, Double> result = new HashMap<>();
        File temp = null;
        try {
            temp = File.createTempFile("ref_", ".xlsx");
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(temp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            loadReferenceFromFile(temp, result);
        } finally {
            if (temp != null && temp.exists()) temp.delete();
        }
        this.supplierTotals = result;
        this.referenceData = result;
    }

    /**
     * 从本地文件加载基准表
     */
    public void loadReference(File file) throws Exception {
        Map<String, Double> result = new HashMap<>();
        loadReferenceFromFile(file, result);
        this.supplierTotals = result;
        this.referenceData = result;
    }

    /**
     * 从文件读取基准表数据的公共逻辑
     */
    private void loadReferenceFromFile(File file, Map<String, Double> result) throws Exception {
        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) throw new Exception("基准表未找到表头行");

            int supplierIdx = -1, totalIdx = -1;
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                String h = getCellStringValue(headerRow.getCell(c));
                if ("供应商名称".equals(h)) supplierIdx = c;
                if ("总计".equals(h)) totalIdx = c;
            }
            if (supplierIdx < 0) throw new Exception("基准表未找到'供应商名称'列");
            if (totalIdx < 0) throw new Exception("基准表未找到'总计'列");

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String supplier = getCellStringValue(row.getCell(supplierIdx)).trim();
                if (supplier.isEmpty()) continue;
                double total = getCellNumericValue(row.getCell(totalIdx));
                // 使用基础名称作为key（去除末尾数字）
                String baseName = getBaseName(supplier);
                result.merge(baseName, total, Double::sum);
            }
        }
    }

    /**
     * 从内存中的汇总结果加载基准表
     * 对应 Python: DataChecker.load_reference_from_dataframe()
     */
    public void loadReferenceFromSummary(List<Map<String, Object>> summaryData) throws Exception {
        Map<String, Double> result = new HashMap<>();
        for (Map<String, Object> row : summaryData) {
            Object supplierObj = row.get("供应商名称");
            if (supplierObj == null) continue;
            String supplier = supplierObj.toString().trim();
            if (supplier.isEmpty()) continue;

            double total = 0.0;
            Object totalObj = row.get("总计");
            if (totalObj instanceof Number) {
                total = ((Number) totalObj).doubleValue();
            }

            String baseName = getBaseName(supplier);
            result.merge(baseName, total, Double::sum);
        }
        this.supplierTotals = result;
        this.referenceData = result;
    }

    // ========== 检测列 ==========

    /**
     * 从待核对文件中检测月份列、介绍人列、金额列
     * 对应 Python: DataChecker.detect_columns() 等效逻辑
     *
     * @return Map 包含 monthCol, personCol, amountCol（列名），以及 months, persons
     */
    public Map<String, Object> detectCheckColumns(MultipartFile file, String sheetName) throws Exception {
        log.info("开始检测待核对文件的列信息，工作表: {}", sheetName);
        Map<String, Object> result = new HashMap<>();
        File temp = null;
        try {
            temp = File.createTempFile("check_", ".xlsx");
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = file.getInputStream();
                 OutputStream out = new FileOutputStream(temp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }
            try (Workbook wb = WorkbookFactory.create(temp)) {
                Sheet sheet = findSheetByName(wb, sheetName);
                if (sheet == null) {
                    log.error("工作表不存在: {}", sheetName);
                    throw new Exception("工作表不存在: " + sheetName);
                }

                // 自动检测表头行（前5行内查找同时包含"月份"、"介绍人"、"金额"的行）
                Row headerRow = null;
                List<String> headers = new ArrayList<>();
                int headerRowIdx = -1;
                for (int rowIdx = 0; rowIdx < Math.min(5, sheet.getLastRowNum() + 1); rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;
                    List<String> rowHeaders = new ArrayList<>();
                    boolean hasMonth = false, hasPerson = false, hasAmount = false;
                    for (int c = 0; c < row.getLastCellNum(); c++) {
                        String h = getCellStringValue(row.getCell(c));
                        rowHeaders.add(h);
                        if (!hasMonth && h != null && h.contains("月份")) hasMonth = true;
                        if (!hasPerson && h != null && h.contains("介绍人")) hasPerson = true;
                        if (!hasAmount && h != null && h.contains("金额")) hasAmount = true;
                    }
                    if (hasMonth && hasPerson && hasAmount) {
                        headerRow = row;
                        headers = rowHeaders;
                        headerRowIdx = rowIdx;
                        log.info("检测到表头行在第 {} 行，表头内容: {}", headerRowIdx + 1, headers);
                        break;
                    }
                }
                if (headerRow == null) {
                    throw new Exception("未在表头中找到'月份'、'介绍人'、'金额'列，请确认表头在前5行内");
                }

                List<String> columns = headers;
                int monthColIdx = -1, personColIdx = -1, amountColIdx = -1;
                for (int i = 0; i < headers.size(); i++) {
                    String h = headers.get(i);
                    if (monthColIdx < 0 && h != null && h.contains("月份")) monthColIdx = i;
                    if (personColIdx < 0 && h != null && h.contains("介绍人")) personColIdx = i;
                    if (amountColIdx < 0 && h != null && h.contains("金额")) amountColIdx = i;
                }

                result.put("monthCol", monthColIdx >= 0 ? columns.get(monthColIdx) : null);
                result.put("personCol", personColIdx >= 0 ? columns.get(personColIdx) : null);
                result.put("amountCol", amountColIdx >= 0 ? columns.get(amountColIdx) : null);
                result.put("monthColIdx", monthColIdx);
                result.put("personColIdx", personColIdx);
                result.put("amountColIdx", amountColIdx);
                result.put("columns", columns);

                // 数据从表头行的下一行开始
                int dataStartRow = headerRowIdx + 1;

                // 获取唯一月份
                List<String> months = getUniqueMonths(sheet, monthColIdx, dataStartRow);
                result.put("months", months);
                log.info("检测到月份: {}", months);

                // 获取唯一介绍人
                List<String> persons = getUniquePersons(sheet, personColIdx, dataStartRow);
                result.put("persons", persons);
                log.info("检测到介绍人数量: {}", persons.size());

                // 构建月份-介绍人映射
                Map<String, List<String>> monthPersonMap = new LinkedHashMap<>();
                for (String m : months) {
                    monthPersonMap.put(m, new ArrayList<>());
                }
                for (int r = dataStartRow; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;
                    String monthVal = monthColIdx >= 0 ? getCellStringValue(row.getCell(monthColIdx)) : "";
                    String personVal = personColIdx >= 0 ? getCellStringValue(row.getCell(personColIdx)) : "";
                    if (personVal.isEmpty()) continue;
                    if (personVal.matches("^0+(\\.0+)?$")) continue;
                    for (String m : months) {
                        if (matchMonthInText(m, monthVal)) {
                            if (!monthPersonMap.get(m).contains(personVal)) {
                                monthPersonMap.get(m).add(personVal);
                            }
                        }
                    }
                }
                // 对每个月份的人员列表排序
                for (String m : monthPersonMap.keySet()) {
                    Collections.sort(monthPersonMap.get(m));
                }
                result.put("monthPersonMap", monthPersonMap);
                log.info("构建月份-介绍人映射完成，月份数: {}, 总介绍人数: {}",
                        monthPersonMap.size(),
                        monthPersonMap.values().stream().mapToInt(List::size).sum());
                result.put("success", true);
            }
        } finally {
            if (temp != null && temp.exists()) temp.delete();
        }
        log.info("检测完成，返回结果");
        return result;
    }

    // ========== 执行核对 ==========

    /**
     * 执行核对：按介绍人分组累加金额，与基准表总计对比
     * 对应 Python: DataChecker.check()
     *
     * @param checkFile       待核对文件
     * @param sheetName       工作表名
     * @param selectedMonth   选择的月份
     * @param selectedPersons 选择的介绍人（列表）
     * @param outputDir      输出目录
     * @return 核对结果文件路径
     */
    public String executeCheck(MultipartFile checkFile, String sheetName,
                                      String selectedMonth, List<String> selectedPersons,
                                      String outputDir) throws Exception {
        log.info("开始执行核对，工作表: {}, 月份: {}, 介绍人: {}", sheetName, selectedMonth, selectedPersons);
        File temp = null;
        try {
            temp = File.createTempFile("check_exec_", ".xlsx");
            // 使用流复制避免 Tomcat 临时文件清理导致的 FileNotFoundException
            try (InputStream in = checkFile.getInputStream();
                 OutputStream out = new FileOutputStream(temp)) {
                byte[] buf = new byte[8192];
                int n;
                while ((n = in.read(buf)) != -1) out.write(buf, 0, n);
            }

            // 读取待核对数据
            List<Map<String, Object>> checkData;
            String monthColName, personColName, amountColName;
            try (Workbook wb = WorkbookFactory.create(temp)) {
                Sheet sheet = findSheetByName(wb, sheetName);
                if (sheet == null) throw new Exception("工作表不存在: " + sheetName);

                // 自动检测表头行（前5行内查找同时包含"月份"、"介绍人"、"金额"的行）
                Row headerRow = null;
                List<String> columns = new ArrayList<>();
                int headerRowIdx = -1;
                for (int rowIdx = 0; rowIdx < Math.min(5, sheet.getLastRowNum() + 1); rowIdx++) {
                    Row row = sheet.getRow(rowIdx);
                    if (row == null) continue;
                    List<String> rowHeaders = new ArrayList<>();
                    boolean hasMonth = false, hasPerson = false, hasAmount = false;
                    for (int c = 0; c < row.getLastCellNum(); c++) {
                        String h = getCellStringValue(row.getCell(c));
                        rowHeaders.add(h);
                        if (!hasMonth && h != null && h.contains("月份")) hasMonth = true;
                        if (!hasPerson && h != null && h.contains("介绍人")) hasPerson = true;
                        if (!hasAmount && h != null && h.contains("金额")) hasAmount = true;
                    }
                    if (hasMonth && hasPerson && hasAmount) {
                        headerRow = row;
                        columns = rowHeaders;
                        headerRowIdx = rowIdx;
                        break;
                    }
                }
                if (headerRow == null) {
                    throw new Exception("未在表头中找到'月份'、'介绍人'、'金额'列，请确认表头在前5行内");
                }

                int monthColIdx = -1, personColIdx = -1, amountColIdx = -1;
                for (int i = 0; i < columns.size(); i++) {
                    String h = columns.get(i);
                    if (monthColIdx < 0 && h != null && h.contains("月份")) monthColIdx = i;
                    if (personColIdx < 0 && h != null && h.contains("介绍人")) personColIdx = i;
                    if (amountColIdx < 0 && h != null && h.contains("金额")) amountColIdx = i;
                }

                // 数据从表头行的下一行开始
                int dataStartRow = headerRowIdx + 1;
                checkData = new ArrayList<>();
                for (int r = dataStartRow; r <= sheet.getLastRowNum(); r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;
                    Map<String, Object> rowMap = new HashMap<>();
                    for (int c = 0; c < columns.size(); c++) {
                        rowMap.put(columns.get(c), getCellStringValue(row.getCell(c)));
                    }
                    // 金额作为double存储
                    if (amountColIdx >= 0) {
                        rowMap.put("_amount", getCellNumericValue(row.getCell(amountColIdx)));
                    }
                    // 月份值
                    if (monthColIdx >= 0) {
                        rowMap.put("_monthVal", getCellStringValue(row.getCell(monthColIdx)));
                    }
                    checkData.add(rowMap);
                }
                monthColName = monthColIdx >= 0 ? columns.get(monthColIdx) : "";
                personColName = personColIdx >= 0 ? columns.get(personColIdx) : "";
                amountColName = amountColIdx >= 0 ? columns.get(amountColIdx) : "";
            }

            // 筛选月份
            final String finalMonthColName = monthColName;
            List<Map<String, Object>> filtered = new ArrayList<>();
            for (Map<String, Object> row : checkData) {
                String monthVal = (String) row.get("_monthVal");
                if (matchMonthInText(selectedMonth, monthVal)) {
                    filtered.add(row);
                }
            }

            // 按介绍人分组汇总
            List<Map<String, Object>> results = new ArrayList<>();
            for (String person : selectedPersons) {
                // 查找基准表中匹配的基础名称
                String personBase = getBaseName(person);
                double baseTotal = 0.0;
                List<String> matchedSuppliers = new ArrayList<>();
                for (Map.Entry<String, Double> entry : supplierTotals.entrySet()) {
                    if (getBaseName(entry.getKey()).equals(personBase)) {
                        matchedSuppliers.add(entry.getKey());
                        baseTotal += entry.getValue();
                    }
                }

                // 计算待核对金额
                double checkAmount = 0.0;
                for (Map<String, Object> row : filtered) {
                    String rowPerson = (String) row.get(personColName);
                    if (person.equals(rowPerson)) {
                        Object amtObj = row.get("_amount");
                        if (amtObj instanceof Number) {
                            checkAmount += ((Number) amtObj).doubleValue();
                        }
                    }
                }

                String status;
                if (matchedSuppliers.isEmpty()) {
                    status = "未匹配";
                } else if (Math.abs(checkAmount - baseTotal) < 0.01) {
                    status = "已核对";
                } else {
                    status = "待核对";
                }

                Map<String, Object> resultRow = new LinkedHashMap<>();
                resultRow.put("介绍人", person);
                resultRow.put("待核对金额总计", Math.round(checkAmount * 100.0) / 100.0);
                resultRow.put("供应商名称", matchedSuppliers.isEmpty() ? "" : personBase);
                resultRow.put("对公打款总计", matchedSuppliers.isEmpty() ? "" :
                        (Math.round(baseTotal * 100.0) / 100.0));
                resultRow.put("状态", status);
                results.add(resultRow);
                log.info("介绍人: {}, 待核对金额: {}, 基准表金额: {}, 状态: {}",
                        person, Math.round(checkAmount * 100.0) / 100.0,
                        matchedSuppliers.isEmpty() ? "" : (Math.round(baseTotal * 100.0) / 100.0), status);
            }

            // 写出结果（固定文件名，与Python一致）
            File outputFile = new File(outputDir, "核对结果.xlsx");
            writeCheckResult(outputFile, results);
            log.info("核对完成，结果文件: {}", outputFile.getName());
            return outputFile.getAbsolutePath();

        } finally {
            if (temp != null && temp.exists()) temp.delete();
        }
    }

    // ========== 写出核对结果 ==========

    private void writeCheckResult(File outputFile, List<Map<String, Object>> results) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("核对结果");

            if (results.isEmpty()) {
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    wb.write(fos);
                }
                return;
            }

            // 获取所有列
            Set<String> colSet = new LinkedHashSet<>();
            for (Map<String, Object> row : results) colSet.addAll(row.keySet());
            List<String> columns = new ArrayList<>(colSet);

            // 写表头
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = createHeaderStyle(wb);
            for (int c = 0; c < columns.size(); c++) {
                Cell cell = headerRow.createCell(c);
                cell.setCellValue(columns.get(c));
                cell.setCellStyle(headerStyle);
            }

            // 写数据
            CellStyle dataStyle = createDataStyle(wb);
            CellStyle pendingStyle = createPendingStyle(wb);
            for (int r = 0; r < results.size(); r++) {
                Row row = sheet.createRow(r + 1);
                Map<String, Object> map = results.get(r);
                for (int c = 0; c < columns.size(); c++) {
                    Cell cell = row.createCell(c);
                    Object val = map.get(columns.get(c));
                    if (val instanceof Number) {
                        double d = ((Number) val).doubleValue();
                        // 金额为 0 时保留空白
                        if (Math.abs(d) < 0.0001) {
                            cell.setCellValue("");
                        } else {
                            cell.setCellValue(d);
                        }
                    } else if (val != null) {
                        cell.setCellValue(val.toString());
                    }
                    // 仅"状态"列且值为"待核对"的单元格标黄，其余用普通样式
                    if ("状态".equals(columns.get(c)) && "待核对".equals(val)) {
                        cell.setCellStyle(pendingStyle);
                    } else {
                        cell.setCellStyle(dataStyle);
                    }
                }
            }

            // 自动列宽
            for (int c = 0; c < columns.size(); c++) {
                sheet.autoSizeColumn(c);
                int currentWidth = sheet.getColumnWidth(c);
                sheet.setColumnWidth(c, Math.min(currentWidth + 1024, 50 * 256));
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // ========== 工具方法 ==========

    /**
     * 获取基础名称（去除末尾数字，如"供应商-张三1" -> "供应商-张三"）
     * 对应 Python: get_base_name()
     */
    private static String getBaseName(String name) {
        if (name == null) return "";
        return BLANK_PATTERN.matcher(name.trim()).replaceAll("")
                .replaceAll("\\d+$", "");
    }

    /**
     * 精确匹配月份字符串（避免"1月"匹配到"12月"）
     * 对应 Python: match_month_in_text()
     */
    private static boolean matchMonthInText(String month, String text) {
        if (text == null || text.isEmpty()) return false;
        Pattern p = Pattern.compile("(?<!\\d)" + Pattern.quote(month) + "(?!\\d)");
        return p.matcher(text).find();
    }

    private static List<String> getUniqueMonths(Sheet sheet, int monthColIdx) {
        return getUniqueMonths(sheet, monthColIdx, 1);
    }

    private static List<String> getUniqueMonths(Sheet sheet, int monthColIdx, int startRow) {
        Set<String> found = new LinkedHashSet<>();
        if (monthColIdx < 0) return new ArrayList<>();

        for (int r = startRow; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String val = getCellStringValue(row.getCell(monthColIdx));
            if (val == null || val.isEmpty()) continue;

            Matcher m1 = MONTH_PATTERN.matcher(val);
            while (m1.find()) found.add(m1.group(1));
            Matcher m2 = RANGE_PATTERN.matcher(val);
            while (m2.find()) found.add(m2.group(1));
        }

        // 排序：单月份在前，范围月份在后
        List<String> result = new ArrayList<>(found);
        result.sort((a, b) -> {
            boolean aRange = a.contains("-");
            boolean bRange = b.contains("-");
            if (aRange != bRange) return aRange ? 1 : -1;
            int aNum = Integer.parseInt(a.replace("月", "").split("-")[0]);
            int bNum = Integer.parseInt(b.replace("月", "").split("-")[0]);
            return Integer.compare(aNum, bNum);
        });
        return result;
    }

    private static List<String> getUniquePersons(Sheet sheet, int personColIdx) {
        return getUniquePersons(sheet, personColIdx, 1);
    }

    private static List<String> getUniquePersons(Sheet sheet, int personColIdx, int startRow) {
        Set<String> found = new LinkedHashSet<>();
        if (personColIdx < 0) return new ArrayList<>();

        for (int r = startRow; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String val = getCellStringValue(row.getCell(personColIdx));
            if (val == null || val.isEmpty()) continue;
            if (val.startsWith("#") || val.toLowerCase().matches("nan|none|null")) continue;
            // 排除纯零数值（如 0, 0.0, 0.00 等 Excel  numeric 空值）
            if (val.matches("^0+(\\.0+)?$")) continue;
            found.add(val.trim());
        }
        return new ArrayList<>(found);
    }

    /**
     * 模糊匹配工作表名称（去除隐藏字符后精确匹配，或忽略大小写包含匹配）
     */
    private static Sheet findSheetByName(Workbook wb, String sheetName) {
        if (sheetName == null || sheetName.isEmpty()) return wb.getSheetAt(0);
        String cleanName = BLANK_PATTERN.matcher(sheetName.trim()).replaceAll("");
        // 第一遍：精确匹配（去除隐藏字符后）
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (BLANK_PATTERN.matcher(name.trim()).replaceAll("").equals(cleanName)) {
                return wb.getSheetAt(i);
            }
        }
        // 第二遍：忽略大小写包含匹配
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (name.toLowerCase().contains(sheetName.toLowerCase()) ||
                sheetName.toLowerCase().contains(name.toLowerCase())) {
                return wb.getSheetAt(i);
            }
        }
        return null;
    }

    private static String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        String val;
        switch (cell.getCellType()) {
            case STRING: val = cell.getStringCellValue(); break;
            case NUMERIC: val = String.valueOf(cell.getNumericCellValue()); break;
            case BOOLEAN: val = String.valueOf(cell.getBooleanCellValue()); break;
            case FORMULA:
                CellType resultType = cell.getCachedFormulaResultType();
                switch (resultType) {
                    case STRING: val = cell.getStringCellValue(); break;
                    case NUMERIC: val = String.valueOf(cell.getNumericCellValue()); break;
                    case BOOLEAN: val = String.valueOf(cell.getBooleanCellValue()); break;
                    default: return "";
                }
                break;
            default: return "";
        }
        // 去除首尾空白及不可见字符（零宽空格等）
        return BLANK_PATTERN.matcher(val.trim()).replaceAll("");
    }

    private static double getCellNumericValue(Cell cell) {
        if (cell == null) return 0.0;
        switch (cell.getCellType()) {
            case NUMERIC: return cell.getNumericCellValue();
            case STRING:
                try { return Double.parseDouble(cell.getStringCellValue().trim()); }
                catch (NumberFormatException e) { return 0.0; }
            case FORMULA:
                CellType resultType = cell.getCachedFormulaResultType();
                if (resultType == CellType.NUMERIC) {
                    return cell.getNumericCellValue();
                } else if (resultType == CellType.STRING) {
                    try { return Double.parseDouble(cell.getStringCellValue().trim()); }
                    catch (NumberFormatException e) { return 0.0; }
                }
                return 0.0;
            default: return 0.0;
        }
    }

    // ========== Excel 样式 ==========

    private static CellStyle createHeaderStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        org.apache.poi.ss.usermodel.Font font = wb.createFont();
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);
        // #4472c4 = RGB(68, 114, 196)
        ((XSSFCellStyle)style).setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x44, (byte)0x72, (byte)0xC4}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createDataStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createPendingStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
}
