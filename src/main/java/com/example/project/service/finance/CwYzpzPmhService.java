package com.example.project.service.finance;

import com.example.project.util.ExcelUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 对应 Python: cw_yzpz_pmh.py (ExcelProcessor) + cw_yzpz_pmh_main.py (GUI)
 * 功能：预支平账
 *
 * 核心流程（多轮核对）：
 * 1. prepare:  加载文件 → 创建核对工作表(Sheet3/4/5...)
 * 2. copyNames: 从外部文件复制姓名字段到姓名表
 * 3. match:     姓名匹配 → 筛选平账额为空的行 → 写入数据表
 * 4. processDeduction: 处理需扣回/预支扣款/预支扣回数据 → 汇总对比
 * 5. balance:   平账 → 更新工资表 → 金额拆分 → 标红 → 保存记录
 */
@Service
public class CwYzpzPmhService {

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy.M.d");
    // 最大轮次
    private static final int MAX_ROUND = 30;

    // Sheet4 列索引常量（0-based，与Python openpyxl的1-based对应）
    private static final int COL_ADVANCE_NAME      = 0;  // A列: 姓名（匹配后）
    private static final int COL_ADVANCE_VALUE     = 1;  // B列: 预支数额
    private static final int COL_DED_NAME          = 4;  // E列: 扣款姓名
    private static final int COL_DED_VALUE         = 5;  // F列: 扣款金额
    private static final int COL_ADVANCE_SUM_NAME  = 8;  // I列: 汇总姓名
    private static final int COL_ADVANCE_SUM_VALUE = 9;  // J列: 求和项:预支数额
    private static final int COL_COMPARE           = 10; // K列: 姓名是否相同
    private static final int COL_SUM               = 11; // L列: 求和(J+N)
    private static final int COL_DED_SUM_NAME      = 12; // M列: 汇总姓名(复制E列)
    private static final int COL_DED_SUM_VALUE     = 13; // N列: 求和项:扣款
    private static final int COL_TYPE_MARKER       = 51; // AZ列: 列类型标记

    // ==================== 准备阶段 ====================

    /**
     * 准备阶段：创建核对工作表
     * @param file        主工资表文件
     * @param targetSheet 目标工作表名
     * @param outputDir   输出目录
     * @param round       当前核对轮次（1-base）
     * @return 处理后的文件路径
     */
    public String prepare(MultipartFile file, String targetSheet, int round, String outputDir) throws Exception {
        if (file == null || file.isEmpty()) throw new IllegalArgumentException("文件不能为空");
        if (round < 1 || round > MAX_ROUND) throw new IllegalArgumentException("轮次范围: 1-" + MAX_ROUND);
        if (targetSheet == null || targetSheet.isEmpty()) throw new IllegalArgumentException("请指定工作表名");

        File inputFile = saveToTemp(file);
        try {
            File outputFile = new File(outputDir, "yzpz_prepared_" + System.currentTimeMillis() + ".xlsx");
            doPrepare(inputFile, outputFile, targetSheet, round);
            return outputFile.getAbsolutePath();
        } finally {
            if (inputFile.exists()) inputFile.delete();
        }
    }

    private void doPrepare(File inputFile, File outputFile, String targetSheet, int round) throws Exception {
        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheet(targetSheet);
            if (sheet == null) throw new Exception("未找到工作表: " + targetSheet);

            // 创建核对工作表（第 round 轮: Sheet3/4, Sheet5/6, Sheet7/8...）
            String namesSheetName = roundSheetName(round, 0); // Sheet3/5/7...
            String dataSheetName  = roundSheetName(round, 1); // Sheet4/6/8...
            int dataNum = 3 + (round - 1) * 2 + 1; // 用于清理循环

            // 重建当前轮次的 Sheet（确保干净）
            if (wb.getSheet(namesSheetName) != null) {
                wb.removeSheetAt(wb.getSheetIndex(namesSheetName));
            }
            if (wb.getSheet(dataSheetName) != null) {
                wb.removeSheetAt(wb.getSheetIndex(dataSheetName));
            }
            wb.createSheet(namesSheetName);
            wb.createSheet(dataSheetName);

            // 清理比当前轮次更高的核对 Sheet（防止误操作产生的空 Sheet）
            int minSheetToRemove = dataNum + 1; // Sheet7, Sheet8, ...
            List<String> toRemove = new ArrayList<>();
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String sName = wb.getSheetName(i);
                if (sName.startsWith("Sheet")) {
                    try {
                        int num = Integer.parseInt(sName.substring(5));
                        if (num >= minSheetToRemove) {
                            toRemove.add(sName);
                        }
                    } catch (NumberFormatException e) {
                        // 忽略非数字后缀的 Sheet 名
                    }
                }
            }
            // 按编号降序排序，从后往前删除避免索引变化
            toRemove.sort((a, b) -> {
                int na = Integer.parseInt(a.substring(5));
                int nb = Integer.parseInt(b.substring(5));
                return Integer.compare(nb, na);
            });
            for (String sName : toRemove) {
                int idx = wb.getSheetIndex(sName);
                if (idx >= 0) wb.removeSheetAt(idx);
            }

            // 保存
            try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
        }
    }

    // ==================== 复制姓名 ====================

    /**
     * 从外部核对文件复制姓名到主文件
     * @param mainFile       主工资表文件
     * @param checkFile      外部核对文件
     * @param checkSheet     外部核对文件的工作表名
     * @param round          当前轮次
     * @param outputDir      输出目录
     * @return 处理后的文件路径
     */
    public String copyNames(MultipartFile mainFile, MultipartFile checkFile,
                             String checkSheet, int round, String outputDir) throws Exception {
        File mainTemp = saveToTemp(mainFile);
        File checkTemp = saveToTemp(checkFile);
        try {
            File outputFile = new File(outputDir, "yzpz_names_" + System.currentTimeMillis() + ".xlsx");
            doCopyNames(mainTemp, checkTemp, checkSheet, round, outputFile);
            return outputFile.getAbsolutePath();
        } finally {
            if (mainTemp.exists()) mainTemp.delete();
            if (checkTemp.exists()) checkTemp.delete();
        }
    }

    private void doCopyNames(File mainFile, File checkFile, String checkSheet, int round, File outputFile) throws Exception {
        try (FileInputStream mainFis = new FileInputStream(mainFile);
             Workbook mainWb = new XSSFWorkbook(mainFis);
             FileInputStream checkFis = new FileInputStream(checkFile);
             Workbook checkWb = new XSSFWorkbook(checkFis)) {

            String namesSheetName = roundSheetName(round, 0);

            // 获取或重建姓名表
            if (mainWb.getSheet(namesSheetName) != null) {
                mainWb.removeSheetAt(mainWb.getSheetIndex(namesSheetName));
            }
            Sheet namesSheet = mainWb.createSheet(namesSheetName);

            // 从核对文件查找"姓名"列
            Sheet srcSheet = checkWb.getSheet(checkSheet);
            if (srcSheet == null) {
                if (checkWb.getNumberOfSheets() > 0) srcSheet = checkWb.getSheetAt(0);
                else throw new Exception("核对文件没有工作表");
            }

            Row srcHeader = getHeaderRow(srcSheet);

            int nameCol = findColumnByHeader(srcHeader, "姓名");

            if (nameCol < 0) throw new Exception("未找到'姓名'列");

            // 复制姓名
            int destRow = 1;
            for (int r = 2; r <= srcSheet.getLastRowNum(); r++) { // 从第3行开始（0-based: 2）
                Row row = srcSheet.getRow(r);
                if (row == null) continue;
                String name = ExcelUtils.getCellStringValue(row.getCell(nameCol));
                if (name != null && !name.trim().isEmpty()) {
                    namesSheet.createRow(destRow - 1).createCell(0).setCellValue(name.trim());
                    destRow++;
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) { mainWb.write(fos); }
        }
    }

    // ==================== 姓名匹配 ====================

    /**
     * 姓名匹配：根据姓名表筛选数据到数据表
     * @param mainFile    主工资表文件
     * @param mergedSheet 合并工作表名
     * @param units       单位筛选（逗号分隔，空=不筛选）
     * @param round       当前轮次
     * @param outputDir   输出目录
     * @return {file, matchedUnits} 或 单位列表供前端选择
     */
    public Map<String, Object> matchNames(MultipartFile mainFile, String mergedSheet,
                                           String units, int round, String outputDir) throws Exception {
        File mainTemp = saveToTemp(mainFile);
        try {
            File outputFile = new File(outputDir, "yzpz_matched_" + System.currentTimeMillis() + ".xlsx");
            Map<String, Object> result = doMatchNames(mainTemp, mergedSheet, units, round, outputFile);
            result.put("fileName", outputFile.getName());
            return result;
        } finally {
            if (mainTemp.exists()) mainTemp.delete();
        }
    }

    private Map<String, Object> doMatchNames(File mainFile, String mergedSheet,
                                              String units, int round, File outputFile) throws Exception {
        Map<String, Object> result = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(mainFile);
             Workbook wb = new XSSFWorkbook(fis)) {

            String namesSheetName = roundSheetName(round, 0);
            String dataSheetName  = roundSheetName(round, 1);

            Sheet namesSheet = wb.getSheet(namesSheetName);
            if (namesSheet == null) throw new Exception("姓名表 " + namesSheetName + " 不存在，请先复制姓名");

            Sheet merged = wb.getSheet(mergedSheet);
            if (merged == null) throw new Exception("未找到工作表: " + mergedSheet);

            Row mergedHeader = getHeaderRow(merged);

            // 查找列
            int nameCol = findColumnByHeader(mergedHeader, "姓名");
            int unitCol = findColumnByHeader(mergedHeader, "单位");
            int balanceCol = findColumnContains(mergedHeader, "平账额");
            int advanceCol = findColumnContains(mergedHeader, "预支数额");

            if (nameCol < 0) throw new Exception("未找到'姓名'列");
            if (balanceCol < 0) throw new Exception("未找到'平账额'列");
            if (advanceCol < 0) throw new Exception("未找到'预支数额'列");
            int maxRow = merged.getLastRowNum();

            // 收集姓名表中的姓名
            Set<String> nameSet = new HashSet<>();
            for (int r = 0; r <= namesSheet.getLastRowNum(); r++) {
                Row row = namesSheet.getRow(r);
                if (row != null) {
                    String n = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_NAME));
                    if (n != null && !n.trim().isEmpty()) nameSet.add(n.trim());
                }
            }

            // 扫描匹配行并收集单位
            Set<String> matchedUnits = new TreeSet<>();
            List<Integer> matchedRows = new ArrayList<>();

            for (int r = 2; r <= maxRow; r++) { // 从第3行开始
                Row row = merged.getRow(r);
                if (row == null) continue;
                String name = ExcelUtils.getCellStringValue(row.getCell(nameCol));
                if (name == null || !nameSet.contains(name.trim())) continue;

                // 单位筛选
                if (!matchesUnitFilter(row, unitCol, units)) continue;

                // 收集单位
                if (unitCol >= 0) {
                    String uv = ExcelUtils.getCellStringValue(row.getCell(unitCol));
                    if (uv != null && !uv.trim().isEmpty()) matchedUnits.add(uv.trim());
                }

                // 筛选平账额为空的行
                String balance = ExcelUtils.getCellStringValue(row.getCell(balanceCol));
                if (balance == null || balance.trim().isEmpty()) {
                    matchedRows.add(r);
                }
            }

            // 如果用户未指定单位且有多个单位，把单位列表返回给前端
            if (units == null && matchedUnits.size() > 1) {
                result.put("success", true);
                result.put("units", new ArrayList<>(matchedUnits));
                result.put("needSelectUnits", true);
                result.put("message", String.join("|", matchedUnits));
                return result;
            }

            // 写入数据表
            if (wb.getSheet(dataSheetName) != null) {
                wb.removeSheetAt(wb.getSheetIndex(dataSheetName));
            }
            Sheet dataSheet = wb.createSheet(dataSheetName);
            // 标题行
            dataSheet.createRow(0).createCell(COL_ADVANCE_NAME).setCellValue("姓名");
            dataSheet.getRow(0).createCell(COL_ADVANCE_VALUE).setCellValue("预支数额");

            int destRow = 1;
            for (int rowIndex : matchedRows) {
                Row srcRow = merged.getRow(rowIndex);
                Row dest = dataSheet.createRow(destRow++);
                // A列：姓名（字符串）
                dest.createCell(COL_ADVANCE_NAME).setCellValue(ExcelUtils.getCellStringValue(srcRow.getCell(nameCol)));
                // B列：预支数额（数值类型）
                Cell advCell = srcRow.getCell(advanceCol);
                if (advCell != null) {
                    if (advCell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                        dest.createCell(COL_ADVANCE_VALUE).setCellValue(advCell.getNumericCellValue());
                    } else {
                        String val = ExcelUtils.getCellStringValue(advCell);
                        dest.createCell(COL_ADVANCE_VALUE).setCellValue(parseDouble(val));
                    }
                } else {
                    dest.createCell(COL_ADVANCE_VALUE).setCellValue(0.0);
                }
            }

            result.put("success", true);
            result.put("matchCount", matchedRows.size());
            result.put("message", "已匹配 " + matchedRows.size() + " 条数据");

            try (FileOutputStream fos = new FileOutputStream(outputFile)) { wb.write(fos); }
            result.put("filePath", outputFile.getAbsolutePath());
            return result;
        }
    }

    // ==================== 处理扣款数据 ====================

    /**
     * 处理需扣回/预支扣款/预支扣回数据
     * @param mainFile     主工资表文件
     * @param dedFile      扣款核对文件
     * @param dedSheet     扣款文件工作表
     * @param columnType   列类型: 需扣回/预支扣款/预支扣回 (null=自动)
     * @param round        当前轮次
     * @param outputDir    输出目录
     * @return 处理后的文件路径
     */
    public String processDeduction(MultipartFile mainFile, MultipartFile dedFile,
                                    String dedSheet, String columnType,
                                    int round, String outputDir) throws Exception {
        File mainTemp = saveToTemp(mainFile);
        File dedTemp = saveToTemp(dedFile);
        try {
            File outputFile = new File(outputDir, "yzpz_deduction_" + System.currentTimeMillis() + ".xlsx");
            doProcessDeduction(mainTemp, dedTemp, dedSheet, columnType, round, outputFile);
            return outputFile.getAbsolutePath();
        } finally {
            if (mainTemp.exists()) mainTemp.delete();
            if (dedTemp.exists()) dedTemp.delete();
        }
    }

    private void doProcessDeduction(File mainFile, File dedFile,
                                     String dedSheet, String columnType, int round, File outputFile) throws Exception {
        try (FileInputStream mainFis = new FileInputStream(mainFile);
             Workbook mainWb = new XSSFWorkbook(mainFis);
             FileInputStream dedFis = new FileInputStream(dedFile);
             Workbook dedWb = new XSSFWorkbook(dedFis)) {

            // 按轮次计算数据表名称
            String dataSheetName = roundSheetName(round, 1);

            Sheet dataSheet = mainWb.getSheet(dataSheetName);
            if (dataSheet == null) throw new Exception(dataSheetName + " 不存在，请先完成姓名匹配");

            Sheet srcSheet = dedWb.getSheet(dedSheet);
            if (srcSheet == null) {
                if (dedWb.getNumberOfSheets() > 0) srcSheet = dedWb.getSheetAt(0);
                else throw new Exception("扣款文件没有工作表");
            }

            Row srcHeader = getHeaderRow(srcSheet);

            // 查找列
            int nameCol = findColumnByHeader(srcHeader, "姓名");
            int deductionCol = findColumnContains(srcHeader, "需扣回");
            int advDedCol = findColumnContains(srcHeader, "预支扣款");
            int advRetCol = findColumnContains(srcHeader, "预支扣回");

            if (nameCol < 0) throw new Exception("未找到'姓名'列");
            if (deductionCol < 0 && advDedCol < 0 && advRetCol < 0)
                throw new Exception("未找到'需扣回'、'预支扣款'或'预支扣回'列");

            // 探测有数据的列
            List<int[]> availableCols = new ArrayList<>();
            int[] cols = {deductionCol, advDedCol, advRetCol};
            String[] colNames = {"需扣回", "预支扣款", "预支扣回"};
            for (int i = 0; i < cols.length; i++) {
                if (cols[i] < 0) continue;
                for (int r = 2; r <= srcSheet.getLastRowNum(); r++) {
                    Row row = srcSheet.getRow(r);
                    if (row != null) {
                        String v = ExcelUtils.getCellStringValue(row.getCell(cols[i]));
                        if (v != null && !v.trim().isEmpty()) {
                            availableCols.add(new int[]{cols[i], i});
                            break;
                        }
                    }
                }
            }

            if (availableCols.isEmpty()) throw new Exception("没有有效扣款数据");

            // 选择目标列
            int targetCol;
            String targetColName;
            if (availableCols.size() > 1 && columnType != null) {
                int idx = -1;
                for (int i = 0; i < colNames.length; i++) {
                    if (colNames[i].equals(columnType)) { idx = i; break; }
                }
                if (idx < 0) throw new Exception("未找到列类型: " + columnType);
                int finalIdx = idx;
                targetCol = availableCols.stream().filter(a -> a[1] == finalIdx).findFirst()
                        .orElseThrow(() -> new Exception("未找到匹配列"))[0];
                targetColName = columnType;
            } else {
                targetCol = availableCols.get(0)[0];
                targetColName = colNames[availableCols.get(0)[1]];
            }

            // 收集扣款数据
            List<String[]> dedRows = new ArrayList<>();
            for (int r = 2; r <= srcSheet.getLastRowNum(); r++) {
                Row row = srcSheet.getRow(r);
                if (row == null) continue;
                String name = ExcelUtils.getCellStringValue(row.getCell(nameCol));
                String val = ExcelUtils.getCellStringValue(row.getCell(targetCol));
                if (name != null && !name.trim().isEmpty() && !"总计".equals(name.trim())
                        && val != null && !val.trim().isEmpty()) {
                    dedRows.add(new String[]{name.trim(), val});
                }
            }

            if (dedRows.isEmpty()) throw new Exception("没有有效的扣款数据");

            // 在 AZ1 保存列类型标记
            Row azRow = dataSheet.getRow(0);
            if (azRow == null) azRow = dataSheet.createRow(0);
            azRow.createCell(COL_TYPE_MARKER).setCellValue(targetColName);

            // ========== 步骤1：写入扣款数据到第5、6列 ==========
            dataSheet.getRow(0).createCell(COL_DED_NAME).setCellValue("姓名");
            dataSheet.getRow(0).createCell(COL_DED_VALUE).setCellValue(targetColName);
            for (int i = 0; i < dedRows.size(); i++) {
                Row r = getOrCreateRow(dataSheet, i + 1);
                r.createCell(COL_DED_NAME).setCellValue(dedRows.get(i)[0]);
                r.createCell(COL_DED_VALUE).setCellValue(dedRows.get(i)[1]);
            }

            // ========== 步骤2：从第1-2列汇总数据到 name_advance_dict ==========
            Map<String, Double> advanceSumMap = new LinkedHashMap<>();
            for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;
                String n = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_NAME));
                if (n == null || n.trim().isEmpty() || "总计".equals(n.trim())) continue;
                String v = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_VALUE));
                if (v != null && !v.trim().isEmpty()) {
                    advanceSumMap.merge(n.trim(), parseDouble(v), Double::sum);
                }
            }

            // ========== 步骤3：设置第13、14列标题，复制第5、6列到第13、14列 ==========
            dataSheet.getRow(0).createCell(COL_DED_SUM_NAME).setCellValue("姓名");
            dataSheet.getRow(0).createCell(COL_DED_SUM_VALUE).setCellValue("求和项：" + targetColName);

            int maxRow13_14 = 1;
            double totalDeduction = 0;
            for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;
                String n = ExcelUtils.getCellStringValue(row.getCell(COL_DED_NAME));
                if (n == null || n.trim().isEmpty() || "总计".equals(n.trim())) continue;
                String v = ExcelUtils.getCellStringValue(row.getCell(COL_DED_VALUE));
                row.createCell(COL_DED_SUM_NAME).setCellValue(n);
                row.createCell(COL_DED_SUM_VALUE).setCellValue(v);
                maxRow13_14 = r;
                totalDeduction += parseDouble(v);
            }

            // ========== 步骤4：设置第9、10列标题，按第13列顺序写入第9、10列 ==========
            dataSheet.getRow(0).createCell(COL_ADVANCE_SUM_NAME).setCellValue("姓名");
            dataSheet.getRow(0).createCell(COL_ADVANCE_SUM_VALUE).setCellValue("求和项：预支数额");

            int maxRow9_10 = 1;
            double totalAdvance = 0;
            for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;
                String name13 = ExcelUtils.getCellStringValue(row.getCell(COL_DED_SUM_NAME));
                if (name13 == null || name13.trim().isEmpty() || "总计".equals(name13.trim())) continue;
                String nameTrim = name13.trim();
                Cell mCell = row.getCell(COL_DED_SUM_NAME);
                if (mCell != null) {
                    if (mCell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                        row.createCell(COL_ADVANCE_SUM_NAME).setCellValue(mCell.getNumericCellValue());
                    } else {
                        String mValue = ExcelUtils.getCellStringValue(mCell);
                        row.createCell(COL_ADVANCE_SUM_NAME).setCellValue(mValue);
                    }
                } else {
                    row.createCell(COL_ADVANCE_SUM_NAME).setCellValue(name13);
                }
                double adv = advanceSumMap.getOrDefault(nameTrim, 0.0);
                row.createCell(COL_ADVANCE_SUM_VALUE).setCellValue(adv);
                maxRow9_10 = r;
                totalAdvance += adv;
            }

            // ========== 步骤5：确定最大数据行和总计行位置 ==========
            int maxDataRow = Math.max(maxRow9_10, maxRow13_14);
            int totalRow = maxDataRow + 1;

            if (maxDataRow > 1) {
                Row totalRowObj = getOrCreateRow(dataSheet, totalRow);
                totalRowObj.createCell(COL_ADVANCE_SUM_NAME).setCellValue("总计");
                totalRowObj.createCell(COL_ADVANCE_SUM_VALUE).setCellValue(totalAdvance);
                totalRowObj.createCell(COL_DED_SUM_NAME).setCellValue("总计");
                totalRowObj.createCell(COL_DED_SUM_VALUE).setCellValue(totalDeduction);
            }

            // ========== 步骤6：第11列比较 I列 和 M列 姓名 ==========
            for (int r = 1; r < totalRow; r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;
                String name9 = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_SUM_NAME));
                String name13 = ExcelUtils.getCellStringValue(row.getCell(COL_DED_SUM_NAME));
                if (name9 != null && name13 != null) {
                    row.createCell(COL_COMPARE).setCellValue(name9.trim().equals(name13.trim()));
                }
            }

            // ========== 步骤7：第12列计算 J列 + N列 ==========
            for (int r = 1; r < totalRow; r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;
                double sumValue = 0;
                String val10 = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_SUM_VALUE));
                String val14 = ExcelUtils.getCellStringValue(row.getCell(COL_DED_SUM_VALUE));
                if (val10 != null && !val10.trim().isEmpty()) {
                    try { sumValue += Double.parseDouble(val10); } catch (Exception ignored) {}
                }
                if (val14 != null && !val14.trim().isEmpty()) {
                    try { sumValue += Double.parseDouble(val14); } catch (Exception ignored) {}
                }
                row.createCell(COL_SUM).setCellValue(sumValue);
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) { mainWb.write(fos); }
        }
    }

    // ==================== 平账 ====================

    /**
     * 执行平账
     * @param mainFile     主工资表文件
     * @param salarySheet  工资工作表名
     * @param month        所属月份
     * @param round        当前轮次
     * @param outputDir    输出目录
     * @return 处理后的文件路径
     */
    public String balance(MultipartFile mainFile, String salarySheet,
                           String month, int round, String units, String outputDir) throws Exception {
        File mainTemp = saveToTemp(mainFile);
        try {
            File outputFile = new File(outputDir, "yzpz_balanced_" + System.currentTimeMillis() + ".xlsx");
            doBalance(mainTemp, salarySheet, month, round, units, outputFile);
            return outputFile.getAbsolutePath();
        } finally {
            if (mainTemp.exists()) mainTemp.delete();
        }
    }

    private void doBalance(File mainFile, String salarySheet, String month,
                            int round, String units, File outputFile) throws Exception {
        try (FileInputStream fis = new FileInputStream(mainFile);
             Workbook wb = new XSSFWorkbook(fis)) {

            String dataSheetName = roundSheetName(round, 1);

            Sheet dataSheet = wb.getSheet(dataSheetName);
            if (dataSheet == null) throw new Exception(dataSheetName + " 不存在，请先完成处理");

            Sheet salarySheetObj = wb.getSheet(salarySheet);
            if (salarySheetObj == null) throw new Exception("未找到工作表: " + salarySheet);

            Row salaryHeader = getHeaderRow(salarySheetObj);

            // 查找工资表列
            int nameCol = findColumnByHeader(salaryHeader, "姓名");
            int advanceCol = findColumnContains(salaryHeader, "预支数额");
            int balanceCol = findColumnContains(salaryHeader, "平账额");
            int settledCol = findColumnContains(salaryHeader, "是否平账");
            if (settledCol < 0) settledCol = findColumnContains(salaryHeader, "工资表是否平账");
            int dateCol = findColumnContains(salaryHeader, "工资表所属月份");
            if (dateCol < 0) dateCol = findColumnContains(salaryHeader, "工资表月份");

            if (nameCol < 0) throw new Exception("未找到'姓名'列");
            if (advanceCol < 0) throw new Exception("未找到'预支数额'列");
            if (balanceCol < 0) throw new Exception("未找到'平账额'列");

            int unitCol = findColumnByHeader(salaryHeader, "单位");

            // 读取 AZ1 标记
            String typeMarker = ExcelUtils.getCellStringValue(dataSheet.getRow(0).getCell(COL_TYPE_MARKER));
            boolean useMatching = "需扣回".equals(typeMarker);

            // 读取第12列（SUM）、第13列（姓名）、第14列（扣款总额）
            Map<String, Double> splitData = new HashMap<>();
            for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                Row row = dataSheet.getRow(r);
                if (row == null) continue;
                String name13 = ExcelUtils.getCellStringValue(row.getCell(COL_DED_SUM_NAME));
                String sumVal = ExcelUtils.getCellStringValue(row.getCell(COL_SUM));
                String dedVal = ExcelUtils.getCellStringValue(row.getCell(COL_DED_SUM_VALUE));
                if (name13 != null && !name13.trim().isEmpty() && !"总计".equals(name13.trim())
                        && sumVal != null && !sumVal.isEmpty()
                        && dedVal != null && !dedVal.isEmpty()) {
                    double sumNum = parseDouble(sumVal);
                    if (sumNum > 0) {
                        double dedNum = parseDouble(dedVal);
                        splitData.put(name13.trim(), Math.abs(dedNum));
                    }
                }
            }

            // 读取数据表姓名+预支数额
            Map<String, List<double[]>> sheet4Data = new LinkedHashMap<>();
            if (useMatching) {
                Set<String> dedNames = new HashSet<>();
                for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                    Row row = dataSheet.getRow(r);
                    if (row == null) continue;
                    String n = ExcelUtils.getCellStringValue(row.getCell(COL_DED_NAME));
                    if (n != null && !n.trim().isEmpty() && !"总计".equals(n.trim())) {
                        dedNames.add(n.trim());
                    }
                }
                for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                    Row row = dataSheet.getRow(r);
                    if (row == null) continue;
                    String n = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_NAME));
                    String v = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_VALUE));
                    if (n != null && dedNames.contains(n.trim()) && v != null && !v.trim().isEmpty()) {
                        sheet4Data.computeIfAbsent(n.trim(), k -> new ArrayList<>())
                                .add(new double[]{parseDouble(v), (double) r, 0.0});
                    }
                }
            } else {
                for (int r = 1; r <= dataSheet.getLastRowNum(); r++) {
                    Row row = dataSheet.getRow(r);
                    if (row == null) continue;
                    String n = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_NAME));
                    String v = ExcelUtils.getCellStringValue(row.getCell(COL_ADVANCE_VALUE));
                    if (n != null && !n.trim().isEmpty() && !"总计".equals(n.trim()) && v != null && !v.trim().isEmpty()) {
                        sheet4Data.computeIfAbsent(n.trim(), k -> new ArrayList<>())
                                .add(new double[]{parseDouble(v), (double) r, 0.0});
                    }
                }
            }

            String today = LocalDate.now().format(DATE_FMT);
            int updatedCount = 0;
            List<Integer> updatedRows = new ArrayList<>();
            Set<Integer> noBalanceRows = new HashSet<>();

            // 扫描工资表（平账额为空的行）
            List<SalaryRecord> salaryRecs = new ArrayList<>();
            for (int r = 2; r <= salarySheetObj.getLastRowNum(); r++) {
                Row row = salarySheetObj.getRow(r);
                if (row == null) continue;
                String bal = ExcelUtils.getCellStringValue(row.getCell(balanceCol));
                if (bal == null || bal.trim().isEmpty()) {
                    // 单位筛选
                    if (!matchesUnitFilter(row, unitCol, units)) continue;
                    salaryRecs.add(new SalaryRecord(r,
                            ExcelUtils.getCellStringValue(row.getCell(nameCol)),
                            ExcelUtils.getCellStringValue(row.getCell(advanceCol))));
                }
            }

            // 第一步：处理拆分数据（与Python逻辑一致）
            // splitData: 姓名 -> 需要平账的金额（第14列绝对值）
            // remaining: 需要平账的剩余额度（逐行扣减）
            // 当 origAdv <= remaining：该行全部平账，remaining -= origAdv
            // 当 origAdv > remaining：拆分，原行=remaining(平账)，新行=origAdv-remaining(不平账)
            // 当 remaining <= 0：额度已用完，后续行不参与平账
            List<Object[]> newRowsInfo = new ArrayList<>(); // [源行号, 不平账部分金额]

            for (Map.Entry<String, Double> entry : splitData.entrySet()) {
                String nameStr = entry.getKey();
                double noBalanceAmt = entry.getValue();  // 不需要平账的金额（第14列绝对值）
                if (!sheet4Data.containsKey(nameStr)) continue;

                List<SalaryRecord> matchingRecs = salaryRecs.stream()
                        .filter(r -> nameStr.equals(r.name)).collect(Collectors.toList());
                if (matchingRecs.isEmpty()) continue;

                double totalAdv = matchingRecs.stream()
                        .mapToDouble(r -> parseDouble(r.advance)).sum();
                if (totalAdv <= 0) continue;

                // remaining = 不需要平账的金额（与Python的remaining_balance一致）
                double remaining = noBalanceAmt;

                for (SalaryRecord rec : matchingRecs) {
                    if (remaining <= 0) {
                        // 与Python一致：不需要平账的额度已用完，该行不参与平账
                        noBalanceRows.add(rec.rowIndex);
                        continue;
                    }

                    double origAdv = parseDouble(rec.advance);
                    if (origAdv <= 0) continue;

                    if (origAdv <= remaining) {
                        // 该行全部参与平账，从 remaining 扣减
                        // 与Python一致：不加入 noBalanceRows，后续步骤2正常平账
                        remaining -= origAdv;
                    } else {
                        // 该行只能部分平账
                        // 与Python一致：balance_part = remaining（原行保留平账部分）
                        //               no_balance_part = origAdv - remaining（新行保留不平账部分）
                        double balancePart = remaining;            // 平账部分（保留在原行）
                        double noBalancePart = origAdv - remaining; // 不平账部分（追加到新行）
                        Row row = salarySheetObj.getRow(rec.rowIndex);
                        if (row != null) {
                            Cell advCell = row.getCell(advanceCol);
                            if (advCell == null) advCell = row.createCell(advanceCol);
                            advCell.setCellValue(balancePart); // 原行保留平账部分
                        }
                        // 收集新增行信息：不平账部分的金额
                        newRowsInfo.add(new Object[]{rec.rowIndex, noBalancePart});
                        remaining = 0;
                    }
                }
            }

            // 批量新增行（不平账部分，与Python一致：追加到末尾，不加入salaryRecs）
            // 注意：不能直接用 getLastRowNum()，因为可能包含空白"幽灵行"
            // 需要从底部向上扫描找到真正有数据的最后一行
            int actualLastRow = salarySheetObj.getLastRowNum();
            for (int checkRow = salarySheetObj.getLastRowNum(); checkRow >= 0; checkRow--) {
                Row row = salarySheetObj.getRow(checkRow);
                if (row == null) continue;
                boolean rowHasData = false;
                // 只检查前20列，覆盖主要数据区
                for (int c = 0; c < Math.min(row.getLastCellNum(), 20); c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null) {
                        String val = ExcelUtils.getCellStringValue(cell);
                        if (val != null && !val.trim().isEmpty()) {
                            rowHasData = true;
                            break;
                        }
                    }
                }
                if (rowHasData) {
                    actualLastRow = checkRow;
                    break;
                }
            }

            int appendIdx = 0;
            for (Object[] info : newRowsInfo) {
                int srcRowIdx = (Integer) info[0];
                double noBalancePart = (Double) info[1];
                Row srcRow = salarySheetObj.getRow(srcRowIdx);
                if (srcRow == null) continue;
                int newRowIdx = actualLastRow + (++appendIdx);
                Row newRow = salarySheetObj.createRow(newRowIdx);
                // 复制所有单元格
                for (int c = 0; c < srcRow.getLastCellNum(); c++) {
                    Cell srcCell = srcRow.getCell(c);
                    if (srcCell != null) {
                        Cell destCell = newRow.createCell(c);
                        ExcelUtils.copyCellValue(srcCell, destCell);
                    }
                }
                // 清空平账相关列
                if (newRow.getCell(balanceCol) != null) newRow.getCell(balanceCol).setBlank();
                if (settledCol >= 0 && newRow.getCell(settledCol) != null) newRow.getCell(settledCol).setBlank();
                if (dateCol >= 0 && newRow.getCell(dateCol) != null) newRow.getCell(dateCol).setBlank();
                // 设置预支数额为不平账部分
                Cell advanceCell = newRow.getCell(advanceCol);
                if (advanceCell == null) advanceCell = newRow.createCell(advanceCol);
                advanceCell.setCellValue(noBalancePart);
                // ✅ 与Python一致：不添加到 salaryRecs，避免错误平账
            }

            // 第二步：正常平账
            for (SalaryRecord rec : salaryRecs) {
                if (noBalanceRows.contains(rec.rowIndex)) continue;
                if (rec.name == null) continue;

                String nameTrim = rec.name.trim();
                List<double[]> records = sheet4Data.get(nameTrim);
                if (records == null) continue;

                for (double[] record : records) {
                    if (record[2] > 0) continue; // used flag ≈ 已使用
                    double salaryAdv = parseDouble(rec.advance);
                    double sheet4Adv = record[0];
                    if (Math.abs(salaryAdv - sheet4Adv) < 0.01) {
                        record[2] = 1; // mark as used
                        Row row = salarySheetObj.getRow(rec.rowIndex);
                        // 与Python一致：重新读取预支数额单元格的当前值（拆分后可能已被修改）
                        Cell advCell = row.getCell(advanceCol);
                        String realAdvance = ExcelUtils.getCellStringValue(advCell);
                        Cell balCell = row.getCell(balanceCol);
                        if (balCell == null) balCell = row.createCell(balanceCol);
                        balCell.setCellValue(realAdvance);
                        if (settledCol >= 0) {
                            Cell setCell = row.getCell(settledCol);
                            if (setCell == null) setCell = row.createCell(settledCol);
                            setCell.setCellValue(today);
                        }
                        if (dateCol >= 0) {
                            Cell dateCell = row.getCell(dateCol);
                            if (dateCell == null) dateCell = row.createCell(dateCol);
                            dateCell.setCellValue(month);
                        }
                        updatedCount++;
                        updatedRows.add(rec.rowIndex);
                        break;
                    }
                }
            }

            // 标红：保留原字体名称/大小，只修改颜色
            if (!updatedRows.isEmpty()) {
                for (int rowNum : updatedRows) {
                    Row row = salarySheetObj.getRow(rowNum);
                    if (row == null) continue;
                    for (int c = 0; c < row.getLastCellNum(); c++) {
                        Cell cell = row.getCell(c);
                        if (cell == null) continue;
                        CellStyle oldStyle = cell.getCellStyle();
                        if (oldStyle == null) continue;
                        // 基于原样式创建新样式，只改字体颜色
                        CellStyle newStyle = wb.createCellStyle();
                        newStyle.cloneStyleFrom(oldStyle);
                        Font oldFont = wb.getFontAt(oldStyle.getFontIndexAsInt());
                        Font redFont = wb.createFont();
                        redFont.setFontName(oldFont.getFontName());
                        redFont.setFontHeight(oldFont.getFontHeight());
                        redFont.setBold(oldFont.getBold());
                        redFont.setItalic(oldFont.getItalic());
                        redFont.setUnderline(oldFont.getUnderline());
                        redFont.setColor(IndexedColors.RED.getIndex());
                        newStyle.setFont(redFont);
                        cell.setCellStyle(newStyle);
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // ==================== 文件保存 ====================

    // ==================== 检测方法 ====================

    /**
     * 检测扣款文件中的可用列类型（需扣回、预支扣款、预支扣回）
     * @param dedFile 扣款文件
     * @param dedSheet 扣款文件工作表名
     * @return 可用的列类型列表
     */
    public List<String> detectDeductionColumns(MultipartFile dedFile, String dedSheet) throws Exception {
        File temp = saveToTemp(dedFile);
        try (FileInputStream fis = new FileInputStream(temp);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheet(dedSheet);
            if (sheet == null) {
                if (wb.getNumberOfSheets() > 0) sheet = wb.getSheetAt(0);
                else throw new Exception("扣款文件没有工作表");
            }

            Row header = getHeaderRow(sheet);

            String[] colNames = {"需扣回", "预支扣款", "预支扣回"};
            List<String> available = new ArrayList<>();

            for (String colName : colNames) {
                int colIdx = findColumnContains(header, colName);
                if (colIdx >= 0) {
                    // 检查该列是否有有效数据（第3行起）
                    boolean hasData = false;
                    for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                        Row row = sheet.getRow(r);
                        if (row == null) continue;
                        Cell cell = row.getCell(colIdx);
                        if (cell == null) continue;
                        if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                            double val = cell.getNumericCellValue();
                            if (val != 0.0) { hasData = true; break; }
                        } else {
                            String v = ExcelUtils.getCellStringValue(cell);
                            if (v != null && !v.trim().isEmpty() && !v.trim().equals("0") && !v.trim().equals("0.0") && !v.trim().equals("0.00")) {
                                hasData = true;
                                break;
                            }
                        }
                    }
                    if (hasData) available.add(colName);
                }
            }
            return available;
        } finally {
            if (temp.exists()) temp.delete();
        }
    }

    /**
     * 检测主文件中工资表单位列的唯一值
     * @param file 主文件
     * @param sheetName 工资表工作表名
     * @return 单位唯一值列表
     */
    public List<String> detectUnits(MultipartFile file, String sheetName) throws Exception {
        File temp = saveToTemp(file);
        try (FileInputStream fis = new FileInputStream(temp);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) throw new Exception("未找到工作表: " + sheetName);

            Row header = getHeaderRow(sheet);

            int unitCol = findColumnByHeader(header, "单位");
            if (unitCol < 0) return new ArrayList<>(); // 没有单位列

            Set<String> units = new TreeSet<>();
            for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row != null) {
                    String v = ExcelUtils.getCellStringValue(row.getCell(unitCol));
                    if (v != null && !v.trim().isEmpty()) {
                        units.add(v.trim());
                    }
                }
            }
            return new ArrayList<>(units);
        } finally {
            if (temp.exists()) temp.delete();
        }
    }

    /**
     * 保存工作簿（另存为）
     */
    public String saveAs(MultipartFile file, String outputDir) throws Exception {
        File temp = saveToTemp(file);
        try {
            File outputFile = new File(outputDir, "yzpz_saved_" + System.currentTimeMillis() + ".xlsx");
            try (FileInputStream fis = new FileInputStream(temp);
                 Workbook wb = new XSSFWorkbook(fis);
                 FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
            return outputFile.getAbsolutePath();
        } finally {
            if (temp.exists()) temp.delete();
        }
    }

    // ==================== 工具方法 ====================

    static class SalaryRecord {
        int rowIndex;
        String name;
        String advance;
        SalaryRecord(int r, String n, String a) { rowIndex = r; name = n; advance = a; }
    }

    private File saveToTemp(MultipartFile file) throws IOException {
        File tmp = File.createTempFile("yzpz_", ".xlsx");
        file.transferTo(tmp.toPath());
        return tmp;
    }

    private int findColumnByHeader(Row headerRow, String name) {
        if (headerRow == null) return -1;
        for (int c = 0; c < headerRow.getLastCellNum(); c++) {
            if (name.equals(ExcelUtils.getCellStringValue(headerRow.getCell(c)))) return c;
        }
        return -1;
    }

    private int findColumnContains(Row headerRow, String keyword) {
        if (headerRow == null) return -1;
        for (int c = 0; c < headerRow.getLastCellNum(); c++) {
            String v = ExcelUtils.getCellStringValue(headerRow.getCell(c));
            if (v != null && v.contains(keyword)) return c;
        }
        return -1;
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row != null ? row : sheet.createRow(rowIndex);
    }

    private double parseDouble(String s) {
        if (s == null || s.isEmpty()) return 0;
        try { return Double.parseDouble(s.replaceAll("[^\\d.\\-]", "")); }
        catch (NumberFormatException e) { return 0; }
    }

    // ==================== 工具方法（消除重复） ====================

    /** 计算轮次对应的 Sheet 名称：names=3+(round-1)*2, data=names+1 */
    private String roundSheetName(int round, int offset) {
        return "Sheet" + (3 + (round - 1) * 2 + offset);
    }

    /** 获取工作表的标题行（优先第2行，兜底第1行） */
    private Row getHeaderRow(Sheet sheet) {
        Row header = sheet.getRow(1);
        return header != null ? header : sheet.getRow(0);
    }

    /** 检查行是否匹配逗号分隔的单位筛选条件 */
    private boolean matchesUnitFilter(Row row, int unitCol, String units) {
        if (units == null || units.isEmpty() || unitCol < 0) return true;
        String unitVal = ExcelUtils.getCellStringValue(row.getCell(unitCol));
        if (unitVal == null || unitVal.trim().isEmpty()) return false;
        for (String u : units.split(",")) {
            if (u.trim().equals(unitVal.trim())) return true;
        }
        return false;
    }
}
