package com.example.project.service.finance.cwgjx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 财务工具箱 - 代工费拆分
 * 代工费表格 + 发放记录表格 处理（cw_dgf_cf_xzy.py + dgf_cf_frame.py）
 */
@Service
public class CwgjxDgfCfService {

    private static final boolean DEBUG = true; // 调试开关，发布时可改为 false

    private void log(String msg) {
        if (DEBUG) System.out.println("[CwgjxDgfCfService] " + msg);
    }

    private static final Set<String> HIGHLIGHT_KEY_WORDS = new HashSet<>(Arrays.asList(
            "公司", "个体工商户", "个体户", "服务部"   // 对齐Python: cw_dgf_cf_xzy.py 第22行
    ));

    private static final double LIMIT = 4995.0;  // 拆分阈值，与Python一致

    private static final Map<String, String> COMPANY_NAME_MAP = new HashMap<>();
    static {
        COMPANY_NAME_MAP.put("刘先锋个人现金卡", "刘先锋现金卡支付");
        COMPANY_NAME_MAP.put("温州鑫锦途供应链服务有限公司", "温州鑫锦途");
        COMPANY_NAME_MAP.put("浙江锦途人力资源有限公司", "浙江锦途");
        COMPANY_NAME_MAP.put("温州锦途服务外包有限公司", "温州锦途外包");
        COMPANY_NAME_MAP.put("衢州鑫途服务外包有限公司", "衢州鑫途外包");
        COMPANY_NAME_MAP.put("浙江锦途人力资源有限公司乐清分公司", "乐清分公司");
        COMPANY_NAME_MAP.put("芜湖才库人力资源有限公司", "芜湖才库");
        COMPANY_NAME_MAP.put("衢州驰锦企业服务有限公司", "衢州驰锦");
        COMPANY_NAME_MAP.put("安徽锦途企业服务集团有限公司", "安徽锦途服务");
        COMPANY_NAME_MAP.put("安徽锦途企业管理咨询有限公司", "安徽锦途");
        COMPANY_NAME_MAP.put("安徽锦途人力资源有限公司", "安徽人力");
        COMPANY_NAME_MAP.put("安庆市同益劳务服务有限公司", "安庆同益");
        COMPANY_NAME_MAP.put("台州众通人才服务有限公司", "台州众通");
        COMPANY_NAME_MAP.put("台州锦途人力资源服务外包有限公司", "台州锦途外包");
        COMPANY_NAME_MAP.put("温州鑫途企业服务有限公司", "温州鑫途");
        COMPANY_NAME_MAP.put("温州驰锦企业服务有限公司", "温州驰锦");
        COMPANY_NAME_MAP.put("浙江盛威安保服务有限公司", "盛威安保");
        COMPANY_NAME_MAP.put("温州市合静人力资源有限公司", "合静");
        COMPANY_NAME_MAP.put("铜陵锦途劳务有限公司", "铜陵锦途");
    }

    /**
     * 批量处理（对齐 Python dgf_cf_frame.py 的多文件 for 循环逻辑）
     * 多个代工费文件串行处理，bState 跨文件累加，共享同一个发放记录表格
     * 注意：file1List 和 file2 已由 Controller 层保存到磁盘，此方法直接读取
     */
    public ProcessResult processBatch(File[] file1List, String[] sheet1Names,
                                       File file2, String sheet2Name, String sheetBName,
                                       String uploadDir) throws Exception {
        log("========== 开始批量处理代工费拆分 ==========");
        log("代工费文件数: " + file1List.length);
        log("发放记录文件: " + file2.getName());

        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        // file2 已由 Controller 预先保存到磁盘，直接使用
        File temp2 = file2;
        log("发放记录文件已由 Controller 保存完成");

        // 读取 B 表初始状态（只读，后续实时更新）
        Workbook wb2 = null;
        Map<String, Double> bState = null;
        Map<String, String> lookupMap = null;
        int totalSummary = 0;
        int totalSplits = 0;
        List<String> allStep4Details = new ArrayList<>();
        List<String> allStep5Details = new ArrayList<>();
        List<String> allStep5Result = new ArrayList<>();
        String lastRemark1 = "";
        String lastCompanyPrefix = "";

        try {
            wb2 = WorkbookFactory.create(new FileInputStream(temp2));
            bState = initBState(wb2, sheetBName);
            lookupMap = buildLookupMap(wb2, "供应商信息表（定稿）");
            log("B表状态读取完成，共 " + bState.size() + " 个收款人");

            Set<String> usedNames = new HashSet<>();

            List<String> summaryPaths = new ArrayList<>();

            // === 串行处理每个代工费文件（对齐 Python 第296行 for 循环） ===
            for (int i = 0; i < file1List.length; i++) {
                File file1 = file1List[i];
                String sheet1Name = sheet1Names[i];

                log("=".repeat(60));
                log("【文件 " + (i+1) + "/" + file1List.length + "】" + file1.getName());
                log("=".repeat(60));

                // file1 已由 Controller 预先保存到磁盘，直接使用
                File temp1 = file1;
                log("代工费文件已由 Controller 保存完成, sheet1=" + sheet1Name);

                try (Workbook wb1 = WorkbookFactory.create(new FileInputStream(temp1))) {
                    // 步骤1：汇总
                    log("【步骤1】汇总'介绍人'和'金额'...");
                    List<SummaryRecord> summary = processFirstExcel(wb1, sheet1Name);
                    log("  汇总完成，共 " + summary.size() + " 条记录");
                    totalSummary += summary.size();

                    // 保存汇总后的代工费表格
                    String origName = stripUuidPrefix(file1.getName());
                    String baseName = origName.contains(".") ? origName.substring(0, origName.lastIndexOf('.')) : origName;
                    String summaryPath = uploadDir + File.separator + baseName + "_已汇总.xlsx";
                    try (FileOutputStream fos = new FileOutputStream(summaryPath)) { wb1.write(fos); }
                    summaryPaths.add(summaryPath);
                    log("  代工费表格已保存: " + summaryPath);

                    // 步骤2：公司前缀和备注1
                    log("【步骤2】读取公司前缀并解析备注1...");
                    lastCompanyPrefix = getCompanyPrefix(wb1);
                    lastRemark1 = parseRemark1FromFilename(file1.getName());
                    log("  公司前缀：" + lastCompanyPrefix);
                    log("  解析备注1：" + lastRemark1);

                    // 步骤3：构建追加记录
                    List<AppendedRecord> splitRecords = new ArrayList<>();
                    for (SummaryRecord sr : summary) {
                        String receiver = lookupMap.get(sr.name);
                        splitRecords.add(new AppendedRecord(sr.name, sr.amount, receiver));
                        usedNames.add(sr.name);
                    }
                    log("  已构建追加记录，共 " + splitRecords.size() + " 条");

                    // 步骤4：预览拆分（bState 实时累加）
                    log("【步骤4】逐条判断追加记录并拆分...");
                    Map<String, Double> previewState = new HashMap<>(bState);
                    for (AppendedRecord rec : splitRecords) {
                        String receiver = rec.receiver;
                        double before = receiver != null ? previewState.getOrDefault(receiver, 0.0) : 0.0;
                        double after = before + rec.amount;
                        String detail = "    需拆分: " + rec.name + " (" + receiver + ") 累计 " + before + " + " + rec.amount + " = " + after;
                        log(detail);
                        allStep4Details.add(detail);
                        if (receiver != null) previewState.put(receiver, after);
                    }
                    log("  判断完成，需拆分 " + splitRecords.size() + " 条。");

                    // 步骤5：回写拆分行（bState 实实更新并跨文件累加）
                    log("【步骤5】回写拆分行到工作簿A...");
                    SplitResult splitResult = writeSplitRows(wb2, sheet2Name, splitRecords,
                            lookupMap, bState, usedNames);
                    totalSplits += splitResult.records.size();
                    allStep5Details.addAll(splitResult.step5Details);
                    allStep5Result.addAll(splitResult.step5Result);
                    log("  拆分结果（本文件）：");
                    for (SplitRecord rec : splitResult.records) {
                        String receiverDisplay = rec.receiver != null ? rec.receiver : "";
                        log("    " + rec.name + "   " + rec.amount + "   " + receiverDisplay);
                    }
                    log("  回写完成，新增拆分行");
                    log("  发放记录表格（工作簿A）已保存");

                    // 步骤6：填充备注列
                    log("【步骤6】为拆分记录填充备注列...");
                    fillSplitRows(wb2, sheet2Name, splitResult.splitStartRow,
                            splitResult.records, lastRemark1, lastCompanyPrefix);
                    log("  填充完成，工作簿A已保存");
                    log("");

                } finally {
                    // file1 已由 Controller 管理生命周期，Service 不负责删除
                    log("代工费文件处理完毕: " + file1.getName());
                }
            }

            // 保存最终文件（用原文件名）
            String origFile2Name = stripUuidPrefix(file2.getName());
            String baseFile2Name = origFile2Name.contains(".") ? origFile2Name.substring(0, origFile2Name.lastIndexOf('.')) : origFile2Name;
            String outputPath2 = uploadDir + File.separator + baseFile2Name + "_拆分结果.xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputPath2)) { wb2.write(fos); }
            log("发放记录表已保存: " + outputPath2);

            ProcessResult result = new ProcessResult();
            result.outputPath = outputPath2;
            result.summaryPaths = summaryPaths;
            result.summaryPath = summaryPaths.isEmpty() ? null : summaryPaths.get(summaryPaths.size() - 1);
            result.summaryCount = totalSummary;
            result.splitCount = totalSplits;
            result.remark1 = lastRemark1;
            result.companyPrefix = lastCompanyPrefix;
            result.step4Details = allStep4Details;
            result.step5Details = allStep5Details;
            result.step5Result = allStep5Result;
            log("========== 批量处理完毕 ==========");
            return result;

        } finally {
            if (wb2 != null) wb2.close();
            // temp2 由 Controller 管理生命周期，Service 不负责删除
        }
    }

    public ProcessResult process(MultipartFile file1, String sheet1Name,
                                 MultipartFile file2, String sheet2Name, String sheetBName,
                                 String uploadDir) throws Exception {
        log("========== 开始处理代工费拆分 ==========");
        log("file1=" + file1.getOriginalFilename() + ", sheet1Name=" + sheet1Name);
        log("file2=" + file2.getOriginalFilename() + ", sheet2Name=" + sheet2Name + ", sheetBName=" + sheetBName);

        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();
        File temp1 = new File(dir, UUID.randomUUID() + "_" + file1.getOriginalFilename());
        File temp2 = new File(dir, UUID.randomUUID() + "_" + file2.getOriginalFilename());
        log("保存临时文件: " + temp1.getName() + ", " + temp2.getName());
        file1.transferTo(temp1);
        file2.transferTo(temp2);
        log("临时文件保存完成，file1大小=" + temp1.length() + ", file2大小=" + temp2.length());

        try (Workbook wb1 = WorkbookFactory.create(new FileInputStream(temp1));
             Workbook wb2 = WorkbookFactory.create(new FileInputStream(temp2))) {
            log("Workbook 加载完成: wb1=" + wb1.getSheetName(0) + ", wb2=" + wb2.getSheetName(0));

            // 步骤1：汇总表1（介绍人 + 金额）
            log("【步骤1】汇总'介绍人'和'金额'...");
            List<SummaryRecord> summary = processFirstExcel(wb1, sheet1Name);
            log("  汇总完成，共 " + summary.size() + " 条记录：");
            for (SummaryRecord r : summary) log("    - " + r.name + ": " + r.amount);
            log("  代工费表格已保存（汇总结果已写入原文件下方）");
            log("");

            // 步骤2：读取公司前缀并解析备注1
            log("【步骤2】读取公司前缀并解析备注1...");
            String companyPrefix = getCompanyPrefix(wb1);
            log("  公司前缀：" + companyPrefix);
            String remark1 = parseRemark1FromFilename(file1.getOriginalFilename());
            log("  解析备注1：" + remark1);
            log("");

            // 步骤3：选择发放记录表格的汇总工作簿（读取B状态 + 构建追加记录）
            log("【步骤3】选择发放记录表格的汇总工作簿...");
            log("  工作簿（B）：" + sheetBName);
            Map<String, Double> bState = initBState(wb2, sheetBName);
            Map<String, String> lookupMap = buildLookupMap(wb2, "供应商信息表（定稿）");
            List<AppendedRecord> splitRecords = new ArrayList<>();
            Set<String> usedNames = new HashSet<>();
            for (SummaryRecord sr : summary) {
                String receiver = lookupMap.get(sr.name);
                splitRecords.add(new AppendedRecord(sr.name, sr.amount, receiver));
                usedNames.add(sr.name);
            }
            log("  B表状态读取完成");
            log("  已构建追加记录，共 " + splitRecords.size() + " 条");
            log("");

            // 步骤4：逐条判断追加记录并拆分
            log("【步骤4】逐条判断追加记录并拆分...");
            List<String> step4Details = new ArrayList<>();
            // 预览（不修改bState）
            for (AppendedRecord rec : splitRecords) {
                String receiver = rec.receiver;
                double before = receiver != null ? bState.getOrDefault(receiver, 0.0) : 0.0;
                double after = before + rec.amount;
                String detail = "    需拆分: " + rec.name + " (" + receiver + ") 累计 " + before + " + " + rec.amount + " = " + after;
                log(detail);
                step4Details.add(detail);
            }
            log("  判断完成，需拆分 " + splitRecords.size() + " 条。");
            log("  工作簿B仅用于参考判断，未修改文件");
            log("");

            // 执行实际拆分
            log("【步骤5】回写拆分行到工作簿A...");
            SplitResult splitResult = writeSplitRows(wb2, sheet2Name, splitRecords, lookupMap, bState, usedNames);
            List<SplitRecord> newRecords = splitResult.records;
            int splitStartRow = splitResult.splitStartRow;
            log("  拆分结果：");
            splitResult.step5Result.add("  拆分结果：");
            for (SplitRecord rec : newRecords) {
                String receiverDisplay = rec.receiver != null ? rec.receiver : "";
                String resultLine = "    " + rec.name + "   " + rec.amount + "   " + receiverDisplay;
                log(resultLine);
                splitResult.step5Result.add(resultLine);
            }
            log("  回写完成，新增拆分行");
            log("  发放记录表格（工作簿A）已保存");
            log("");

            // 步骤6：为拆分记录填充备注列
            log("【步骤6】为拆分记录填充备注列...");
            fillSplitRows(wb2, sheet2Name, splitStartRow, newRecords, remark1, companyPrefix);
            log("  填充完成，工作簿A已保存");
            log("");

            log("全部处理完毕！");

            // 保存文件（用原始文件名）
            String name1 = file1.getOriginalFilename();
            String base1 = name1.contains(".") ? name1.substring(0, name1.lastIndexOf('.')) : name1;
            String outputPath1 = uploadDir + File.separator + base1 + "_已汇总.xlsx";
            try (FileOutputStream fos1 = new FileOutputStream(outputPath1)) { wb1.write(fos1); }
            log("代工费表格已保存: " + outputPath1);

            String name2 = file2.getOriginalFilename();
            String base2 = name2.contains(".") ? name2.substring(0, name2.lastIndexOf('.')) : name2;
            String outputPath2 = uploadDir + File.separator + base2 + "_拆分结果.xlsx";
            try (FileOutputStream fos2 = new FileOutputStream(outputPath2)) { wb2.write(fos2); }
            log("发放记录表已保存: " + outputPath2);

            // 构造返回结果
            ProcessResult result = new ProcessResult();
            result.outputPath = outputPath2;
            result.summaryPath = outputPath1;
            result.summaryCount = summary.size();
            result.splitCount = newRecords.size();
            result.remark1 = remark1;
            result.companyPrefix = companyPrefix;
            result.step4Details = step4Details;
            result.step5Details = splitResult.step5Details;
            result.step5Result = splitResult.step5Result;
            log("========== 处理完毕 ==========");
            return result;
        } catch (Throwable t) {
            log("处理失败: " + t.getClass().getSimpleName() + " - " + t.getMessage());
            t.printStackTrace();
            throw t;
        } finally {
            temp1.delete(); temp2.delete();
            log("临时文件已删除");
        }
    }

    private List<SummaryRecord> processFirstExcel(Workbook wb, String sheetName) throws Exception {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) throw new Exception("工作表不存在: " + sheetName);

        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        log("  公式求值器已创建");

        HeaderInfo headerInfo = findHeaderCols(sheet, Arrays.asList("介绍人", "金额"));
        if (headerInfo.headerRow < 0) throw new Exception("未找到表头: 介绍人/金额");
        int headerRow = headerInfo.headerRow;
        int jieshaoCol = headerInfo.colIndices[0], hejiCol = headerInfo.colIndices[1];

        // 用 data_only 模式读取公式计算结果（POI 默认读取公式计算结果，等价于 data_only=True）
        Map<String, Double> summary = new LinkedHashMap<>();
        log("  表头行: " + headerRow + ", 介绍人列: " + jieshaoCol + ", 金额列: " + hejiCol + ", 总行数: " + (sheet.getLastRowNum() + 1));
        // 调试：打印表头行所有非空单元格
        Row hdrRow = sheet.getRow(headerRow);
        if (hdrRow != null) {
            StringBuilder hdrCells = new StringBuilder("  表头内容: ");
            for (Cell c : hdrRow) {
                if (c != null) hdrCells.append("col").append(c.getColumnIndex()).append("=[").append(getStr(c)).append("] ");
            }
            log(hdrCells.toString().trim());
        }
        int scannedRows = 0, emptyName = 0, numName = 0, nullAmt = 0;
        for (int r = headerRow + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            scannedRows++;

            // 对公式单元格尝试逐个求值（单个公式失败不影响其他单元格）
            Cell nameCell = row.getCell(jieshaoCol);
            Cell amtCell = row.getCell(hejiCol);
            tryEvaluateCell(nameCell, evaluator);
            tryEvaluateCell(amtCell, evaluator);

            String name = getStr(nameCell);
            Double amount = parseDbl(amtCell);
            // 调试：打印前3行的原始值（含单元格类型）
            if (scannedRows <= 3) {
                String nType = nameCell == null ? "null" : nameCell.getCellType().toString();
                String aType = amtCell == null ? "null" : amtCell.getCellType().toString();
                log("    行" + r + " name='" + name + "'(" + nType + ") amount=" + amount + "(" + aType + ")");
            }
            if (name.isEmpty()) { emptyName++; continue; }
            if (name.matches("\\d*\\.?\\d+")) { numName++; continue; }
            if (amount == null) { nullAmt++; continue; }
            summary.merge(name, amount, Double::sum);
        }
        if (scannedRows > 0) log("  扫描 " + scannedRows + " 行, 过滤: 空名" + emptyName + " 数字名" + numName + " 空金额" + nullAmt);

        // 在原 sheet 末尾空 5 行后写回汇总结果（等价于 Python process_first_excel 的写回逻辑）
        int lastDataRow = findLastDataRow(sheet, headerRow, Arrays.asList(jieshaoCol, hejiCol));
        int writeStart = lastDataRow + 5 + 1;  // 空 5 行，再下一行开始写

        // 写汇总表头
        Row headerRowCell = sheet.getRow(writeStart) != null ? sheet.getRow(writeStart) : sheet.createRow(writeStart);
        headerRowCell.createCell(0).setCellValue("介绍人");
        headerRowCell.createCell(1).setCellValue("求和项：金额");

        // 写汇总数据
        int dataRowIdx = writeStart + 1;
        for (Map.Entry<String, Double> e : summary.entrySet()) {
            Row dataRow = sheet.getRow(dataRowIdx) != null ? sheet.getRow(dataRowIdx) : sheet.createRow(dataRowIdx);
            dataRow.createCell(0).setCellValue(e.getKey());
            dataRow.createCell(1).setCellValue(e.getValue());
            dataRowIdx++;
        }

        // 构造返回结果
        List<SummaryRecord> result = new ArrayList<>();
        for (Map.Entry<String, Double> e : summary.entrySet()) result.add(new SummaryRecord(e.getKey(), e.getValue()));
        return result;
    }

    private String getCompanyPrefix(Workbook wb) {
        Sheet sheet = wb.getSheet("个税统计");
        if (sheet == null) return "";
        HeaderInfo info1 = findHeaderCols(sheet, Arrays.asList("发薪单位"));
        HeaderInfo info2 = null;
        if (info1.headerRow < 0) {
            info2 = findHeaderCols(sheet, Arrays.asList("发薪公司"));
            if (info2.headerRow < 0) return "";
        }
        HeaderInfo info = info1.headerRow >= 0 ? info1 : info2;
        int col = info.colIndices[0];
        int headerRow = info.headerRow;
        for (int r = headerRow + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String v = getStr(row.getCell(col));
            if (!v.isEmpty()) return COMPANY_NAME_MAP.getOrDefault(v, v);
        }
        return "";
    }

    private Map<String, Double> initBState(Workbook wb, String sheetName) {
        Map<String, Double> state = new HashMap<>();
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) return state;
        HeaderInfo headerInfo = findHeaderCols(sheet, Arrays.asList("收款姓名", "求和项：打款金额"));
        if (headerInfo.headerRow < 0) return state;
        int headerRow = headerInfo.headerRow;
        int receiverCol = headerInfo.colIndices[0], amountCol = headerInfo.colIndices[1];
        for (int r = headerRow + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String name = getStr(row.getCell(receiverCol));
            Double amount = parseDbl(row.getCell(amountCol));
            if (!name.isEmpty() && amount != null) state.put(name, amount);
        }
        return state;
    }

    private SplitResult writeSplitRows(Workbook wb, String sheetName, List<AppendedRecord> splitRecords,
                                          Map<String, String> lookupMap, Map<String, Double> bState, Set<String> usedNames) throws Exception {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) throw new Exception("工作表不存在: " + sheetName);
        HeaderInfo headerInfo = findHeaderCols(sheet, Arrays.asList("供应商名称", "打款金额", "银行卡号", "收款姓名", "身份证号", "电话号码"));
        if (headerInfo.headerRow < 0) throw new Exception("未找到必要表头");
        int headerRow = headerInfo.headerRow;
        int supplierCol = headerInfo.colIndices[0], dkC = headerInfo.colIndices[1];
        int bankCardCol = headerInfo.colIndices[2], receiverCol = headerInfo.colIndices[3];
        int idCardCol = headerInfo.colIndices[4], phoneCol = headerInfo.colIndices[5];
        int dkH = -1;
        for (int c = 0; c <= sheet.getRow(headerRow).getLastCellNum(); c++) {
            if (c == dkC) continue;
            if ("打款金额".equals(getStr(sheet.getRow(headerRow).getCell(c)))) { dkH = c; break; }
        }
        if (dkH < 0) dkH = dkC;

        int lastRow = findLastDataRow(sheet, headerRow, Arrays.asList(supplierCol, dkC));
        int splitStartRow = lastRow + 1;  // 等价于 Python 的 last_row + 1
        String bLetter = CellReference.convertNumToColString(supplierCol);
        Set<String> used = new HashSet<>(usedNames);
        List<SplitRecord> result = new ArrayList<>();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        SplitResult splitResult = new SplitResult(result, splitStartRow);  // 提前创建，用于收集日志

        // 获取有效的供应商名称集合
        Set<String> validSupplierNames = new HashSet<>(lookupMap.keySet());

        for (AppendedRecord rec : splitRecords) {
            double remaining = rec.amount;  // 剩余待分配金额
            boolean firstSplit = true;

            while (remaining > 0) {
                String candidate = null;
                String chosenReceiver = null;
                double allocateAmount = 0;

                // 原名称（仅第一次尝试）
                if (firstSplit) {
                    if (shouldHighlightStopSplit(rec.receiver)) {
                        // 原收款人含关键词（公司/个体户/刘先锋），直接整条给原名称
                        candidate = rec.name;
                        chosenReceiver = rec.receiver;
                        allocateAmount = remaining;
                        remaining = 0;
                    } else {
                        // 尝试原名称能塞多少
                        double origCurrent = bState.getOrDefault(rec.receiver, 0.0);
                        if (origCurrent < LIMIT) {
                            double canAlloc = Math.min(LIMIT - origCurrent, remaining);
                            if (canAlloc > 0) {
                                candidate = rec.name;
                                chosenReceiver = rec.receiver;
                                allocateAmount = canAlloc;
                                remaining -= canAlloc;
                                // 立即更新bState
                                bState.put(rec.receiver, origCurrent + canAlloc);
                            }
                        }
                    }

                    firstSplit = false;

                    if (candidate != null) {
                        // 写入行
                        lastRow = writeSplitRow(sheet, lastRow, supplierCol, dkC, dkH, bankCardCol, receiverCol, idCardCol, phoneCol, bLetter, candidate, allocateAmount, chosenReceiver, evaluator);
                        result.add(new SplitRecord(candidate, allocateAmount, chosenReceiver));
                        used.add(candidate);
                        String detail = "    拆分写入: " + rec.name + " -> " + candidate + " | 金额=" + allocateAmount + " | 收款人=" + chosenReceiver;
                        log(detail);
                        splitResult.step5Details.add(detail);

                        if (remaining <= 0) break;
                        // 继续找后缀
                    }
                }

                // 后缀1~20
                boolean foundSuffix = false;
                for (int suffix = 1; suffix <= 20; suffix++) {
                    String testName = rec.name + suffix;
                    if (used.contains(testName)) continue;
                    if (!validSupplierNames.contains(testName)) continue;

                    String newReceiver = lookupMap.get(testName);

                    // 含关键词：停止拆分，剩余全部给当前后缀
                    if (shouldHighlightStopSplit(newReceiver)) {
                        candidate = testName;
                        chosenReceiver = newReceiver;
                        allocateAmount = remaining;
                        remaining = 0;
                        foundSuffix = true;
                        break;
                    }

                    // 检查能否分配
                    double currentTotal = bState.getOrDefault(newReceiver, 0.0);
                    if (currentTotal >= LIMIT) {
                        continue;  // 已满，跳过
                    }

                    double canAlloc = Math.min(LIMIT - currentTotal, remaining);
                    if (canAlloc > 0) {
                        candidate = testName;
                        chosenReceiver = newReceiver;
                        allocateAmount = canAlloc;
                        remaining -= canAlloc;
                        // 立即更新bState
                        bState.put(newReceiver, currentTotal + canAlloc);
                        foundSuffix = true;
                        break;
                    }
                }

                if (candidate == null) {
                    // 后缀1~20都失败，回退到原名称（累加已分配+剩余）
                    candidate = rec.name;
                    chosenReceiver = rec.receiver;
                    allocateAmount = remaining;
                    remaining = 0;
                }

                // 写入行
                lastRow = writeSplitRow(sheet, lastRow, supplierCol, dkC, dkH, bankCardCol, receiverCol, idCardCol, phoneCol, bLetter, candidate, allocateAmount, chosenReceiver, evaluator);
                result.add(new SplitRecord(candidate, allocateAmount, chosenReceiver));
                used.add(candidate);
                String detail2 = "    拆分写入: " + rec.name + " -> " + candidate + " | 金额=" + allocateAmount + " | 收款人=" + chosenReceiver;
                log(detail2);
                splitResult.step5Details.add(detail2);
            }
        }
        return splitResult;
    }

    /**
     * 写入单行拆分数据（辅助方法）
     */
    private int writeSplitRow(Sheet sheet, int lastRow, int supplierCol, int dkC, int dkH,
                              int bankCardCol, int receiverCol, int idCardCol, int phoneCol,
                              String bLetter, String candidate, double amount,
                              String chosenReceiver, FormulaEvaluator evaluator) {
        lastRow++;
        // 找到第一个可用行：B列和C列没有真实数据，复用VLOOKUP公式空行
        while (lastRow <= sheet.getLastRowNum()) {
            Row existingRow = sheet.getRow(lastRow);
            if (existingRow == null) break;  // 行不存在，直接创建新行

            boolean hasRealData = false;
            for (int c : Arrays.asList(supplierCol, dkC)) {
                Cell cell = existingRow.getCell(c);
                if (cell == null) continue;
                if (cell.getCellType() == CellType.BLANK) continue;
                if (cell.getCellType() == CellType.FORMULA &&
                    cell.getCachedFormulaResultType() == CellType.ERROR) continue;
                if (cell.getCellType() == CellType.STRING) {
                    String val = cell.getStringCellValue().trim();
                    if (val.isEmpty() || "#N/A".equals(val)) continue;
                }
                hasRealData = true;
                break;
            }

            if (!hasRealData) break;  // B列和C列都没有真实数据，复用该行
            lastRow++;  // 有真实数据，继续找下一行
        }

        Row row = sheet.getRow(lastRow);
        if (row == null) row = sheet.createRow(lastRow);
        row.createCell(supplierCol).setCellValue(candidate);
        row.createCell(dkC).setCellValue(amount);
        row.createCell(dkH).setCellValue(amount);
        writeVlookupFormulas(sheet, lastRow, supplierCol, bankCardCol, receiverCol, idCardCol, phoneCol, bLetter, evaluator);

        boolean isLxf = isLxfCashCard(chosenReceiver);
        boolean highlight = shouldHighlight(chosenReceiver);
        if (isLxf) {
            // 刘先锋现金卡支付：使用 #7cded7 颜色，B列到L列（1-11列）
            applyRowHighlight(sheet, lastRow, 1, 11, createLxfCashCardStyle(sheet.getWorkbook()));
        } else if (highlight) {
            applyRowHighlight(sheet, lastRow, 1, 11);
        }

        return lastRow;
    }

    /**
     * 判断是否停止拆分：含"公司"/"个体工商户"/"个体户"/"刘先锋现金卡支付"任意一个
     * 等价于Python的_should_highlight_stop_split
     */
    private boolean shouldHighlightStopSplit(String name) {
        if (name == null) return false;
        if (shouldHighlight(name)) return true;
        return name.contains("刘先锋现金卡支付");
    }

    private void fillSplitRows(Workbook wb, String sheetName, int splitStartRow, List<SplitRecord> records, String remark1, String companyPrefix) throws Exception {
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) return;
        HeaderInfo headerInfo = findHeaderCols(sheet, Arrays.asList("备注1", "备注2", "是否申报个税", "申报个税月份"));
        if (headerInfo.headerRow < 0) return;
        int r1Col = headerInfo.colIndices[0], r2Col = headerInfo.colIndices[1];
        int taxCol = headerInfo.colIndices[2], monthCol = headerInfo.colIndices[3];

        // 预先创建三种带背景色的居中样式，确保备注列（I到L列）颜色与B-H列保持一致
        CellStyle yellowCenterStyle = createYellowStyle(wb);
        CellStyle lxfCenterStyle = createLxfCashCardStyle(wb);
        CellStyle normalCenterStyle = wb.createCellStyle();
        normalCenterStyle.setAlignment(HorizontalAlignment.CENTER);
        normalCenterStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        int startRow = splitStartRow;  // 使用 writeSplitRows 返回的精确起始行
        for (int i = 0; i < records.size(); i++) {
            SplitRecord rec = records.get(i);
            Row row = sheet.getRow(startRow + i);
            if (row == null) row = sheet.createRow(startRow + i);
            boolean hasKw = shouldHighlight(rec.receiver);
            boolean isLxf = rec.receiver != null && rec.receiver.contains("刘先锋现金卡支付");
            String r2, tax, month;
            if (isLxf) {
                r2 = "刘先锋现金卡支付"; tax = "刘先锋现金卡支付"; month = "/";
            } else {
                r2 = hasKw ? (companyPrefix + "对公打款") : (companyPrefix + "代发打款");
                if (hasKw) { tax = "否，对公打款，后续对接开票"; month = "/"; }
                else { tax = "是"; month = null; }
            }

            // 根据收款人判断使用哪种样式，确保I-L列背景色与B-H列一致
            CellStyle styleToUse = isLxf ? lxfCenterStyle : (hasKw ? yellowCenterStyle : normalCenterStyle);

            // 备注列（I到L列）填充值并设置样式（保留背景色+居中）
            Cell r1Cell = row.createCell(r1Col);
            r1Cell.setCellValue(remark1);
            r1Cell.setCellStyle(styleToUse);

            Cell r2Cell = row.createCell(r2Col);
            r2Cell.setCellValue(r2);
            r2Cell.setCellStyle(styleToUse);

            Cell taxCell = row.createCell(taxCol);
            taxCell.setCellValue(tax);
            taxCell.setCellStyle(styleToUse);

            if (month != null) {
                Cell monthCell = row.createCell(monthCol);
                monthCell.setCellValue(month);
                monthCell.setCellStyle(styleToUse);
            }
        }
    }

    private Map<String, String> buildLookupMap(Workbook wb, String sheetName) {
        Map<String, String> map = new HashMap<>();
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) return map;
        for (int r = 1; r <= Math.min(sheet.getLastRowNum(), 5000); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            String key = getStr(row.getCell(0));
            String val = getStr(row.getCell(1));
            if (!key.isEmpty()) map.put(key, val);
        }
        return map;
    }

    private void writeVlookupFormulas(Sheet sheet, int row, int supplierCol, int bankCardCol, int receiverCol, int idCardCol, int phoneCol, String bLetter, FormulaEvaluator evaluator) {
        String range = "'供应商信息表（定稿）'!$A:$E";

        Row r = sheet.getRow(row);
        Workbook wb = sheet.getWorkbook();
        
        // 创建居中样式
        CellStyle centerStyle = wb.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        // B列到L列（0-based: 1到11）设置居中
        for (int col = 1; col <= 11; col++) {
            Cell cell = r.getCell(col);
            if (cell == null) {
                cell = r.createCell(col);
            }
            cell.setCellStyle(centerStyle);
        }

        // 银行卡号：第4列
        Cell bankCardCell = r.getCell(bankCardCol);
        if (bankCardCell == null) {
            bankCardCell = r.createCell(bankCardCol);
        }
        bankCardCell.setCellFormula("VLOOKUP(" + bLetter + (row + 1) + "," + range + ",4,0)");
        bankCardCell.setCellStyle(centerStyle);
        
        // 收款姓名：第2列
        Cell receiverCell = r.getCell(receiverCol);
        if (receiverCell == null) {
            receiverCell = r.createCell(receiverCol);
        }
        receiverCell.setCellFormula("VLOOKUP(" + bLetter + (row + 1) + "," + range + ",2,0)");
        receiverCell.setCellStyle(centerStyle);
        
        // 身份证号：第3列
        Cell idCardCell = r.getCell(idCardCol);
        if (idCardCell == null) {
            idCardCell = r.createCell(idCardCol);
        }
        idCardCell.setCellFormula("VLOOKUP(" + bLetter + (row + 1) + "," + range + ",3,0)");
        idCardCell.setCellStyle(centerStyle);
        
        // 电话号码：第5列
        Cell phoneCell = r.getCell(phoneCol);
        if (phoneCell == null) {
            phoneCell = r.createCell(phoneCol);
        }
        phoneCell.setCellFormula("VLOOKUP(" + bLetter + (row + 1) + "," + range + ",5,0)");
        phoneCell.setCellStyle(centerStyle);

        // 立即计算公式，确保D-G列显示正确值
        evaluator.evaluateFormulaCell(bankCardCell);
        evaluator.evaluateFormulaCell(receiverCell);
        evaluator.evaluateFormulaCell(idCardCell);
        evaluator.evaluateFormulaCell(phoneCell);
    }

    private boolean shouldHighlight(String name) {
        if (name == null) return false;
        for (String kw : HIGHLIGHT_KEY_WORDS) if (name.contains(kw)) return true;
        return false;
    }

    private boolean isLxfCashCard(String name) {
        return name != null && name.contains("刘先锋现金卡支付");
    }

    private void applyRowHighlight(Sheet sheet, int row, int startCol, int endCol) {
        applyRowHighlight(sheet, row, startCol, endCol, createYellowStyle(sheet.getWorkbook()));
    }

    private void applyRowHighlight(Sheet sheet, int row, int startCol, int endCol, CellStyle style) {
        Row r = sheet.getRow(row);
        if (r == null) return;
        for (int c = startCol; c <= endCol; c++) {
            Cell cell = r.getCell(c);
            if (cell == null) cell = r.createCell(c);
            cell.setCellStyle(style);
        }
    }

    private CellStyle createYellowStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 添加居中对齐
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private CellStyle createSteelBlueStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        if (wb instanceof XSSFWorkbook) {
            XSSFColor steelBlue = new XSSFColor(new byte[]{(byte)0x46, (byte)0x82, (byte)0xB4}, null);
            style.setFillForegroundColor(steelBlue);
        } else {
            style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        }
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 添加居中对齐
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    /**
     * 创建刘先锋现金卡支付专用样式（#7cded7）
     * 等价于Python中对刘先锋现金卡支付行的颜色标记
     */
    private CellStyle createLxfCashCardStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        if (wb instanceof XSSFWorkbook) {
            // #7cded7 颜色：RGB(124, 222, 215)
            XSSFColor lxfColor = new XSSFColor(new byte[]{(byte)0x7c, (byte)0xde, (byte)0xd7}, null);
            style.setFillForegroundColor(lxfColor);
        } else {
            style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        }
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 添加居中对齐
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private void copyHeaderRows(Sheet source, Sheet target, int headerRow) {
        for (int r = 0; r <= headerRow; r++) {
            Row sr = source.getRow(r);
            if (sr == null) continue;
            Row dr = target.createRow(r);
            for (int c = 0; c <= sr.getLastCellNum(); c++) {
                Cell sc = sr.getCell(c);
                if (sc != null) {
                    Cell dc = dr.createCell(c);
                    dc.setCellValue(getStr(sc));
                }
            }
        }
        for (CellRangeAddress m : source.getMergedRegions()) {
            if (m.getFirstRow() <= headerRow) target.addMergedRegion(m);
        }
    }

    private int findLastDataRow(Sheet sheet, int headerRow, List<Integer> colsToCheck) {
        for (int r = sheet.getLastRowNum(); r > headerRow; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            // 只检查关键列，过滤各种空值和无效值
            for (int c : colsToCheck) {
                Cell cell = row.getCell(c);
                if (cell == null) continue;
                if (cell.getCellType() == CellType.BLANK) continue;
                // 公式单元格：如果计算结果是错误值则忽略
                if (cell.getCellType() == CellType.FORMULA &&
                    cell.getCachedFormulaResultType() == CellType.ERROR) continue;
                // 字符串单元格：如果内容为空或者是#N/A则忽略
                if (cell.getCellType() == CellType.STRING) {
                    String val = cell.getStringCellValue().trim();
                    if (val.isEmpty() || "#N/A".equals(val)) continue;
                }
                return r;
            }
        }
        return headerRow;
    }

    private HeaderInfo findHeaderCols(Sheet sheet, List<String> keywords) {
        for (int r = 0; r < Math.min(3, sheet.getLastRowNum() + 1); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Map<String, Integer> found = new HashMap<>();
            for (Cell cell : row) {
                if (cell == null) continue;
                String v = normalizeHeader(getStr(cell));
                for (String kw : keywords) {
                    // 精确匹配：归一化后完全相等
                    if (v.equals(normalizeHeader(kw))) { found.put(kw, cell.getColumnIndex()); break; }
                }
            }
            if (found.size() == keywords.size()) {
                int[] result = new int[keywords.size()];
                for (int i = 0; i < keywords.size(); i++) result[i] = found.get(keywords.get(i));
                return new HeaderInfo(r, result);
            }
        }
        int[] empty = new int[keywords.size()];
        Arrays.fill(empty, -1);
        return new HeaderInfo(-1, empty);
    }

    private String normalizeHeader(String s) {
        return s == null ? "" : s.trim().replace(":", "：").replace(" ", "").replace("　", "");
    }

    private String getStr(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case FORMULA:
                // 公式单元格：根据公式计算结果的实际类型来取值
                switch (cell.getCachedFormulaResultType()) {
                    case STRING:
                        return cell.getStringCellValue();
                    case NUMERIC:
                        return String.valueOf(cell.getNumericCellValue());
                    default:
                        return "";
                }
            default:
                return "";
        }
    }

    private Double parseDbl(Cell cell) {
        if (cell == null) return null;
        try {
            switch (cell.getCellType()) {
                case NUMERIC:
                    return cell.getNumericCellValue();
                case STRING: {
                    String s = cell.getStringCellValue().trim();
                    if (s.isEmpty()) return null;
                    return Double.parseDouble(s);
                }
                case FORMULA:
                    // 公式单元格：直接获取公式计算的数值结果（等价于 Python openpyxl data_only=True）
                    return cell.getNumericCellValue();
                default:
                    return null;
            }
        } catch (Exception e) {
            return null;
        }
    }

    /** 逐个安全求值公式单元格，跨工作簿引用失败时静默回退 */
    private void tryEvaluateCell(Cell cell, FormulaEvaluator evaluator) {
        if (cell == null || cell.getCellType() != CellType.FORMULA) return;
        try {
            evaluator.evaluateFormulaCell(cell);
        } catch (Exception ignored) {
            // 跨工作簿引用等无法求值，保留缓存值
        }
    }

    private String parseRemark1FromFilename(String filename) {
        // 先去掉扩展名
        String name = filename.contains(".") ? filename.substring(0, filename.lastIndexOf('.')) : filename;
        // 去掉 Controller 添加的 UUID 前缀（格式：xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx_）
        name = stripUuidPrefix(name);
        // 去掉开头的年份前缀（如 "26年"）
        name = name.replaceAll("^\\d+年", "");
        // 取第一个分隔符之前的部分
        name = name.split("[（\\(-]")[0].trim();
        return name.replace("代工费", "服务费");
    }

    /** 去除 Controller 添加的 UUID 前缀，返回原始文件名 */
    private String stripUuidPrefix(String filename) {
        return filename.replaceFirst("^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}_", "");
    }

    private static class SummaryRecord {
        String name; double amount;
        SummaryRecord(String n, double a) { name = n; amount = a; }
    }
    private static class AppendedRecord {
        String name; double amount; String receiver;
        AppendedRecord(String n, double a, String r) { name = n; amount = a; receiver = r; }
    }
    private static class SplitRecord {
        String name; double amount; String receiver;
        SplitRecord(String n, double a, String r) { name = n; amount = a; receiver = r; }
    }
    private static class SplitResult {
        List<SplitRecord> records;
        int splitStartRow;
        List<String> step5Details;  // 逐条拆分写入日志
        List<String> step5Result;   // 拆分结果汇总日志
        SplitResult(List<SplitRecord> r, int s) {
            records = r; splitStartRow = s;
            step5Details = new ArrayList<>();
            step5Result = new ArrayList<>();
        }
    }

    private static class HeaderInfo {
        int headerRow;        // 表头所在行号（0-based）
        int[] colIndices;     // 各关键字对应的列索引
        HeaderInfo(int row, int[] cols) { headerRow = row; colIndices = cols; }
    }
    public static class ProcessResult {
        public String outputPath;       // 发放记录表路径（含拆分结果）
        public String summaryPath;       // 代工费表格路径（单个/最后一个）
        public List<String> summaryPaths; // 批量模式：所有汇总后代工费表格路径
        public int summaryCount;
        public int splitCount;
        public String remark1;
        public String companyPrefix;
        public List<String> step4Details;  // 步骤4逐条判断详情
        public List<String> step5Details;  // 步骤5逐条拆分写入详情
        public List<String> step5Result;   // 步骤5拆分结果汇总
    }
}
