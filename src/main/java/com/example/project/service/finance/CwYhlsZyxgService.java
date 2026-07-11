package com.example.project.service.finance;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

/**
 * 银行流水专业修改（cw_yhls_zyxg.py）
 * 基于"各种内容"配置工作表，对主数据工作表进行多规则批量处理
 */
@Service
public class CwYhlsZyxgService {

    // 必要列名
    private static final String[] REQUIRED_COLUMNS = {
            "流程实例号", "业务名称", "摘要", "收(付)方名称", "收(付)方账号", "借方金额", "贷方金额"
    };

    /**
     * 检测接口：读取配置表，返回选项列表供前端选择
     */
    public Map<String, Object> detectConfig(MultipartFile file) throws Exception {
        Map<String, Object> result = new HashMap<>();

        try (Workbook wb = WorkbookFactory.create(file.getInputStream())) {
            // 检查"各种内容"工作表
            Sheet configSheet = wb.getSheet("各种内容");
            if (configSheet == null) {
                throw new Exception("Excel文件中缺少工作表: 各种内容");
            }

            // 获取主数据工作表名（非"各种内容"的第一个工作表）
            String mainSheetName = null;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String name = wb.getSheetName(i);
                if (!"各种内容".equals(name)) {
                    mainSheetName = name;
                    break;
                }
            }
            if (mainSheetName == null) {
                throw new Exception("Excel文件中缺少主数据工作表");
            }

            // AS列(索引28)选项
            List<String> asOptions = new ArrayList<>();
            for (int rowIdx = 1; rowIdx <= configSheet.getLastRowNum(); rowIdx++) {
                Row row = configSheet.getRow(rowIdx);
                if (row == null) continue;
                Cell cell = row.getCell(28);
                if (cell != null) {
                    String v = getCellString(cell).trim();
                    if (!v.isEmpty() && !asOptions.contains(v)) asOptions.add(v);
                }
            }

            // J列(索引9)选项
            List<String> jOptions = new ArrayList<>();
            for (int rowIdx = 1; rowIdx <= configSheet.getLastRowNum(); rowIdx++) {
                Row row = configSheet.getRow(rowIdx);
                if (row == null) continue;
                Cell cell = row.getCell(9);
                if (cell != null) {
                    String v = getCellString(cell).trim();
                    if (!v.isEmpty() && !jOptions.contains(v)) jOptions.add(v);
                }
            }

            result.put("success", true);
            result.put("mainSheetName", mainSheetName);
            result.put("asOptions", asOptions);
            result.put("jOptions", jOptions);
        }
        return result;
    }

    /**
     * 处理接口：执行所有规则
     */
    public String process(MultipartFile file,
                          String selectedAs,
                          String selectedCompany,
                          String inputNameFilter,
                          String inputSummaryText,
                          String inputSummary,
                          String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        try (Workbook wb = WorkbookFactory.create(new FileInputStream(tempFile))) {
            Sheet configSheet = wb.getSheet("各种内容");
            if (configSheet == null) {
                throw new Exception("Excel文件中缺少工作表: 各种内容");
            }

            // 找主数据工作表
            Sheet mainSheet = null;
            String mainSheetName = null;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                String name = wb.getSheetName(i);
                if (!"各种内容".equals(name)) {
                    mainSheet = wb.getSheetAt(i);
                    mainSheetName = name;
                    break;
                }
            }
            if (mainSheet == null) throw new Exception("Excel文件中缺少主数据工作表");

            // ========== 读取配置 ==========
            // A列(0): 代发其他短语
            List<String> replacePhrases = readColumnStrings(configSheet, 0, 1);
            // B列(1): 退款文本
            List<String> refundTexts = readColumnStrings(configSheet, 1, 1);
            String refundText = refundTexts.isEmpty() ? "交易失败退款" : refundTexts.get(0);
            // D列(3): 手续费类型
            List<String> feeTypes = readColumnStrings(configSheet, 3, 1);
            // E列(4): 手续费替换内容
            List<String> feeReplacements = readColumnStrings(configSheet, 4, 1);
            // G列(6): 投资款目标名称
            List<String> targetNames = readColumnStrings(configSheet, 6, 1);
            // I列(8)和J列(9): 往来款映射
            Map<String, String> transferMapping = new HashMap<>();
            for (int rowIdx = 1; rowIdx <= configSheet.getLastRowNum(); rowIdx++) {
                Row row = configSheet.getRow(rowIdx);
                if (row == null) continue;
                String source = getCellString(row.getCell(8)).trim();
                String target = getCellString(row.getCell(9)).trim();
                if (!source.isEmpty() && !target.isEmpty()) {
                    transferMapping.put(source, target);
                }
            }
            // L列(11): 内部转账账号配置（未直接使用，AL列才是实际账号列表）
            // N列(13): 需要清除的名称内容
            List<String> namesToClear = readColumnStrings(configSheet, 13, 1);
            // P列(15): 交易类型筛选
            List<String> transactionTypes = readColumnStrings(configSheet, 15, 1);
            // R列(17): 收(付)方名称筛选
            List<String> paymentNames = readColumnStrings(configSheet, 17, 1);
            // T列(19): 新增收(付)方名称筛选
            List<String> paymentNamesT = readColumnStrings(configSheet, 19, 1);
            // V列(21): 新增摘要筛选
            List<String> summaryTexts = readColumnStrings(configSheet, 21, 1);
            // X列(23): 联动代发收(付)方名称
            List<String> lianDongPaymentNames = readColumnStrings(configSheet, 23, 1);
            // Z列(25): 标黄关键词
            List<String> highlightNames = readColumnStrings(configSheet, 25, 1);
            // AB列(27): 交易类型关键词
            List<String> transactionKeywords = readColumnStrings(configSheet, 27, 1);
            // AD列(29): 行标黄_摘要
            List<String> summaryMatchList = readColumnStrings(configSheet, 29, 1);
            // AF列(31): 行标黄_收(付)方名称
            List<String> gsGsWzj = readColumnStrings(configSheet, 31, 1);
            // BB列(53): 交易类型筛选条件
            List<String> bbFilters = readColumnStrings(configSheet, 53, 1);
            // BC列(54): 摘要替换内容
            List<String> bcReplacements = readColumnStrings(configSheet, 54, 1);
            // BG列(58): 跳过处理关键词
            List<String> bgKeywords = readColumnStrings(configSheet, 58, 1);
            // AL列(37): 内部转账账号列表
            List<String> alAccounts = readColumnStrings(configSheet, 37, 1);

            // 交易类型映射 BB->BC
            Map<String, String> transactionTypeMapping = new HashMap<>();
            for (int i = 0; i < Math.min(bbFilters.size(), bcReplacements.size()); i++) {
                transactionTypeMapping.put(bbFilters.get(i), bcReplacements.get(i));
            }

            // 手续费映射
            Map<String, String> feeMapping = new HashMap<>();
            for (int i = 0; i < Math.min(feeTypes.size(), feeReplacements.size()); i++) {
                feeMapping.put(feeTypes.get(i), feeReplacements.get(i));
            }
            for (int i = feeReplacements.size(); i < feeTypes.size(); i++) {
                feeMapping.put(feeTypes.get(i), "银行手续费");
            }

            // ========== 读取主数据并构建列索引 ==========
            Row headerRow = mainSheet.getRow(0);
            if (headerRow == null) throw new Exception("主数据工作表没有表头");

            Map<String, Integer> colIndexMap = new HashMap<>();
            for (Cell cell : headerRow) {
                String h = getCellString(cell).trim();
                if (!h.isEmpty()) colIndexMap.put(h, cell.getColumnIndex());
            }

            // 检查必要列
            for (String col : REQUIRED_COLUMNS) {
                if (!colIndexMap.containsKey(col)) {
                    throw new Exception("主数据工作表中缺少必要的列: " + col);
                }
            }

            int colFlowId = colIndexMap.get("流程实例号");
            int colBizName = colIndexMap.get("业务名称");
            int colSummary = colIndexMap.get("摘要");
            int colPaymentName = colIndexMap.get("收(付)方名称");
            int colPaymentAccount = colIndexMap.get("收(付)方账号");
            int colDebit = colIndexMap.get("借方金额");
            int colCredit = colIndexMap.get("贷方金额");
            Integer colTransType = colIndexMap.get("交易类型");
            int colD = 3; // D列固定索引（与Python的iloc[:, 3]对应）

            int lastRow = mainSheet.getLastRowNum();

            // ========== 辅助数据结构 ==========
            // 流程实例号 -> 行索引列表
            Map<String, List<Integer>> flowIdGroups = new HashMap<>();
            for (int r = 1; r <= lastRow; r++) {
                Row row = mainSheet.getRow(r);
                if (row == null) continue;
                String fid = getCellString(row.getCell(colFlowId)).trim();
                if (!fid.isEmpty()) {
                    flowIdGroups.computeIfAbsent(fid, k -> new ArrayList<>()).add(r);
                }
            }
            // 流程实例号重复≥2次的
            Set<String> duplicateFlowIds = new HashSet<>();
            for (Map.Entry<String, List<Integer>> e : flowIdGroups.entrySet()) {
                if (e.getValue().size() >= 2) duplicateFlowIds.add(e.getKey());
            }

            // ========== 处理每一行 ==========
            for (int r = 1; r <= lastRow; r++) {
                Row row = mainSheet.getRow(r);
                if (row == null) continue;

                String summary = getCellString(row.getCell(colSummary));
                String paymentName = getCellString(row.getCell(colPaymentName));
                String bizName = getCellString(row.getCell(colBizName));
                String transType = colTransType != null ? getCellString(row.getCell(colTransType)) : "";
                double debit = getCellDouble(row.getCell(colDebit));
                double credit = getCellDouble(row.getCell(colCredit));

                // 全局跳过标记
                boolean skip = false;
                if (!bgKeywords.isEmpty()) {
                    for (String kw : bgKeywords) {
                        if (summary.contains(kw)) { skip = true; break; }
                    }
                }

                // ===== 规则0: 清除指定收(付)方名称内容 =====
                if (!skip && !namesToClear.isEmpty()) {
                    for (String textToClear : namesToClear) {
                        if (paymentName.contains(textToClear)) {
                            paymentName = paymentName.replace(textToClear, "");
                            row.getCell(colPaymentName).setCellValue(paymentName);
                            break;
                        }
                    }
                }

                // ===== 规则1: 代发其他 =====
                if (!skip) {
                    String flowId = getCellString(row.getCell(colFlowId)).trim();
                    if (duplicateFlowIds.contains(flowId) && !"联动代发".equals(bizName)) {
                        for (String phrase : replacePhrases) {
                            if (summary.contains(phrase)) {
                                // 找同组其他行
                                List<Integer> groupRows = flowIdGroups.get(flowId);
                                for (int otherR : groupRows) {
                                    if (otherR != r) {
                                        Row otherRow = mainSheet.getRow(otherR);
                                        if (otherRow != null) {
                                            String replacement = getCellString(otherRow.getCell(colSummary));
                                            row.getCell(colSummary).setCellValue(replacement + "_" + refundText);
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                        }
                    }
                }

                // 重新读取可能已更新的字段
                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                bizName = getCellString(row.getCell(colBizName));
                transType = colTransType != null ? getCellString(row.getCell(colTransType)) : "";
                debit = getCellDouble(row.getCell(colDebit));
                credit = getCellDouble(row.getCell(colCredit));

                // ===== 规则2: 联动代发 =====
                if (!skip && "联动代发".equals(bizName) && paymentName.contains("公司")) {
                    row.getCell(colSummary).setCellValue(paymentName);
                }

                // 重新读取
                summary = getCellString(row.getCell(colSummary));

                // ===== 规则3: 银行手续费 =====
                if (!skip && !feeMapping.isEmpty()) {
                    String mapped = feeMapping.get(summary);
                    if (mapped != null) {
                        row.getCell(colSummary).setCellValue(mapped);
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));

                // ===== 规则4: 投资款 =====
                if (!skip && !targetNames.isEmpty()) {
                    if (targetNames.contains(paymentName) && summary.contains("投资款")) {
                        row.getCell(colSummary).setCellValue("投资款，" + paymentName);
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                debit = getCellDouble(row.getCell(colDebit));
                credit = getCellDouble(row.getCell(colCredit));

                // ===== 规则5: 往来款 =====
                if (!skip && !transferMapping.isEmpty()) {
                    String pName = paymentName.trim();
                    List<String> transferCompanies = new ArrayList<>();
                    for (String name : transferMapping.keySet()) {
                        if (!name.equals(selectedAs != null ? selectedAs.trim() : "")) {
                            transferCompanies.add(name);
                        }
                    }

                    if (transferCompanies.contains(pName)) {
                        String mappedName = transferMapping.get(pName);
                        if (debit != 0) {
                            row.getCell(colSummary).setCellValue(
                                    "往来款，" + selectedCompany + "转" + mappedName);
                        }
                        if (credit != 0) {
                            row.getCell(colSummary).setCellValue(
                                    "往来款，" + mappedName + "转" + selectedCompany);
                        }
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                debit = getCellDouble(row.getCell(colDebit));
                credit = getCellDouble(row.getCell(colCredit));

                // ===== 规则6: 内部转账 =====
                if (!skip && inputNameFilter != null && !inputNameFilter.isEmpty()
                        && inputSummaryText != null && !inputSummaryText.isEmpty()
                        && !alAccounts.isEmpty()) {
                    String account = getCellString(row.getCell(colPaymentAccount)).trim();
                    if (inputNameFilter.equals(paymentName) && alAccounts.contains(account)) {
                        String lastFour = account.length() >= 4 ? account.substring(account.length() - 4) : account;
                        if (debit != 0) {
                            row.getCell(colSummary).setCellValue(
                                    "内部转账，" + inputSummaryText + "转入" + lastFour);
                        }
                        if (credit != 0) {
                            row.getCell(colSummary).setCellValue(
                                    "内部转账，" + lastFour + "转入" + inputSummaryText);
                        }
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                debit = getCellDouble(row.getCell(colDebit));
                credit = getCellDouble(row.getCell(colCredit));
                bizName = getCellString(row.getCell(colBizName));
                transType = colTransType != null ? getCellString(row.getCell(colTransType)) : "";

                // ===== 规则7: 劳务费(P+R列) =====
                if (!skip && !transactionTypes.isEmpty() && !paymentNames.isEmpty()) {
                    if (credit != 0 && transactionTypes.contains(transType)
                            && paymentNames.contains(paymentName)
                            && !summary.contains("往来款")) {
                        row.getCell(colSummary).setCellValue(paymentName + "_劳务费");
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                debit = getCellDouble(row.getCell(colDebit));

                // ===== 规则8: 劳务费(T+V列) =====
                if (!skip && !paymentNamesT.isEmpty() && !summaryTexts.isEmpty()) {
                    if (paymentNamesT.contains(paymentName) && summaryTexts.contains(summary)
                            && !summary.contains("往来款")) {
                        row.getCell(colSummary).setCellValue(paymentName + "_劳务费");
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                bizName = getCellString(row.getCell(colBizName));
                debit = getCellDouble(row.getCell(colDebit));

                // ===== 规则9: 联动代发劳务费 =====
                if (!skip && !lianDongPaymentNames.isEmpty()) {
                    if ("联动代发".equals(bizName) && lianDongPaymentNames.contains(paymentName)
                            && !summary.contains("往来款")) {
                        row.getCell(colSummary).setCellValue(paymentName + "_劳务费");
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                bizName = getCellString(row.getCell(colBizName));
                transType = colTransType != null ? getCellString(row.getCell(colTransType)) : "";

                // ===== 规则10: 锦途代发月薪 =====
                if (!skip && "联动代发".equals(bizName) && !summary.endsWith("_劳务费")) {
                    if (inputSummary != null && !inputSummary.isEmpty()) {
                        if (transType.contains("退")) {
                            row.getCell(colSummary).setCellValue(inputSummary + "_交易失败退款");
                        } else {
                            row.getCell(colSummary).setCellValue(inputSummary);
                        }
                    }
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));
                credit = getCellDouble(row.getCell(colCredit));

                // ===== 规则11: 工伤理赔 =====
                if (!skip && paymentName.contains("保险") && credit != 0
                        && !summary.contains("保费") && !summary.contains("退")) {
                    row.getCell(colSummary).setCellValue("xx员工工伤理赔，xx");
                }

                summary = getCellString(row.getCell(colSummary));
                paymentName = getCellString(row.getCell(colPaymentName));

                // ===== 规则12: 公积金 =====
                if (!skip && paymentName.contains("公积金")) {
                    // D列内容（第4列，索引3）
                    String dContent = "";
                    Row rrr = mainSheet.getRow(r);
                    if (rrr != null) {
                        Cell dCell = rrr.getCell(colD);
                        dContent = getCellString(dCell);
                    }
                    if (dContent.length() >= 7) {
                        dContent = dContent.substring(0, 7);
                    }
                    dContent = dContent.replace("-", "");
                    row.getCell(colSummary).setCellValue("缴纳" + dContent + "住房公积金");
                }

                summary = getCellString(row.getCell(colSummary));
                transType = colTransType != null ? getCellString(row.getCell(colTransType)) : "";

                // ===== 规则13: BB列交易类型映射 =====
                if (!skip && !transactionTypeMapping.isEmpty()) {
                    String mapped = transactionTypeMapping.get(transType);
                    if (mapped != null) {
                        row.getCell(colSummary).setCellValue(mapped);
                    }
                }
            }

            // ========== 标黄处理 ==========
            // 黄色样式
            XSSFColor yellowColor = new XSSFColor(new byte[]{(byte)255, (byte)255, (byte)0}, null);
            CellStyle yellowStyle = wb.createCellStyle();
            yellowStyle.setFillForegroundColor(yellowColor);
            yellowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            for (int r = 1; r <= lastRow; r++) {
                Row row = mainSheet.getRow(r);
                if (row == null) continue;

                boolean shouldHighlight = false;
                String paymentName = getCellString(row.getCell(colPaymentName));
                String transType = colTransType != null ? getCellString(row.getCell(colTransType)) : "";
                String summary = getCellString(row.getCell(colSummary));
                String bizName = getCellString(row.getCell(colBizName));
                double debit = getCellDouble(row.getCell(colDebit));

                // 标黄条件1: 收(付)方名称包含Z列关键词
                for (String kw : highlightNames) {
                    if (paymentName.contains(kw)) { shouldHighlight = true; break; }
                }

                // 标黄条件2: 交易类型包含AB列关键词
                if (!shouldHighlight) {
                    for (String kw : transactionKeywords) {
                        if (transType.contains(kw)) { shouldHighlight = true; break; }
                    }
                }

                // 标黄条件3: 摘要完全匹配AD列内容
                if (!shouldHighlight && summaryMatchList.contains(summary)) {
                    shouldHighlight = true;
                }

                // 标黄条件4: 收(付)方名称完全匹配AF列内容
                if (!shouldHighlight && gsGsWzj.contains(paymentName)) {
                    shouldHighlight = true;
                }

                // 标黄条件5: 业务名称="联动代发" 且 借方金额有值
                if (!shouldHighlight && "联动代发".equals(bizName) && debit != 0) {
                    shouldHighlight = true;
                }

                if (shouldHighlight) {
                    for (int c = 0; c <= row.getLastCellNum(); c++) {
                        Cell cell = row.getCell(c);
                        if (cell == null) cell = row.createCell(c);
                        cell.setCellStyle(yellowStyle);
                    }
                }
            }

            // ========== 保存 ==========
            String outputPath = uploadDir + File.separator +
                    mainSheetName + "_完整处理结果.xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                wb.write(fos);
            }

            // 清理临时文件
            tempFile.delete();
            return outputPath;
        }
    }

    // ========== 辅助方法 ==========

    private List<String> readColumnStrings(Sheet sheet, int colIndex, int startRow) {
        List<String> result = new ArrayList<>();
        for (int r = startRow; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(colIndex);
            if (cell != null) {
                String v = getCellString(cell).trim();
                if (!v.isEmpty()) result.add(v);
            }
        }
        return result;
    }

    private String getCellString(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                }
                double d = cell.getNumericCellValue();
                if (d == (long) d) return String.valueOf((long) d);
                return String.valueOf(d);
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            case FORMULA: return cell.getCellFormula();
            default: return "";
        }
    }

    private double getCellDouble(Cell cell) {
        if (cell == null) return 0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        }
        if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return 0;
            }
        }
        if (cell.getCellType() == CellType.FORMULA) {
            try {
                return cell.getNumericCellValue();
            } catch (Exception e) {
                return 0;
            }
        }
        return 0;
    }

    private String getHeaderAt(Row headerRow, int colIndex) {
        Cell cell = headerRow.getCell(colIndex);
        return cell != null ? getCellString(cell).trim() : "";
    }
}
