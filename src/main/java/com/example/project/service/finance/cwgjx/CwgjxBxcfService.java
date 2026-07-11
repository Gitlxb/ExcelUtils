package com.example.project.service.finance.cwgjx;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;

/**
 * 财务工具箱 - 报销拆分
 * 拆分合并单元格、拆分第二行多字段、删除第一行、按列拆分（bxcf_frame.py）
 */
@Service
public class CwgjxBxcfService {

    // 支付账户名称简写映射
    private static final Map<String, String> ACCOUNT_NAME_MAP = new HashMap<>();
    static {
        ACCOUNT_NAME_MAP.put("刘先锋个人现金卡", "刘先锋现金卡");
        ACCOUNT_NAME_MAP.put("温州鑫锦途供应链服务有限公司", "温州鑫锦途");
        ACCOUNT_NAME_MAP.put("浙江锦途人力资源有限公司", "浙江锦途");
        ACCOUNT_NAME_MAP.put("温州锦途服务外包有限公司", "温州锦途");
        ACCOUNT_NAME_MAP.put("衢州鑫途服务外包有限公司", "衢州鑫途");
        ACCOUNT_NAME_MAP.put("浙江锦途人力资源有限公司乐清分公司", "浙江锦途（乐清分公司）");
        ACCOUNT_NAME_MAP.put("芜湖才库人力资源有限公司", "芜湖才库");
        ACCOUNT_NAME_MAP.put("衢州驰锦企业服务有限公司", "衢州驰锦");
        ACCOUNT_NAME_MAP.put("安徽锦途企业服务集团有限公司", "安徽锦途");
        ACCOUNT_NAME_MAP.put("安庆市同益劳务服务有限公司", "安庆同益");
        ACCOUNT_NAME_MAP.put("台州众通人才服务有限公司", "台州众通");
        ACCOUNT_NAME_MAP.put("台州锦途人力资源服务外包有限公司", "台州锦途");
        ACCOUNT_NAME_MAP.put("温州鑫途企业服务有限公司", "温州鑫途");
        ACCOUNT_NAME_MAP.put("温州驰锦企业服务有限公司", "温州驰锦");
        ACCOUNT_NAME_MAP.put("浙江盛威安保服务有限公司", "盛威安保");
        ACCOUNT_NAME_MAP.put("温州市合静人力资源有限公司", "合静");
        ACCOUNT_NAME_MAP.put("铜陵锦途劳务有限公司", "铜陵锦途");
    }

    /**
     * 仅执行前三步处理（拆分合并单元格、拆分第二行、删除第一行、移动申请报销金额列）
     * 返回处理后的列名列表，供前端预览并选择拆分列
     */
    public List<String> processPreview(MultipartFile file, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();
        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        List<String> cols = new ArrayList<>();
        try (Workbook wb = new XSSFWorkbook(new FileInputStream(tempFile))) {
            Sheet srcSheet = wb.getSheetAt(0);

            // 步骤1：拆分合并单元格
            unmergeAndFill(srcSheet);

            // 步骤2：拆分第二行多字段（空格分隔）
            splitRow2Fields(srcSheet);

            // 步骤3：删除第一行（shiftRows 前先保存原始最后行号）
            int originalLastRowNum = srcSheet.getLastRowNum();
            srcSheet.shiftRows(1, originalLastRowNum, -1);
            // 删除 shiftRows 后在末尾留下的空行（原最后一行上移后，原位置变成空行）
            Row spare = srcSheet.getRow(originalLastRowNum);
            if (spare != null) {
                srcSheet.removeRow(spare);
            }

            // 步骤4：读取列名，移动"申请报销金额"到C列右侧
            Row header = srcSheet.getRow(0);
            int amountIdx = -1;
            for (int c = 0; c <= header.getLastCellNum(); c++) {
                String v = getStr(header.getCell(c));
                cols.add(v);
                if ("申请报销金额".equals(v)) amountIdx = cols.size() - 1;
            }

            if (amountIdx >= 0 && amountIdx != 3) {
                String amountCol = cols.remove(amountIdx);
                if (cols.size() >= 3) {
                    cols.add(3, amountCol);
                } else {
                    cols.add(amountCol);
                }
            }
        } finally {
            tempFile.delete();
        }
        return cols;
    }

    public String process(MultipartFile file, String splitColumn, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();
        File tempFile = new File(dir, UUID.randomUUID() + "_" + file.getOriginalFilename());
        file.transferTo(tempFile);

        try (Workbook wb = new XSSFWorkbook(new FileInputStream(tempFile));
             Workbook newWb = new XSSFWorkbook()) {

            Sheet srcSheet = wb.getSheetAt(0);

            // 1. 复制原始数据工作表（保留格式）
            Sheet originalCopy = newWb.createSheet("原始数据");
            copySheet(srcSheet, originalCopy, newWb);

            // 2. 处理合并单元格
            unmergeAndFill(srcSheet);

            // 3. 拆分第二行多字段（空格分隔）
            splitRow2Fields(srcSheet);

            // 4. 删除第一行（shiftRows 前先保存原始最后行号）
            int originalLastRowNum = srcSheet.getLastRowNum();
            srcSheet.shiftRows(1, originalLastRowNum, -1);
            // 删除 shiftRows 后在末尾留下的空行
            Row spare = srcSheet.getRow(originalLastRowNum);
            if (spare != null) {
                srcSheet.removeRow(spare);
            }

            // 5. 读取数据并移动"申请报销金额"列
            Sheet processed = newWb.createSheet("处理后数据");
            Map<Integer, String> colMap = copyDataWithMoveAmount(srcSheet, processed);

            // 5.1 为"处理后数据"工作表添加合计行和标黄
            CellStyle yellow = newWb.createCellStyle();
            yellow.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            yellow.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            int amountColIdx = -1;
            for (Map.Entry<Integer, String> e : colMap.entrySet()) {
                if ("申请报销金额".equals(e.getValue())) {
                    amountColIdx = e.getKey();
                    break;
                }
            }

            if (amountColIdx >= 0) {
                double totalAmount = 0;
                for (int r = 1; r <= processed.getLastRowNum(); r++) {
                    Row row = processed.getRow(r);
                    if (row == null) continue;
                    Cell amtCell = row.getCell(amountColIdx);
                    if (amtCell != null && amtCell.getCellType() == CellType.NUMERIC) {
                        totalAmount += amtCell.getNumericCellValue();
                    }
                }

                int totalRowIdx = processed.getLastRowNum() + 1;
                Row totalRow = processed.createRow(totalRowIdx);

                int memColIdx = -1;
                for (Map.Entry<Integer, String> e : colMap.entrySet()) {
                    if ("报销成员".equals(e.getValue())) {
                        memColIdx = e.getKey();
                        break;
                    }
                }

                if (memColIdx >= 0) {
                    Cell memCell = totalRow.createCell(memColIdx);
                    memCell.setCellValue("合计");
                    memCell.setCellStyle(yellow);
                } else {
                    Cell memCell = totalRow.createCell(2);
                    memCell.setCellValue("合计");
                    memCell.setCellStyle(yellow);
                }

                Cell amtCell = totalRow.createCell(amountColIdx);
                amtCell.setCellValue(totalAmount);
                amtCell.setCellStyle(yellow);
            }

            // 6. 按指定列拆分
            CellStyle yellowSplit = newWb.createCellStyle();
            yellowSplit.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
            yellowSplit.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            
            Map<String, List<RowData>> groups = groupByColumn(processed, splitColumn, splitColumn.equals("支付账户"));
            
            for (Map.Entry<String, List<RowData>> entry : groups.entrySet()) {
                String groupName = entry.getKey();
                // 简写逻辑已前移到 groupByColumn 中，此处直接使用
                groupName = groupName.replaceAll("[\\\\/:*?\"" +
                        "<>|]", "_");
                if (groupName.length() > 31) groupName = groupName.substring(0, 31);
            
                Sheet gs = newWb.createSheet(groupName);
                writeGroupData(gs, entry.getValue(), yellowSplit);
            }

            String outputPath = uploadDir + File.separator + "报销拆分结果.xlsx";
            try (FileOutputStream fos = new FileOutputStream(outputPath)) { newWb.write(fos); }
            return outputPath;
        } finally {
            tempFile.delete();
        }
    }

    private void unmergeAndFill(Sheet sheet) {
        List<CellRangeAddress> merged = new ArrayList<>(sheet.getMergedRegions());
        for (CellRangeAddress range : merged) {
            Row firstRow = sheet.getRow(range.getFirstRow());
            if (firstRow == null) continue;
            Cell topLeft = firstRow.getCell(range.getFirstColumn());
            Object value = topLeft != null ? getValue(topLeft) : null;
            for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                    Row row = sheet.getRow(r);
                    if (row == null) row = sheet.createRow(r);
                    Cell cell = row.getCell(c);
                    if (cell == null) cell = row.createCell(c);
                    if (value != null) setValue(cell, value);
                }
            }
            sheet.removeMergedRegion(sheet.getMergedRegions().indexOf(range));
        }
    }

    private void splitRow2Fields(Sheet sheet) {
        Row row2 = sheet.getRow(1);
        if (row2 == null) return;
        List<Object[]> inserts = new ArrayList<>();
        for (int c = row2.getLastCellNum() - 1; c >= 0; c--) {
            Cell cell = row2.getCell(c);
            if (cell == null || cell.getCellType() != CellType.STRING) continue;
            String v = cell.getStringCellValue();
            if (v == null || !v.trim().contains(" ")) continue;
            String[] fields = v.trim().split("\\s+");
            if (fields.length <= 1) continue;
            inserts.add(new Object[]{c, fields.length - 1, fields});
        }
        for (Object[] ins : inserts) {
            int col = (int) ins[0], count = (int) ins[1];
            String[] fields = (String[]) ins[2];
            sheet.shiftColumns(col + 1, sheet.getRow(0).getLastCellNum(), count);
            for (int i = 0; i < fields.length; i++) {
                for (int r = 0; r <= 2; r++) {
                    Row row = sheet.getRow(r);
                    if (row == null) row = sheet.createRow(r);
                    Cell cell = row.getCell(col + i);
                    if (cell == null) cell = row.createCell(col + i);
                    if (r == 1) cell.setCellValue(fields[i]);
                }
            }
        }
    }

    private Map<Integer, String> copyDataWithMoveAmount(Sheet src, Sheet dst) {
        Map<Integer, String> colMap = new LinkedHashMap<>();
        Row header = src.getRow(0);
        int amountCol = -1;
        for (int c = 0; c <= header.getLastCellNum(); c++) {
            Cell cell = header.getCell(c);
            String v = cell != null ? getStr(cell) : "";
            colMap.put(c, v);
            if (v.equals("申请报销金额")) amountCol = c;
        }

        // 移动"申请报销金额"到第3列（索引3，即D列右侧的C列右侧）
        Map<Integer, String> newColMap = new LinkedHashMap<>();
        int nc = 0;
        for (Map.Entry<Integer, String> e : colMap.entrySet()) {
            if (e.getValue().equals("申请报销金额")) continue;
            newColMap.put(nc++, e.getValue());
            if (nc == 3) newColMap.put(nc++, "申请报销金额");
        }
        if (nc <= 3) newColMap.put(nc, "申请报销金额");

        // 写入表头
        Row dh = dst.createRow(0);
        for (Map.Entry<Integer, String> e : newColMap.entrySet()) dh.createCell(e.getKey()).setCellValue(e.getValue());

        // 写入数据行（使用连续的行号）
        int dstRowNum = 1;
        for (int r = 1; r <= src.getLastRowNum(); r++) {
            Row sr = src.getRow(r);
            if (sr == null) continue;
            
            Row dr = dst.createRow(dstRowNum++);
            int dc = 0;
            boolean amountWritten = false;
            
            for (Map.Entry<Integer, String> e : colMap.entrySet()) {
                int srcCol = e.getKey();
                String colName = e.getValue();
                
                // 跳过申请报销金额列（稍后在正确位置写入）
                if (colName.equals("申请报销金额")) continue;
                
                // 在dc==3时先写入申请报销金额
                if (dc == 3 && !amountWritten && amountCol >= 0) {
                    Cell sac = sr.getCell(amountCol);
                    if (sac != null) copyCell(sac, dr.createCell(dc));
                    dc++;
                    amountWritten = true;
                }
                
                Cell sc = sr.getCell(srcCol);
                if (sc != null) copyCell(sc, dr.createCell(dc++));
            }
            
            // 如果申请报销金额还没有写入（在最后或列数<=3的情况）
            if (!amountWritten && amountCol >= 0) {
                Cell sac = sr.getCell(amountCol);
                if (sac != null) copyCell(sac, dr.createCell(dc));
            }
        }
        return newColMap;
    }

    private Map<String, List<RowData>> groupByColumn(Sheet sheet, String colName, boolean isAccountColumn) {
        Map<String, List<RowData>> groups = new LinkedHashMap<>();
        Row header = sheet.getRow(0);
        int colIdx = -1;
        for (int c = 0; c <= header.getLastCellNum(); c++) {
            if (colName.equals(getStr(header.getCell(c)))) { colIdx = c; break; }
        }
        if (colIdx == -1) throw new RuntimeException("找不到列: " + colName);

        List<String> columns = new ArrayList<>();
        for (int c = 0; c <= header.getLastCellNum(); c++) {
            String v = getStr(header.getCell(c));
            if (!v.isEmpty()) columns.add(v);
        }

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            
            Map<String, Object> data = new LinkedHashMap<>();
            for (int c = 0; c < columns.size(); c++) {
                data.put(columns.get(c), getValue(row.getCell(c)));
            }
            
            // 获取分组列的值（用于分组键）
            String key = getStr(row.getCell(colIdx));
            if (isAccountColumn) {
                key = ACCOUNT_NAME_MAP.getOrDefault(key, key);
            }
            
            // 即使分组列为空也保留数据，只是不加入任何分组
            if (!key.isEmpty()) {
                groups.computeIfAbsent(key, k -> new ArrayList<>()).add(new RowData(data));
            }
        }
        return groups;
    }

    private void writeGroupData(Sheet sheet, List<RowData> rows, CellStyle yellow) {
        if (rows.isEmpty()) return;
        Set<String> cols = rows.get(0).data.keySet();
        // 表头
        Row h = sheet.createRow(0);
        int ci = 0;
        for (String c : cols) h.createCell(ci++).setCellValue(c);

        // 写入数据行并累计总金额
        double total = 0;
        for (int i = 0; i < rows.size(); i++) {
            Row dr = sheet.createRow(i + 1);
            int cj = 0;
            for (String c : cols) {
                Object v = rows.get(i).data.get(c);
                setValue(dr.createCell(cj++), v);
            }
            Object amt = rows.get(i).data.get("申请报销金额");
            if (amt instanceof Number) total += ((Number) amt).doubleValue();
        }

        // 合计行（数据行之后）
        int tr = rows.size() + 1;
        Row totalRow = sheet.createRow(tr);
        int memIdx = -1, amtIdx = -1;
        int idx = 0;
        for (String c : cols) {
            if ("报销成员".equals(c)) memIdx = idx;
            if ("申请报销金额".equals(c)) amtIdx = idx;
            idx++;
        }
        if (memIdx >= 0) totalRow.createCell(memIdx).setCellValue("合计");
        if (amtIdx >= 0) {
            Cell ac = totalRow.createCell(amtIdx);
            ac.setCellValue(total); ac.setCellStyle(yellow);
        }
        if (memIdx >= 0) totalRow.getCell(memIdx).setCellStyle(yellow);

        // 3个空行（与Python一致）
        for (int e = 1; e <= 3; e++) {
            sheet.createRow(tr + e);
        }

        // 报销成员去重统计（与Python一致）
        List<String> uniqueMembers = new ArrayList<>();
        if (memIdx >= 0) {
            for (RowData rd : rows) {
                Object m = rd.data.get("报销成员");
                if (m != null) {
                    String ms = m.toString().trim();
                    if (!ms.isEmpty() && !"合计".equals(ms) && !uniqueMembers.contains(ms)) {
                        uniqueMembers.add(ms);
                    }
                }
            }
        }
        int statRowIdx = tr + 4; // 合计行 + 3空行 之后
        double grandTotal = 0;
        for (String member : uniqueMembers) {
            Row sr = sheet.createRow(statRowIdx++);
            double memberAmount = 0;
            for (RowData rd : rows) {
                Object m = rd.data.get("报销成员");
                if (m != null && member.equals(m.toString().trim())) {
                    Object a = rd.data.get("申请报销金额");
                    if (a instanceof Number) memberAmount += ((Number) a).doubleValue();
                }
            }
            grandTotal += memberAmount;
            idx = 0;
            for (String c : cols) {
                if ("报销成员".equals(c)) sr.createCell(idx).setCellValue(member);
                else if ("申请报销金额".equals(c)) {
                    Cell ac = sr.createCell(idx);
                    ac.setCellValue(memberAmount);
                    // 统计行不标黄，只有合计行标黄
                }
                idx++;
            }
        }

        // 总计行
        if (!uniqueMembers.isEmpty()) {
            Row gr = sheet.createRow(statRowIdx);
            idx = 0;
            for (String c : cols) {
                if ("报销成员".equals(c)) gr.createCell(idx).setCellValue("总计");
                else if ("申请报销金额".equals(c)) {
                    Cell ac = gr.createCell(idx);
                    ac.setCellValue(grandTotal);
                    // 总计行不标黄，只有合计行标黄
                }
                idx++;
            }
        }

        // 标黄合计行（与Python一致：找到"合计"所在行，标黄第3、4列）
        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cCell = row.getCell(memIdx >= 0 ? memIdx : 2);
            if (cCell != null && "合计".equals(cCell.getStringCellValue())) {
                Cell c3 = row.getCell(2);
                Cell c4 = row.getCell(3);
                if (c3 != null) c3.setCellStyle(yellow);
                if (c4 != null) c4.setCellStyle(yellow);
                break;
            }
        }
    }

    private void copySheet(Sheet src, Sheet dst, Workbook dstWb) {
        // 复制单元格数据和样式
        for (int r = 0; r <= src.getLastRowNum(); r++) {
            Row sr = src.getRow(r);
            if (sr == null) continue;
            Row dr = dst.createRow(r);
            dr.setHeight(sr.getHeight()); // 保留行高
            for (int c = 0; c <= sr.getLastCellNum(); c++) {
                Cell sc = sr.getCell(c);
                if (sc != null) {
                    Cell dc = dr.createCell(c);
                    copyCell(sc, dc);
                    // 复制单元格样式
                    CellStyle newStyle = dstWb.createCellStyle();
                    newStyle.cloneStyleFrom(sc.getCellStyle());
                    dc.setCellStyle(newStyle);
                }
            }
        }
        // 复制合并单元格
        for (CellRangeAddress m : src.getMergedRegions()) {
            dst.addMergedRegion(new CellRangeAddress(m.getFirstRow(), m.getLastRow(), m.getFirstColumn(), m.getLastColumn()));
        }
        // 复制列宽
        if (src instanceof org.apache.poi.xssf.usermodel.XSSFSheet &&
            dst instanceof org.apache.poi.xssf.usermodel.XSSFSheet) {
            for (int c = 0; c <= src.getRow(0).getLastCellNum(); c++) {
                dst.setColumnWidth(c, src.getColumnWidth(c));
                if (src.getColumnWidth(c) > 0) {
                    // 复制列宽（含隐藏列）
                    dst.setColumnWidth(c, src.getColumnWidth(c));
                }
            }
        }
    }

    private Object getValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue();
            case NUMERIC: return cell.getNumericCellValue();
            case BOOLEAN: return cell.getBooleanCellValue();
            default: return "";
        }
    }

    private void setValue(Cell cell, Object value) {
        if (value == null) { cell.setBlank(); return; }
        if (value instanceof Number) cell.setCellValue(((Number) value).doubleValue());
        else cell.setCellValue(value.toString());
    }

    private String getStr(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING: return cell.getStringCellValue().trim();
            case NUMERIC: {
                double val = cell.getNumericCellValue();
                // 如果是整数，去掉小数部分
                if (val == Math.floor(val) && !Double.isInfinite(val)) {
                    return String.valueOf((long) val);
                }
                return String.valueOf(val);
            }
            case BOOLEAN: return String.valueOf(cell.getBooleanCellValue());
            default: return "";
        }
    }

    private void copyCell(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING: dst.setCellValue(src.getStringCellValue()); break;
            case NUMERIC: dst.setCellValue(src.getNumericCellValue()); break;
            case BOOLEAN: dst.setCellValue(src.getBooleanCellValue()); break;
            case FORMULA: dst.setCellFormula(src.getCellFormula()); break;
            default: dst.setBlank(); break;
        }
    }

    private static class RowData {
        Map<String, Object> data;
        RowData(Map<String, Object> d) { data = d; }
    }
}
