package com.example.project.service.finance.jxkh;

import com.example.project.util.ExcelUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
 * 绩效考核 - 合并表格（frame_hbbg.py）
 * 多Excel文件合并，保留原始格式
 * 与现有 CwHbbgZhfService 差异：
 * 1. 多文件直接上传（非ZIP）
 * 2. 支持 .xls 自动转换
 * 3. 先处理每个文件（填月份）再合并
 */
@Service
public class JxkhHbbgService {

    private static final Logger log = LoggerFactory.getLogger(JxkhHbbgService.class);
    private static final Pattern MONTH_PATTERN = Pattern.compile("(\\d{4})年(\\d{1,2})月");

    public String processAndMerge(List<MultipartFile> files, String targetCol, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();
        File tempDir = new File(dir, "temp_hbbg_" + System.currentTimeMillis());
        tempDir.mkdirs();

        List<FileResult> results = new ArrayList<>();
        List<File> tempsToClean = new ArrayList<>();

        try {
            int success = 0, skip = 0, error = 0;
            for (MultipartFile file : files) {
                String origName = file.getOriginalFilename();
                if (origName == null) origName = "unknown";
                // 安全处理文件名，防止路径遍历
                String safeName = origName.replaceAll(".*[/\\\\]", "");
                File tempFile = new File(tempDir, safeName);
                file.transferTo(tempFile);

                FileResult res = processSingleFile(tempFile, targetCol, tempDir, tempsToClean);
                res.fileName = safeName;
                results.add(res);
                if ("success".equals(res.status)) success++;
                else if ("skip".equals(res.status)) skip++;
                else error++;
            }

            if (success == 0) {
                StringBuilder sb = new StringBuilder("没有成功处理的文件！成功:" + success + " 跳过:" + skip + " 错误:" + error + "\n\n详细错误:\n");
                for (FileResult r : results) {
                    sb.append("  [").append(r.status).append("] ").append(r.fileName).append(": ").append(r.message).append("\n");
                    if (r.exception != null) {
                        sb.append("    异常: ").append(r.exception).append("\n");
                    }
                }
                throw new Exception(sb.toString());
            }

            // 合并所有文件（包括跳过的原始文件）
            List<String> pathsToMerge = new ArrayList<>();
            for (FileResult r : results) {
                if ("success".equals(r.status) && r.tempPath != null) {
                    pathsToMerge.add(r.tempPath);
                } else {
                    pathsToMerge.add(r.filePath);
                }
            }

            File outputFile = new File(uploadDir, "汇总表_" + System.currentTimeMillis() + ".xlsx");
            mergeFiles(pathsToMerge, outputFile);

            // 清理临时文件
            for (File f : tempsToClean) {
                if (f.exists()) f.delete();
            }

            return outputFile.getAbsolutePath();
        } finally {
            deleteDirectory(tempDir);
        }
    }

    /**
     * 通用合并模式（P0）：多文件原样堆叠合并，不填月份，保留全部格式
     * @param gapRows 文件间空行数
     * @param keepHeaderFileName 保留表头文件（null=全部保留）
     * @param skipRows 跳过表头行数（仅当指定保留表头文件时生效）
     */
    public String genericMerge(List<MultipartFile> files, String uploadDir,
                               int gapRows, String keepHeaderFileName, int skipRows,
                               List<ShiftRule> shiftRules) throws Exception {
        if (files == null || files.isEmpty()) throw new Exception("请选择文件");
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        List<String> tempPaths = new ArrayList<>();
        // 保存原始文件名和临时文件路径的映射
        Map<String, String> fileNameToTempPath = new LinkedHashMap<>();
        
        try {
            for (MultipartFile file : files) {
                String safeName = file.getOriginalFilename() != null ?
                        file.getOriginalFilename().replaceAll(".*[/\\\\]", "") : "unknown";
                File temp = new File(dir, UUID.randomUUID() + "_" + safeName);
                file.transferTo(temp);
                tempPaths.add(temp.getAbsolutePath());
                // 保存映射：原始文件名 → 临时文件路径
                fileNameToTempPath.put(file.getOriginalFilename(), temp.getAbsolutePath());
            }

            // 新增：在合并前，对需要移动列的文件执行列移动
            if (shiftRules != null && !shiftRules.isEmpty()) {
                // 按文件名分组：同一文件的多条规则合并为一次IO
                Map<String, List<ShiftRule>> rulesByFile = new LinkedHashMap<>();
                for (ShiftRule rule : shiftRules) {
                    rulesByFile.computeIfAbsent(rule.getFileName(), k -> new ArrayList<>()).add(rule);
                }

                log.info("列移动规则按文件分组：共 {} 个文件，{} 条规则", rulesByFile.size(), shiftRules.size());

                for (Map.Entry<String, List<ShiftRule>> entry : rulesByFile.entrySet()) {
                    String targetFileName = entry.getKey();
                    List<ShiftRule> fileRules = entry.getValue();

                    String targetTempPath = fileNameToTempPath.get(targetFileName);
                    if (targetTempPath == null) {
                        log.warn("未找到文件 {} 对应的临时文件，跳过此文件的所有规则", targetFileName);
                        continue;
                    }

                    if (fileRules.size() == 1) {
                        // 只有一条规则，直接调用原方法
                        ShiftRule rule = fileRules.get(0);
                        log.info("对文件 {} 的列 {} 执行右移 {} 格", targetFileName,
                                colIndexToLetter(rule.getColIndex()), rule.getShiftCount());
                        shiftColumnRight(new File(targetTempPath), rule.getColIndex(), rule.getShiftCount());
                    } else {
                        // 多条规则，合并为一次IO
                        log.info("对文件 {} 执行 {} 条列移动规则（合并为单次IO）", targetFileName, fileRules.size());
                        shiftMultipleColumns(new File(targetTempPath), fileRules);
                    }
                }
            }

            boolean[] skipFlags = new boolean[files.size()];
            int[] skipAmounts = new int[files.size()];
            if (keepHeaderFileName != null && !keepHeaderFileName.isEmpty() && !"（不指定）".equals(keepHeaderFileName)) {
                for (int i = 0; i < files.size(); i++) {
                    String origName = files.get(i).getOriginalFilename();
                    if (origName != null && origName.equals(keepHeaderFileName)) {
                        skipFlags[i] = false;
                    } else {
                        skipFlags[i] = true;
                        skipAmounts[i] = skipRows;
                    }
                }
            }

            // 将保留表头文件移到合并结果的第一位
            if (keepHeaderFileName != null && !keepHeaderFileName.isEmpty() && !"（不指定）".equals(keepHeaderFileName)) {
                int headerIndex = -1;
                for (int i = 0; i < files.size(); i++) {
                    String origName = files.get(i).getOriginalFilename();
                    if (origName != null && origName.equals(keepHeaderFileName)) {
                        headerIndex = i;
                        break;
                    }
                }

                if (headerIndex > 0) {
                    // 同步重排 tempPaths、skipFlags、skipAmounts
                    String headerTempPath = tempPaths.remove(headerIndex);
                    tempPaths.add(0, headerTempPath);

                    boolean headerSkipFlag = skipFlags[headerIndex];
                    System.arraycopy(skipFlags, 0, skipFlags, 1, headerIndex);
                    skipFlags[0] = headerSkipFlag;

                    int headerSkipAmount = skipAmounts[headerIndex];
                    System.arraycopy(skipAmounts, 0, skipAmounts, 1, headerIndex);
                    skipAmounts[0] = headerSkipAmount;

                    log.info("保留表头文件 {} 已移到合并结果第一位", keepHeaderFileName);
                } else if (headerIndex == 0) {
                    log.info("保留表头文件 {} 已在第一位，无需重排", keepHeaderFileName);
                } else {
                    log.warn("未找到保留表头文件 {}，可能文件名不匹配", keepHeaderFileName);
                }
            }

            File outputFile = new File(uploadDir, "合并结果.xlsx");
            mergeGenericFiles(tempPaths, outputFile, gapRows, skipFlags, skipAmounts);

            // 清理临时文件
            for (String p : tempPaths) {
                try { new File(p).delete(); } catch (Exception ignored) {}
            }
            return outputFile.getAbsolutePath();
        } catch (Exception e) {
            for (String p : tempPaths) {
                try { new File(p).delete(); } catch (Exception ignored) {}
            }
            throw e;
        }
    }

    static class FileResult {
        String status;
        String filePath;
        String fileName;
        String tempPath;
        String message;
        String exception;
    }

    /**
     * 列移动规则
     */
    public static class ShiftRule {
        private String fileName;    // 目标文件名（原始文件名）
        private int colIndex;        // 列索引（0-based）
        private int shiftCount;      // 右移格数

        // Getter & Setter
        public String getFileName() { return fileName; }
        public void setFileName(String fileName) { this.fileName = fileName; }
        public int getColIndex() { return colIndex; }
        public void setColIndex(int colIndex) { this.colIndex = colIndex; }
        public int getShiftCount() { return shiftCount; }
        public void setShiftCount(int shiftCount) { this.shiftCount = shiftCount; }
    }

    private FileResult processSingleFile(File file, String targetColLetter, File tempDir, List<File> tempsToClean) {
        FileResult res = new FileResult();
        res.filePath = file.getAbsolutePath();
        res.fileName = file.getName();
        Workbook wb = null;

        if (targetColLetter == null || targetColLetter.trim().isEmpty()) {
            res.status = "error"; res.message = "未指定目标列";
            return res;
        }

        int targetColIdx = ExcelUtils.excelColumnToIndex(targetColLetter.trim().toUpperCase()) + 1; // 1-based

        try {
            wb = loadWorkbook(file);
            if (wb == null) {
                res.status = "error"; res.message = "无法读取文件（格式不支持）";
                return res;
            }

            Sheet sheet = wb.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();
            if (lastRow < 4) {
                closeQuietly(wb);
                res.status = "error"; res.message = "数据不足5行（实际" + (lastRow + 1) + "行）";
                return res;
            }

            // 步骤1: 第2行获取日期范围
            Row row2 = sheet.getRow(1);
            String dateStr = null;
            if (row2 != null) {
                for (int c = 0; c < row2.getLastCellNum(); c++) {
                    String val = ExcelUtils.getCellStringValue(row2.getCell(c));
                    if (val != null && val.contains("年") && val.contains("月") && val.contains("日")) {
                        dateStr = val; break;
                    }
                }
            }
            if (dateStr == null) {
                closeQuietly(wb);
                res.status = "skip"; res.message = "未在第2行找到日期信息";
                return res;
            }

            Integer month = extractMonth(dateStr);
            if (month == null) {
                closeQuietly(wb);
                res.status = "skip"; res.message = "无法解析月份，日期内容: " + dateStr;
                return res;
            }

            int fillMonth = (month == 1) ? 12 : month - 1;
            String monthStr = fillMonth + "月";

            // 步骤2: 第5行找"身份证件类型"列
            Row row5 = sheet.getRow(4);
            int idTypeCol = -1;
            if (row5 != null) {
                for (int c = 0; c < row5.getLastCellNum(); c++) {
                    String val = ExcelUtils.getCellStringValue(row5.getCell(c));
                    if ("身份证件类型".equals(val.trim())) { idTypeCol = c; break; }
                }
            }
            if (idTypeCol == -1) {
                closeQuietly(wb);
                res.status = "skip"; res.message = "未在第5行找到'身份证件类型'列";
                return res;
            }

            // 步骤3: 找第一个和最后一个"居民身份证"
            Integer startRow = null, endRow = null;
            for (int r = 5; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String val = ExcelUtils.getCellStringValue(row.getCell(idTypeCol));
                if ("居民身份证".equals(val.trim())) {
                    if (startRow == null) startRow = r;
                    endRow = r;
                }
            }
            if (startRow == null) {
                closeQuietly(wb);
                res.status = "skip"; res.message = "在'身份证件类型'列中未找到'居民身份证'";
                return res;
            }

            // 步骤4: 填写月份
            for (int r = startRow; r <= endRow; r++) {
                Row row = sheet.getRow(r);
                if (row == null) row = sheet.createRow(r);
                Cell cell = row.getCell(targetColIdx - 1);
                if (cell == null) cell = row.createCell(targetColIdx - 1);
                cell.setCellValue(monthStr);
            }

            // 保存处理后的文件
            String baseName = file.getName().replaceAll("(?i)\\.xlsx$", "").replaceAll("(?i)\\.xls$", "");
            File processed = new File(tempDir, baseName + "_processed.xlsx");
            try (FileOutputStream fos = new FileOutputStream(processed)) {
                wb.write(fos);
            }
            wb.close();

            res.status = "success";
            res.tempPath = processed.getAbsolutePath();
            res.message = "填写" + (endRow - startRow + 1) + "行，月份=" + monthStr;
            return res;
        } catch (Exception e) {
            closeQuietly(wb);
            res.status = "error";
            res.message = e.getMessage();
            // 记录完整异常信息以便前端展示
            StringWriter sw = new StringWriter();
            e.printStackTrace(new PrintWriter(sw));
            res.exception = sw.toString();
            log.error("处理文件失败: {} - {}", res.fileName, e.getMessage(), e);
            return res;
        }
    }

    /**
     * 加载Excel文件。先让 WorkbookFactory 自动检测格式；
     * 若失败且是 .xls，尝试直接用 HSSFWorkbook 读取。
     */
    private Workbook loadWorkbook(File file) throws Exception {
        String name = file.getName().toLowerCase();
        Exception firstErr = null;

        // 方式1: WorkbookFactory 自动检测（支持 .xls 和 .xlsx）
        try {
            return WorkbookFactory.create(file);
        } catch (Exception e) {
            firstErr = e;
            log.warn("WorkbookFactory 读取失败 [{}]: {}", file.getName(), e.getMessage());
        }

        // 方式2: 若扩展名为 .xls，尝试直接用 HSSFWorkbook（某些OLE2变体）
        if (name.endsWith(".xls")) {
            try (FileInputStream fis = new FileInputStream(file)) {
                return new HSSFWorkbook(fis);
            } catch (Exception e) {
                log.warn("HSSFWorkbook 直接读取失败 [{}]: {}", file.getName(), e.getMessage());
                throw new Exception("无法读取文件 '" + file.getName() + "'。WorkbookFactory: " +
                        (firstErr != null ? firstErr.getMessage() : "未知错误") +
                        "; HSSFWorkbook: " + e.getMessage() +
                        "。请确认文件是标准 Excel 格式（.xls 或 .xlsx），而非 HTML/XML 伪装格式。", firstErr);
            }
        }

        throw new Exception("无法读取文件 '" + file.getName() + "': " +
                (firstErr != null ? firstErr.getMessage() : "未知错误"), firstErr);
    }

    private void closeQuietly(Workbook wb) {
        if (wb != null) {
            try { wb.close(); } catch (Exception ignored) {}
        }
    }

    /**
     * 复制单元格值（含公式缓存结果）
     * 用于列移动场景：复制公式单元格时，同时写入缓存计算结果，
     * 确保打开文件时即使公式引用暂时异常也能显示正确的值。
     */
    private void copyCellValueWithCachedResult(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING:
                dst.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dst.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dst.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                // 复制公式字符串
                dst.setCellFormula(src.getCellFormula());
                // 同时写入缓存计算结果，确保值可见
                copyCachedFormulaResult(src, dst);
                break;
            case ERROR:
                dst.setCellErrorValue(src.getErrorCellValue());
                break;
            default:
                break;
        }
    }

    /**
     * 将公式单元格的缓存计算结果写入目标单元格
     * 这样即使公式引用在目标文件中无效，Excel 打开时仍能显示正确的值。
     */
    private void copyCachedFormulaResult(Cell src, Cell dst) {
        try {
            switch (src.getCachedFormulaResultType()) {
                case NUMERIC:
                    dst.setCellValue(src.getNumericCellValue());
                    break;
                case STRING:
                    dst.setCellValue(src.getStringCellValue());
                    break;
                case BOOLEAN:
                    dst.setCellValue(src.getBooleanCellValue());
                    break;
                case ERROR:
                    dst.setCellErrorValue(src.getErrorCellValue());
                    break;
                default:
                    break;
            }
        } catch (Exception e) {
            // 缓存结果获取失败，不影响主流程
        }
    }

    private void copyCellValue(Cell src, Cell dst) {
        switch (src.getCellType()) {
            case STRING: dst.setCellValue(src.getStringCellValue()); break;
            case NUMERIC: dst.setCellValue(src.getNumericCellValue()); break;
            case BOOLEAN: dst.setCellValue(src.getBooleanCellValue()); break;
            case FORMULA:
                dst.setCellFormula(src.getCellFormula());
                // 同时写入缓存计算结果，确保合并文件中值可见
                copyCachedFormulaResult(src, dst);
                break;
            case ERROR: dst.setCellErrorValue(src.getErrorCellValue()); break;
            default: break;
        }
    }

    /**
     * 手动复制单元格样式，双重缓存避免样式对象爆炸：
     * 1. localMap（IdentityHashMap）：同文件内按引用快速去重
     * 2. globalStyleCache（HashMap<指纹, CellStyle>）：跨文件按属性去重
     */
    private void copyCellStyle(Cell srcCell, Cell destCell, Workbook destWb,
                               Map<CellStyle, CellStyle> localMap,
                               Map<String, CellStyle> globalStyleCache,
                               Map<String, Font> globalFontCache) {
        CellStyle srcStyle = srcCell.getCellStyle();
        if (srcStyle == null) return;

        // 1. 同文件快速去重
        CellStyle cached = localMap.get(srcStyle);
        if (cached != null) {
            destCell.setCellStyle(cached);
            return;
        }

        // 2. 构建字体指纹 & 全局去重
        String fontKey = buildFontKey(srcCell, srcStyle);
        Font destFont = globalFontCache.get(fontKey);
        boolean fontIsNew = (destFont == null);

        // 3. 构建样式指纹 & 全局去重
        String styleKey = buildStyleKey(srcStyle, fontKey);
        CellStyle globalCached = globalStyleCache.get(styleKey);
        if (globalCached != null) {
            localMap.put(srcStyle, globalCached);
            destCell.setCellStyle(globalCached);
            return;
        }

        // 4. 新建样式
        CellStyle newStyle = destWb.createCellStyle();

        // 对齐
        newStyle.setAlignment(srcStyle.getAlignment());
        newStyle.setVerticalAlignment(srcStyle.getVerticalAlignment());
        newStyle.setWrapText(srcStyle.getWrapText());
        newStyle.setRotation(srcStyle.getRotation());
        newStyle.setIndention(srcStyle.getIndention());

        // 边框
        newStyle.setBorderTop(srcStyle.getBorderTop());
        newStyle.setBorderBottom(srcStyle.getBorderBottom());
        newStyle.setBorderLeft(srcStyle.getBorderLeft());
        newStyle.setBorderRight(srcStyle.getBorderRight());
        newStyle.setTopBorderColor(srcStyle.getTopBorderColor());
        newStyle.setBottomBorderColor(srcStyle.getBottomBorderColor());
        newStyle.setLeftBorderColor(srcStyle.getLeftBorderColor());
        newStyle.setRightBorderColor(srcStyle.getRightBorderColor());

        // 填充
        newStyle.setFillForegroundColor(srcStyle.getFillForegroundColor());
        newStyle.setFillPattern(srcStyle.getFillPattern());

        // 字体：全局去重
        if (fontIsNew) {
            destFont = copyFontToWorkbook(srcCell, srcStyle, destWb);
            if (destFont != null) {
                globalFontCache.put(fontKey, destFont);
            }
        }
        if (destFont != null) {
            newStyle.setFont(destFont);
        }

        // 数字格式
        try {
            newStyle.setDataFormat(destWb.createDataFormat().getFormat(
                    srcStyle.getDataFormatString()));
        } catch (Exception e) { /* ignore */ }

        // 隐藏 & 锁定
        newStyle.setHidden(srcStyle.getHidden());
        newStyle.setLocked(srcStyle.getLocked());

        // 双重缓存
        localMap.put(srcStyle, newStyle);
        globalStyleCache.put(styleKey, newStyle);
        destCell.setCellStyle(newStyle);
    }

    /** 构建字体属性指纹，用于全局字体去重 */
    private String buildFontKey(Cell srcCell, CellStyle srcStyle) {
        try {
            Font f = null;
            if (srcStyle instanceof org.apache.poi.xssf.usermodel.XSSFCellStyle) {
                f = ((org.apache.poi.xssf.usermodel.XSSFCellStyle) srcStyle).getFont();
            } else if (srcStyle instanceof org.apache.poi.hssf.usermodel.HSSFCellStyle) {
                f = ((org.apache.poi.hssf.usermodel.HSSFCellStyle) srcStyle)
                        .getFont(srcCell.getSheet().getWorkbook());
            }
            if (f != null) {
                return f.getFontName() + "|" + f.getFontHeightInPoints() + "|"
                        + f.getBold() + "|" + f.getItalic() + "|" + f.getColor()
                        + "|" + f.getUnderline() + "|" + f.getStrikeout();
            }
        } catch (Exception e) { log.warn("构建字体指纹失败: {}", e.getMessage()); }
        return "DEFAULT_FONT";
    }

    /** 构建样式属性指纹，用于全局样式去重 */
    private String buildStyleKey(CellStyle srcStyle, String fontKey) {
        return srcStyle.getAlignment() + "|" + srcStyle.getVerticalAlignment()
                + "|" + srcStyle.getBorderTop() + "|" + srcStyle.getBorderBottom()
                + "|" + srcStyle.getBorderLeft() + "|" + srcStyle.getBorderRight()
                + "|" + srcStyle.getTopBorderColor() + "|" + srcStyle.getBottomBorderColor()
                + "|" + srcStyle.getFillForegroundColor() + "|" + srcStyle.getFillPattern()
                + "|" + srcStyle.getDataFormatString()
                + "|" + srcStyle.getWrapText() + "|" + srcStyle.getHidden()
                + "|" + srcStyle.getLocked() + "|" + fontKey;
    }

    /** 跨工作簿复制字体，返回新建的 Font */
    private Font copyFontToWorkbook(Cell srcCell, CellStyle srcStyle, Workbook destWb) {
        try {
            Font srcFont = null;
            if (srcStyle instanceof org.apache.poi.xssf.usermodel.XSSFCellStyle) {
                srcFont = ((org.apache.poi.xssf.usermodel.XSSFCellStyle) srcStyle).getFont();
            } else if (srcStyle instanceof org.apache.poi.hssf.usermodel.HSSFCellStyle) {
                srcFont = ((org.apache.poi.hssf.usermodel.HSSFCellStyle) srcStyle)
                        .getFont(srcCell.getSheet().getWorkbook());
            }
            if (srcFont != null) {
                Font newFont = destWb.createFont();
                newFont.setBold(srcFont.getBold());
                newFont.setColor(srcFont.getColor());
                newFont.setFontHeight(srcFont.getFontHeight());
                newFont.setFontName(srcFont.getFontName());
                newFont.setItalic(srcFont.getItalic());
                newFont.setStrikeout(srcFont.getStrikeout());
                newFont.setTypeOffset(srcFont.getTypeOffset());
                newFont.setUnderline(srcFont.getUnderline());
                return newFont;
            }
        } catch (Exception e) {
            log.warn("复制字体失败: {}", e.getMessage());
        }
        return null;
    }

    private Integer extractMonth(String dateStr) {
        Matcher m = MONTH_PATTERN.matcher(dateStr);
        return m.find() ? Integer.parseInt(m.group(2)) : null;
    }

    private void mergeFiles(List<String> filePaths, File outputFile) throws Exception {
        try (Workbook newWb = new XSSFWorkbook()) {
            Sheet newSheet = newWb.createSheet("汇总表");
            int currentRow = 0;

            // 全局跨文件缓存：样式指纹 → CellStyle，字体指纹 → Font
            Map<String, CellStyle> globalStyleCache = new HashMap<>();
            Map<String, Font> globalFontCache = new HashMap<>();

            for (int idx = 0; idx < filePaths.size(); idx++) {
                String filePath = filePaths.get(idx);
                try (FileInputStream fis = new FileInputStream(filePath);
                     Workbook srcWb = WorkbookFactory.create(fis)) {

                    Sheet srcSheet = srcWb.getSheetAt(0);
                    int maxRow = srcSheet.getLastRowNum();

                    // 修正：遍历所有行取最大列数，而不是只看第一行
                    int maxCol = 0;
                    for (int r = 0; r <= maxRow; r++) {
                        Row row = srcSheet.getRow(r);
                        if (row != null && row.getLastCellNum() > maxCol) {
                            maxCol = row.getLastCellNum();
                        }
                    }

                    int fileStartRow = currentRow;

                    // 同文件快速去重 + 全局跨文件去重
                    Map<CellStyle, CellStyle> localMap = new IdentityHashMap<>();

                    for (int r = 0; r <= maxRow; r++) {
                        Row srcRow = srcSheet.getRow(r);
                        Row destRow = newSheet.createRow(currentRow);
                        if (srcRow != null) {
                            if (srcRow.getHeight() != -1) destRow.setHeight(srcRow.getHeight());
                            for (int c = 0; c < maxCol; c++) {
                                Cell srcCell = srcRow.getCell(c);
                                if (srcCell == null) continue;
                                Cell destCell = destRow.createCell(c);
                                copyCellValue(srcCell, destCell);
                                copyCellStyle(srcCell, destCell, newWb,
                                        localMap, globalStyleCache, globalFontCache);
                            }
                        }
                        currentRow++;
                    }

                    // 复制合并单元格（带行偏移）
                    for (int mi = 0; mi < srcSheet.getNumMergedRegions(); mi++) {
                        CellRangeAddress merged = srcSheet.getMergedRegion(mi);
                        CellRangeAddress newRegion = new CellRangeAddress(
                                merged.getFirstRow() + fileStartRow,
                                merged.getLastRow() + fileStartRow,
                                merged.getFirstColumn(), merged.getLastColumn()
                        );
                        try {
                            newSheet.addMergedRegion(newRegion);
                        } catch (Exception e) {
                            log.warn("合并单元格复制失败 [{} 行{}-{}]: {}",
                                    filePath, merged.getFirstRow(), merged.getLastRow(), e.getMessage());
                        }
                    }

                    // 复制列宽
                    for (int c = 0; c < maxCol; c++) {
                        int width = srcSheet.getColumnWidth(c);
                        if (width > 0 && newSheet.getColumnWidth(c) < width) {
                            newSheet.setColumnWidth(c, width);
                        }
                    }
                }

                if (idx < filePaths.size() - 1) currentRow += 20;
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                newWb.write(fos);
            }
        }
    }

    /**
     * 将指定文件的某列数据整体向右移动N格
     * 表头也被移动，原位置（含表头）留空
     *
     * <p>注意：必须先读入内存再写回，不能同时对同一文件进行读写，
     * 否则 POI 内部 ZLIB 流会冲突导致 EOFException。
     *
     * @param file       目标Excel文件
     * @param colIndex  要移动的列索引（0-based）
     * @param shiftCount 向右移动的格数
     */
    private void shiftColumnRight(File file, int colIndex, int shiftCount) throws Exception {
        // Step 1: 先将整个文件读入字节数组，释放文件句柄
        byte[] fileBytes;
        try (FileInputStream fis = new FileInputStream(file)) {
            fileBytes = fis.readAllBytes();
        }

        // Step 2: 从内存中的字节数组创建 Workbook（不再持有文件输入流）
        File tempOutputFile = new File(file.getAbsolutePath() + ".shift_tmp");
        try (InputStream is = new ByteArrayInputStream(fileBytes);
             Workbook wb = WorkbookFactory.create(is)) {

            Sheet sheet = wb.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();

            // ✅ 样式缓存：在方法级别创建一次，所有单元格共享，避免 createCellStyle() 爆炸超过 64000 上限
            Map<CellStyle, CellStyle> localStyleMap = new IdentityHashMap<>();
            Map<String, CellStyle> globalStyleCache = new HashMap<>();
            Map<String, Font> globalFontCache = new HashMap<>();

            // 对每一行，从该行列索引最大值开始向左遍历到 colIndex
            // 将每一列的数据向右复制 shiftCount 格（从右向左处理，避免覆盖）
            for (int r = lastRow; r >= 0; r--) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // getLastCellNum() 返回 1-based 的最大值
                int lastCellNum = row.getLastCellNum();
                if (lastCellNum <= colIndex) continue; // 该行没有 colIndex 及右侧的数据

                int maxCol = lastCellNum - 1; // 转为 0-based 最大列索引

                // 从最右列向左遍历到 colIndex，依次右移 shiftCount 格
                for (int c = maxCol; c >= colIndex; c--) {
                    Cell srcCell = row.getCell(c);
                    if (srcCell == null) continue;

                    int targetColIdx = c + shiftCount;
                    Cell destCell = row.createCell(targetColIdx);

                    // 复制值（含公式的结果值）
                    copyCellValueWithCachedResult(srcCell, destCell);

                    // 复制样式（使用方法级缓存，避免重复创建样式对象）
                    copyCellStyle(srcCell, destCell, wb, localStyleMap, globalStyleCache, globalFontCache);

                    // 清空原位置（移除原单元格，留空）
                    row.removeCell(srcCell);
                }
            }

            // Step 3: 写入临时输出文件（不写回原文件！）
            try (FileOutputStream fos = new FileOutputStream(tempOutputFile)) {
                wb.write(fos);
            }
        }

        // Step 4: 用临时文件替换原文件
        if (file.exists() && !file.delete()) {
            throw new IOException("无法删除原文件: " + file.getAbsolutePath());
        }
        if (!tempOutputFile.renameTo(file)) {
            throw new IOException("无法重命名临时文件: " + tempOutputFile.getAbsolutePath());
        }

        log.info("列移动完成：文件={}, 列={}, 右移{}格", file.getName(),
                colIndexToLetter(colIndex), shiftCount);
    }

    /**
     * 对同一文件执行多条列移动规则（只解析一次 xlsx，大幅提升性能）
     * <p>对同一文件的多个列移动规则，合并为一次文件读写操作。
     * 所有列移动在同一 Workbook 内存对象上顺序执行，最后一次性写回。
     *
     * @param file  目标Excel文件（临时文件）
     * @param rules 该文件的所有列移动规则（已按用户设置顺序排序）
     */
    private void shiftMultipleColumns(File file, List<ShiftRule> rules) throws Exception {
        if (rules == null || rules.isEmpty()) return;

        // Step1: 将整个文件读入字节数组，释放文件句柄
        byte[] fileBytes;
        try (FileInputStream fis = new FileInputStream(file)) {
            fileBytes = fis.readAllBytes();
        }

        // Step2: 从内存中的字节数组创建 Workbook（不再持有文件输入流）
        File tempOutputFile = new File(file.getAbsolutePath() + ".shift_tmp");
        try (InputStream is = new ByteArrayInputStream(fileBytes);
             Workbook wb = WorkbookFactory.create(is)) {

            Sheet sheet = wb.getSheetAt(0);
            int lastRow = sheet.getLastRowNum();

            // ✅ 样式缓存：在方法级别创建一次，所有规则/所有单元格共享
            // 同一文件可能执行多条规则（如先移K列再移A列），必须复用同一套缓存
            Map<CellStyle, CellStyle> localStyleMap = new IdentityHashMap<>();
            Map<String, CellStyle> globalStyleCache = new HashMap<>();
            Map<String, Font> globalFontCache = new HashMap<>();

            // Step3: 按规则顺序，在同一 Workbook 上逐一移动列
            // 重要：必须先按 colIndex 降序排序（最右列先处理），
            // 否则右列规则处理完后，左列规则的 colIndex 会指向错误的列
            rules.sort((a, b) -> Integer.compare(b.getColIndex(), a.getColIndex()));

            for (ShiftRule rule : rules) {
                int colIndex = rule.getColIndex();
                int shiftCount = rule.getShiftCount();

                log.info("  → 列 {} 右移 {} 格", colIndexToLetter(colIndex), shiftCount);

                // 对每一行，从该行列索引最大值开始向左遍历到 colIndex
                for (int r = lastRow; r >= 0; r--) {
                    Row row = sheet.getRow(r);
                    if (row == null) continue;

                    int lastCellNum = row.getLastCellNum();
                    if (lastCellNum <= colIndex) continue;

                    int maxCol = lastCellNum - 1;

                    for (int c = maxCol; c >= colIndex; c--) {
                        Cell srcCell = row.getCell(c);
                        if (srcCell == null) continue;

                        int targetColIdx = c + shiftCount;
                        Cell destCell = row.createCell(targetColIdx);

                        // 复制值（含公式缓存结果）
                        copyCellValueWithCachedResult(srcCell, destCell);

                        // 复制样式（使用方法级共享缓存）
                        copyCellStyle(srcCell, destCell, wb, localStyleMap, globalStyleCache, globalFontCache);

                        row.removeCell(srcCell);
                    }
                }
            }

            // Step4: 所有规则执行完毕后，一次性写回临时文件
            try (FileOutputStream fos = new FileOutputStream(tempOutputFile)) {
                wb.write(fos);
            }
        }

        // Step5: 用临时文件替换原文件
        if (file.exists() && !file.delete()) {
            throw new IOException("无法删除原文件: " + file.getAbsolutePath());
        }
        if (!tempOutputFile.renameTo(file)) {
            throw new IOException("无法重命名临时文件: " + tempOutputFile.getAbsolutePath());
        }

        log.info("多规则列移动完成：文件={}, 共执行 {} 条规则（单次IO）",
                file.getName(), rules.size());
    }

    /**
     * 列索引转字母（0=A, 1=B, ...）
     */
    private static String colIndexToLetter(int colIndex) {
        StringBuilder sb = new StringBuilder();
        int idx = colIndex;
        do {
            sb.insert(0, (char) ('A' + idx % 26));
            idx = idx / 26 - 1;
        } while (idx >= 0);
        return sb.toString();
    }

    private void deleteDirectory(File dir) {
        if (dir == null || !dir.exists()) return;
        File[] files = dir.listFiles();
        if (files != null) {
            for (File f : files) {
                if (f.isDirectory()) deleteDirectory(f);
                else f.delete();
            }
        }
        dir.delete();
    }

    /** 精确检测实际数据行数（跳过底部带格式空行），对齐 Python _get_actual_max_row */
    private int getActualMaxRow(Sheet sheet) {
        int maxRow = sheet.getLastRowNum();
        int maxCol = 0;
        for (int r = 0; r <= maxRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > maxCol) maxCol = row.getLastCellNum();
        }
        if (maxCol == 0) return 0;
        for (int r = maxRow; r >= 0; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < maxCol; c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() != CellType.BLANK) return r;
            }
        }
        return 0;
    }

    /**
     * P0: 通用合并 — 支持 gapRows / header skip / 行高复制 / 精确行检测
     */
    private void mergeGenericFiles(List<String> filePaths, File outputFile,
                                    int gapRows, boolean[] skipFlags, int[] skipAmounts) throws Exception {
        if (filePaths.isEmpty()) return;

        try (Workbook newWb = new XSSFWorkbook()) {
            Sheet newSheet = newWb.createSheet("合并结果");
            int currentRow = 0;
            Map<String, CellStyle> globalStyleCache = new HashMap<>();
            Map<String, Font> globalFontCache = new HashMap<>();

            for (int idx = 0; idx < filePaths.size(); idx++) {
                String filePath = filePaths.get(idx);
                boolean shouldSkip = skipFlags != null && idx < skipFlags.length && skipFlags[idx];
                int skipHdr = shouldSkip ? (idx < skipAmounts.length ? skipAmounts[idx] : 0) : 0;

                try (FileInputStream fis = new FileInputStream(filePath);
                     Workbook srcWb = WorkbookFactory.create(fis)) {

                    Sheet srcSheet = srcWb.getSheetAt(0);
                    int maxRow = getActualMaxRow(srcSheet);
                    // Scan all rows for max column
                    int maxCol = 0;
                    for (int r = 0; r < srcSheet.getLastRowNum(); r++) {
                        Row row = srcSheet.getRow(r);
                        if (row != null && row.getLastCellNum() > maxCol) maxCol = row.getLastCellNum();
                    }
                    if (maxRow == 0 || maxCol == 0) continue;
                    if (skipHdr >= maxRow + 1) continue;

                    int rowStart = skipHdr; // 0-based, actual data start
                    int fileStartRow = currentRow;
                    Map<CellStyle, CellStyle> localMap = new IdentityHashMap<>();

                    for (int r = rowStart; r <= maxRow; r++) {
                        Row srcRow = srcSheet.getRow(r);
                        Row destRow = newSheet.createRow(currentRow);
                        if (srcRow != null) {
                            // 行高复制（P1）
                            if (srcRow.getHeight() != -1) destRow.setHeight(srcRow.getHeight());
                            for (int c = 0; c < maxCol; c++) {
                                Cell srcCell = srcRow.getCell(c);
                                if (srcCell == null) continue;
                                Cell destCell = destRow.createCell(c);
                                copyCellValue(srcCell, destCell);
                                copyCellStyle(srcCell, destCell, newWb, localMap, globalStyleCache, globalFontCache);
                            }
                        }
                        currentRow++;
                    }

                    // 复制合并单元格（偏移行号，跳过已skip的行）
                    for (int mi = 0; mi < srcSheet.getNumMergedRegions(); mi++) {
                        CellRangeAddress merged = srcSheet.getMergedRegion(mi);
                        if (merged.getFirstRow() < rowStart || merged.getLastRow() < rowStart) continue;
                        int adjFirstRow = Math.max(merged.getFirstRow(), rowStart);
                        int adjLastRow = Math.min(merged.getLastRow(), maxRow);
                        CellRangeAddress newRegion = new CellRangeAddress(
                                adjFirstRow - rowStart + fileStartRow,
                                adjLastRow - rowStart + fileStartRow,
                                merged.getFirstColumn(), merged.getLastColumn());
                        try { newSheet.addMergedRegion(newRegion); } catch (Exception ignored) {}
                    }

                    // 复制列宽
                    for (int c = 0; c < maxCol; c++) {
                        int width = srcSheet.getColumnWidth(c);
                        if (width > 0 && newSheet.getColumnWidth(c) < width) {
                            newSheet.setColumnWidth(c, width);
                        }
                    }
                }

                if (idx < filePaths.size() - 1) currentRow += gapRows;
            }

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                newWb.write(fos);
            }
        }
    }
}
