package com.example.project.service.finance.jxkh;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.*;

/**
 * 绩效考核 - 数据匹配（cw_sjpp_zhf.py）
 * 从原始表自动生成完整版表格，设置 VLOOKUP 公式链
 */
@Service
public class JxkhSjppService {

    public String generate(MultipartFile sourceFile, MultipartFile prevFile, String uploadDir) throws Exception {
        File dir = new File(uploadDir);
        if (!dir.exists()) dir.mkdirs();

        File tempSource = new File(dir, UUID.randomUUID() + "_" + sourceFile.getOriginalFilename());
        sourceFile.transferTo(tempSource);

        File tempPrev = null;
        if (prevFile != null && !prevFile.isEmpty()) {
            tempPrev = new File(dir, UUID.randomUUID() + "_" + prevFile.getOriginalFilename());
            prevFile.transferTo(tempPrev);
        }

        try {
            return processWorkbook(tempSource, tempPrev, uploadDir);
        } finally {
            if (tempSource.exists()) tempSource.delete();
            if (tempPrev != null && tempPrev.exists()) tempPrev.delete();
        }
    }

    private String processWorkbook(File sourceFile, File prevFile, String uploadDir) throws Exception {
        try (Workbook wb = WorkbookFactory.create(sourceFile)) {
            String[] sheetNames = new String[wb.getNumberOfSheets()];
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                sheetNames[i] = wb.getSheetName(i);
            }

            String sheetPm = "运营中心-项目经理";
            String sheetDaigong = findSheetByPrefix(sheetNames, "代工费-");
            String sheetBaishi = findSheetByPrefix(sheetNames, "百事通报备名单-");
            String sheetEmp = findSheetByPrefix(sheetNames, "员工信息汇总-");

            if (sheetDaigong == null) throw new Exception("未找到以'代工费-'开头的工作表");
            if (sheetBaishi == null) throw new Exception("未找到以'百事通报备名单-'开头的工作表");
            if (sheetEmp == null) throw new Exception("未找到以'员工信息汇总-'开头的工作表");

            Sheet wsEmp = wb.getSheet(sheetEmp);
            Sheet wsBs = wb.getSheet(sheetBaishi);
            Sheet wsDg = wb.getSheet(sheetDaigong);

            processEmployeeSheet(wsEmp, sheetBaishi);
            processBaishiSheet(wsBs, sheetDaigong, sheetPm, sheetEmp, prevFile, sheetBaishi);
            processDaigongSheet(wsDg, sheetBaishi);

            // POI 不支持设置 calcMode，公式在 Excel 打开时会自动重算

            String baseName = sourceFile.getName().replaceAll("\\.xlsx$", "").replaceAll("\\.xls$", "");
            File outFile = new File(uploadDir, baseName + "_自动生成（已核对）.xlsx");
            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                wb.write(fos);
            }
            return outFile.getAbsolutePath();
        }
    }

    private String findSheetByPrefix(String[] names, String prefix) {
        for (String name : names) {
            if (name.startsWith(prefix)) return name;
        }
        return null;
    }

    private int findLastUsedRow(Sheet sheet, int maxCol) {
        for (int r = sheet.getLastRowNum(); r >= 0; r--) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < maxCol && c <= row.getLastCellNum(); c++) {
                Cell cell = row.getCell(c);
                if (cell != null && cell.getCellType() != CellType.BLANK) {
                    return r;
                }
            }
        }
        return 0;
    }

    private void copyCellStyle(Sheet sheet, int srcRow, int srcCol, int dstRow, int dstCol) {
        Row sRow = sheet.getRow(srcRow);
        Row dRow = sheet.getRow(dstRow);
        if (sRow == null || dRow == null) return;
        Cell sCell = sRow.getCell(srcCol);
        Cell dCell = dRow.getCell(dstCol);
        if (sCell == null || dCell == null) return;
        // 在 POI 中复制样式需要创建新样式
        CellStyle newStyle = sheet.getWorkbook().createCellStyle();
        newStyle.cloneStyleFrom(sCell.getCellStyle());
        dCell.setCellStyle(newStyle);
    }

    private void ensureColumnWidth(Sheet sheet, String colLetter, int width) {
        int colIdx = colLetterToIndex(colLetter);
        sheet.setColumnWidth(colIdx, width * 256);
    }

    private int colLetterToIndex(String col) {
        int result = 0;
        for (char c : col.toUpperCase().toCharArray()) {
            result = result * 26 + (c - 'A' + 1);
        }
        return result - 1; // 0-based
    }

    private String quoteSheetName(String name) {
        return "'" + name.replace("'", "''") + "'";
    }

    private String buildExternalSheetRef(File file, String sheetName) {
        String absPath = file.getAbsolutePath().replace("/", "\\");
        String folder = new File(absPath).getParent();
        String fileName = new File(absPath).getName();
        String extRef = folder + "\\[" + fileName + "]" + sheetName;
        extRef = extRef.replace("'", "''");
        return "'" + extRef + "'";
    }

    // ========== 员工信息汇总表处理 ==========
    private void processEmployeeSheet(Sheet wsEmp, String baishiSheetName) {
        int lastRow = findLastUsedRow(wsEmp, 40);
        String[] cols = {"AA", "AB", "AC", "AD", "AE", "AF"};
        for (String col : cols) {
            int colIdx = colLetterToIndex(col);
            copyCellStyle(wsEmp, 0, 25, 0, colIdx);
            if (lastRow >= 1) {
                copyCellStyle(wsEmp, 1, 25, 1, colIdx);
            }
        }

        setCellValue(wsEmp, 0, colLetterToIndex("AA"), "姓名");
        setCellValue(wsEmp, 0, colLetterToIndex("AB"), "项目+姓名");
        setCellValue(wsEmp, 0, colLetterToIndex("AC"), "工厂实际结算");
        setCellValue(wsEmp, 0, colLetterToIndex("AD"), "项目驻场");
        setCellValue(wsEmp, 0, colLetterToIndex("AE"), "项目经理");
        setCellValue(wsEmp, 0, colLetterToIndex("AF"), "总人数（上月延续）");

        ensureColumnWidth(wsEmp, "AA", 13);
        ensureColumnWidth(wsEmp, "AB", 24);
        ensureColumnWidth(wsEmp, "AC", 24);
        ensureColumnWidth(wsEmp, "AD", 13);
        ensureColumnWidth(wsEmp, "AE", 13);
        ensureColumnWidth(wsEmp, "AF", 18);

        String qBs = quoteSheetName(baishiSheetName);
        for (int r = 1; r <= lastRow; r++) {
            int excelRow = r + 1; // 1-based for formula
            setFormula(wsEmp, r, colLetterToIndex("AA"), "=C" + excelRow);
            setFormula(wsEmp, r, colLetterToIndex("AB"), "=F" + excelRow + "&AA" + excelRow);
            setFormula(wsEmp, r, colLetterToIndex("AC"), "=VLOOKUP(AB" + excelRow + "," + qBs + "!$AB:$AE,1,0)");
            setFormula(wsEmp, r, colLetterToIndex("AD"), "=VLOOKUP(AB" + excelRow + "," + qBs + "!$AC:$AE,2,0)");
            setFormula(wsEmp, r, colLetterToIndex("AE"), "=VLOOKUP(AB" + excelRow + "," + qBs + "!$AC:$AE,3,0)");
            setFormula(wsEmp, r, colLetterToIndex("AF"), "=VLOOKUP(AB" + excelRow + "," + qBs + "!$AC:$AE,1,0)");
        }
    }

    // ========== 百事通报备名单表处理 ==========
    private void processBaishiSheet(Sheet wsBs, String daigongSheetName, String pmSheetName,
                                     String empSheetName, File prevFile, String prevBaishiName) {
        int lastRow = findLastUsedRow(wsBs, 35);
        String[] cols = {"Z", "AA", "AB", "AC", "AD", "AE", "AG"};
        for (String col : cols) {
            int colIdx = colLetterToIndex(col);
            copyCellStyle(wsBs, 0, 22, 0, colIdx);
            if (lastRow >= 1) {
                copyCellStyle(wsBs, 1, 22, 1, colIdx);
            }
        }

        setCellValue(wsBs, 0, colLetterToIndex("Z"), "姓名");
        setCellValue(wsBs, 0, colLetterToIndex("AA"), "企业+姓名");
        setCellValue(wsBs, 0, colLetterToIndex("AB"), "工厂实际结算");
        setCellValue(wsBs, 0, colLetterToIndex("AC"), "总人数（上月延续）");
        setCellValue(wsBs, 0, colLetterToIndex("AD"), "项目驻场");
        setCellValue(wsBs, 0, colLetterToIndex("AE"), "项目经理");
        setCellValue(wsBs, 0, colLetterToIndex("AG"), "员工信息匹配");

        ensureColumnWidth(wsBs, "Z", 12);
        ensureColumnWidth(wsBs, "AA", 24);
        ensureColumnWidth(wsBs, "AB", 24);
        ensureColumnWidth(wsBs, "AC", 24);
        ensureColumnWidth(wsBs, "AD", 12);
        ensureColumnWidth(wsBs, "AE", 12);
        ensureColumnWidth(wsBs, "AG", 24);

        String qDg = quoteSheetName(daigongSheetName);
        String qPm = quoteSheetName(pmSheetName);
        String qEmp = quoteSheetName(empSheetName);

        for (int r = 1; r <= lastRow; r++) {
            int excelRow = r + 1;
            setFormula(wsBs, r, colLetterToIndex("Z"), "=C" + excelRow);
            setFormula(wsBs, r, colLetterToIndex("AA"), "=K" + excelRow + "&Z" + excelRow);
            setFormula(wsBs, r, colLetterToIndex("AB"), "=VLOOKUP(AA" + excelRow + "," + qDg + "!$K:$K,1,0)");

            if (prevFile != null && prevBaishiName != null) {
                String extRef = buildExternalSheetRef(prevFile, prevBaishiName);
                setFormula(wsBs, r, colLetterToIndex("AC"), "=VLOOKUP(AA" + excelRow + "," + extRef + "!$AC:$AC,1,0)");
            } else {
                setCellValue(wsBs, r, colLetterToIndex("AC"), "");
            }

            setFormula(wsBs, r, colLetterToIndex("AD"), "=R" + excelRow);
            setFormula(wsBs, r, colLetterToIndex("AE"), "=VLOOKUP(K" + excelRow + "," + qPm + "!$B:$C,2,0)");
            setFormula(wsBs, r, colLetterToIndex("AG"), "=VLOOKUP(AB" + excelRow + "," + qEmp + "!$AB:$AB,1,0)");
        }
    }

    // ========== 代工费表处理 ==========
    private void processDaigongSheet(Sheet wsDg, String baishiSheetName) {
        int lastRow = findLastUsedRow(wsDg, 20);
        for (String col : new String[]{"K", "L"}) {
            int colIdx = colLetterToIndex(col);
            copyCellStyle(wsDg, 0, 9, 0, colIdx);
            if (lastRow >= 1) {
                copyCellStyle(wsDg, 1, 9, 1, colIdx);
            }
        }

        setCellValue(wsDg, 1, colLetterToIndex("K"), "项目名称+姓名");
        setCellValue(wsDg, 1, colLetterToIndex("L"), "是否在报备名单");

        ensureColumnWidth(wsDg, "K", 24);
        ensureColumnWidth(wsDg, "L", 24);

        String qBs = quoteSheetName(baishiSheetName);
        for (int r = 2; r <= lastRow; r++) {
            int excelRow = r + 1;
            setFormula(wsDg, r, colLetterToIndex("K"), "=D" + excelRow + "&E" + excelRow);
            setFormula(wsDg, r, colLetterToIndex("L"), "=VLOOKUP(K" + excelRow + "," + qBs + "!$AB:$AB,1,0)");
        }
    }

    private void setCellValue(Sheet sheet, int rowIdx, int colIdx, String value) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) row = sheet.createRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        if (cell == null) cell = row.createCell(colIdx);
        cell.setCellValue(value);
    }

    private void setFormula(Sheet sheet, int rowIdx, int colIdx, String formula) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) row = sheet.createRow(rowIdx);
        Cell cell = row.getCell(colIdx);
        if (cell == null) cell = row.createCell(colIdx);
        cell.setCellFormula(formula);
    }
}
