package org.example;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.Duration;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.example.ExcelUtil.*;

public class ExcelHandler {

    private Logger logger = LoggerFactory.getLogger(this.getClass());

    private List<String> handlerSheetNames = new ArrayList<>();

    public void handlerExcelFiles(File[] excelFiles, int chuckSize, int parseType, String cellAddress) throws IOException {
        for (File excelFile : excelFiles) {
            handlerExcelFile(excelFile, chuckSize, parseType, cellAddress);
        }
    }

    public void handlerExcelFile(File srcExcelFile, int chuckSize, int parseType, String cellAddress) throws IOException {
        logger.info("开始处理"+srcExcelFile.getName()+"文件");
        Instant start = Instant.now();
        try(Workbook workbook = WorkbookFactory.create(srcExcelFile);
            FileOutputStream out = new FileOutputStream("TEMP.xlsx");
            SXSSFWorkbook copyWorkbook = new SXSSFWorkbook(100)
        ){
            // 解析类型（1 COC01~XX，2 指定单元格）
            for (Sheet sheet : workbook) {
                int startRowIdx = sheet.getFirstRowNum();
                int endRowIdx = sheet.getLastRowNum();
                logger.info(String.format("起点:%s, 结束:%s", startRowIdx, endRowIdx));
                // 第一个sheet
                if (workbook.getSheetIndex(sheet) == 0){
                    // 遍历合并区域
                    List<List<Integer>> groupList = groupSheetRowsBySize(sheet, chuckSize);
                    List<Set<CellRangeAddress>> groupCellRangeAddressList = groupSheetMergedRegions(sheet, chuckSize);

                    for (int i = 0; i < groupCellRangeAddressList.size(); i++) {
                        // 创建Sheet
                        String sheetName = null;
                        if (parseType == 1) {
                            sheetName = String.format("COC(%d)", i + 1);
                        }else if (parseType == 2) {
                            // 创建Sheet
                            CellAddress cd = new CellAddress(cellAddress);
                            Cell keyCell = sheet.getRow(cd.getRow() + chuckSize * i).getCell(cd.getColumn());
                            sheetName = keyCell.getStringCellValue();
                        }
                        handlerSheetNames.add(sheetName);
                        Sheet copySheet = copyWorkbook.createSheet(sheetName);
                        copySheet( sheet, copySheet, groupCellRangeAddressList.get(i), groupList.get(i).stream().mapToInt(Integer::intValue).toArray());
                        copyPicture(sheet, copySheet, i);
                        adjustPicLoc(copySheet);
                    }
                }else {
                    Sheet targetSheet = copyWorkbook.createSheet(sheet.getSheetName());
                    copySheet(sheet, targetSheet, sheet.getMergedRegions(), IntStream.rangeClosed(sheet.getFirstRowNum(), sheet.getLastRowNum()).toArray());
                    copyPicture(sheet, targetSheet);
                    adjustPicLoc(targetSheet);
                }
            }
            copyWorkbook.write(out);

        }
        ZipSecureFile.setMinInflateRatio(-1.0d);
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("YYYY年MM月dd日HH时mm分ss秒");
        try(Workbook workbook = WorkbookFactory.create(new File("TEMP.xlsx"));
            FileOutputStream fos = new FileOutputStream(FilenameUtils.getBaseName(srcExcelFile.getName())+"_" + LocalDateTime.now().format(dateTimeFormatter) +"."+FilenameUtils.getExtension(srcExcelFile.getName()))
        ){
            List<Sheet> sheets = handlerSheetNames.stream().map(workbook::getSheet).collect(Collectors.toList());
            formatSheets(sheets);
            workbook.write(fos);
        }
        Files.deleteIfExists(Paths.get("TEMP.xlsx"));
        Instant end = Instant.now();
        logger.info("处理"+srcExcelFile.getName()+"文件成功,耗时:"+ Duration.between(start, end).toMillis()+"毫秒");
    }

    // 调整图片位置
    private void adjustPicLoc(Sheet sheet){
        // 移动图片
        Drawing<Shape> drawingPatriarch = (Drawing<Shape>) sheet.getDrawingPatriarch();

        for (Shape shape : drawingPatriarch) {
            if (shape instanceof Picture){
                ClientAnchor clientAnchor = ((Picture) shape).getClientAnchor();
                clientAnchor.setAnchorType(ClientAnchor.AnchorType.MOVE_DONT_RESIZE);
                clientAnchor.setRow1(clientAnchor.getRow1() - sheet.getFirstRowNum());
                clientAnchor.setRow2(clientAnchor.getRow2() - sheet.getFirstRowNum());
            }
        }
    }

    private void formatSheets(List<Sheet> sheets){

        logger.info("开始格式化Workbook");
        for (Sheet sheet : sheets) {
            sheet.shiftRows(sheet.getFirstRowNum(), sheet.getLastRowNum(), - sheet.getFirstRowNum() + 1);
            sheet.shiftRows(sheet.getFirstRowNum(), sheet.getLastRowNum(), -1);
        }
        logger.info("格式化Workbook完成");
    }
}
