package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public final class ExcelUtil {

    private ExcelUtil(){
    }

    private static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    public static void copyRow(Row srcRow, Row targetRow){
        logger.info("开始copyRow");

        for (Cell cell : srcRow) {
            Cell targetCell = targetRow.createCell(cell.getColumnIndex());
            copyCell(cell, targetCell);
        }
        // 调整列宽
        for (int i = 0; i < srcRow.getPhysicalNumberOfCells(); i++){
            targetRow.getSheet().setColumnWidth(i, srcRow.getSheet().getColumnWidth(i));
        }
        // 调整行高
        targetRow.setHeight(srcRow.getHeight());
        logger.info("结束copyRow");
    }

    public static void copyCell(Cell srcCell, Cell targetCell){
        // 复制评论
        copyCellComment(srcCell, targetCell);
        // 复制内容
        copyCellValue(srcCell, targetCell);
        // 复制样式
        copyCellStyle(srcCell, targetCell);
    }

    public static void copyCellComment(Cell srcCell, Cell targetCell){
        logger.info("开始copyCellComment");
        if (Objects.nonNull(srcCell.getCellComment())){
            targetCell.setCellComment(srcCell.getCellComment());
        }
        logger.info("结束copyCellComment");
    }


    public static void copyCellValue(Cell srcCell, Cell targetCell){
        logger.info("开始copyCellValue");
        switch (srcCell.getCellType()){
            case STRING:
                targetCell.setCellValue(srcCell.getStringCellValue());break;
            case NUMERIC:
                targetCell.setCellValue(srcCell.getNumericCellValue());break;
            case FORMULA:
                targetCell.setCellValue(srcCell.getCellFormula());break;
            case BOOLEAN:
                targetCell.setCellValue(srcCell.getBooleanCellValue());break;
            case ERROR:
                targetCell.setCellValue(srcCell.getErrorCellValue());break;
            case BLANK:
            default:
                targetCell.setBlank();
                break;
        }
        logger.info("结束copyCellValue");
    }



    public static void copyCellStyle(Cell srcCell, Cell targetCell){
        logger.info("开始copyCellStyle");
        CellStyle cellStyle = targetCell.getSheet().getWorkbook().createCellStyle();
        cellStyle.cloneStyleFrom(srcCell.getCellStyle());
        targetCell.setCellStyle(cellStyle);
        logger.info("结束copyCellStyle");
    }

    public static List<List<Integer>> groupSheetRowsBySize(Sheet sheet, int chuckSize){
        int startRowIdx = sheet.getFirstRowNum();
        int endRowIdx = sheet.getLastRowNum();
        final AtomicInteger counter = new AtomicInteger();
        return new ArrayList<>(IntStream.range(startRowIdx, endRowIdx).boxed()
                .collect(Collectors.groupingBy(it -> counter.getAndIncrement() / chuckSize))
                .values());
    }

    public static List<Set<CellRangeAddress>> groupSheetMergedRegions(Sheet sheet, int chuckSize){

        List<Set<CellRangeAddress>> rtn = new ArrayList<>();
        List<List<Integer>> groupList = groupSheetRowsBySize(sheet, chuckSize);
        List<CellRangeAddress> cellRangeAddressList = sheet.getMergedRegions();

        for (int i = 0; i < groupList.size(); i++) {

            Set<CellRangeAddress> cellRangeAddressSet = new HashSet<>();
            for (int j = 0; j < groupList.get(i).size(); j++){
                for (Cell cell : sheet.getRow(groupList.get(i).get(j))) {
                    for (CellRangeAddress cellRangeAddress : cellRangeAddressList) {
                        if (cellRangeAddress.isInRange(cell)) {
                            cellRangeAddressSet.add(cellRangeAddress);
                        }
                    }
                }
            }
            rtn.add(cellRangeAddressSet);
        }
        return rtn;
    }

    public static void copyPicture(Sheet srcSheet, Sheet targetSheet){
        logger.info("开始copyPicture");
        // 获取图片位置
        Drawing<Shape> drawingPatriarch = (Drawing<Shape>) srcSheet.getDrawingPatriarch();
        Drawing<Shape> targetDrawingPatriarch = (Drawing<Shape>) targetSheet.createDrawingPatriarch();

        if (Objects.nonNull(drawingPatriarch)){
            for (Shape shape : drawingPatriarch) {
                if (shape instanceof Picture){

                    ClientAnchor clientAnchor = ((Picture) shape).getClientAnchor();
                    PictureData pictureData = ((Picture) shape).getPictureData();

                    targetDrawingPatriarch.createPicture(clientAnchor,
                            targetSheet.getWorkbook().addPicture(pictureData.getData(), pictureData.getPictureType()));

                }
            }
            logger.info("结束copyPicture");
        }
    }

    public static void copyPicture(Sheet srcSheet, Sheet targetSheet, int index){
        logger.info("开始copyPicture");
        // 获取图片位置

        Drawing<Shape> drawingPatriarch = (Drawing<Shape>) srcSheet.getDrawingPatriarch();
        Drawing<Shape> targetDrawingPatriarch = (Drawing<Shape>) targetSheet.createDrawingPatriarch();

        int i = 0;
        for (Shape shape : drawingPatriarch) {
            if (shape instanceof Picture){
                if (i == index){
                    ClientAnchor clientAnchor = ((Picture) shape).getClientAnchor();
                    PictureData pictureData = ((Picture) shape).getPictureData();

                    targetDrawingPatriarch.createPicture(clientAnchor,
                            targetSheet.getWorkbook().addPicture(pictureData.getData(), pictureData.getPictureType()));
                }
            }
            i++;
        }
        logger.info("结束copyPicture");
    }

    public static void copySheet(Sheet srcSheet,
                                 Sheet targetSheet,
                                 Collection<CellRangeAddress> mergedRegions,
                                 int[] rowIndexs){

        // 复制合并区域
        for (CellRangeAddress mergedRegion : mergedRegions) {
            targetSheet.addMergedRegion(mergedRegion.copy());
        }

        for (Integer rowIndex : rowIndexs) {
            Row srcRow = srcSheet.getRow(rowIndex);
            Row targetRow = targetSheet.createRow(rowIndex);
            copyRow(srcRow, targetRow);
        }
    }
}
