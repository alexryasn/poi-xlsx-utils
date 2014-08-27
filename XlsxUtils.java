package org.gazprom.smpo.service.report;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.sql.Timestamp;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by Ryasale on 18.08.2014.
 */
@Service
public class XlsxUtils {

    private String[] months = {"январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"};

    public String[] getMonths() {
        return months;
    }

    public String getMonth(int index) {
        return months[index];
    }

    /**
     * Format floating numbers to String without unnecessary decimal 0.
     * @param d
     * @return
     */
    public String formatDoubleToString(double d)
    {
        if(d == (int) d)
            return String.format("%d",(int)d);
        else
            return String.format("%s",d);
    }

    /**
     * Make "dd.mm.yyy" from long Timestamp
     */
    public String timestampToString(long timestamp) {
        Timestamp t = new Timestamp(timestamp);
        Date date = new Date(t.getTime());
        return new SimpleDateFormat("dd.MM.yyyy").format(date);
    }

    /**
     * Copy row in XLS file
     *
     * @param workbook
     * @param worksheet
     * @param sourceRowNum
     * @param destinationRowNum
     */
    private static void copyRow(HSSFWorkbook workbook, HSSFSheet worksheet, int sourceRowNum, int destinationRowNum) {
        // Get the source / new row
        HSSFRow newRow = worksheet.getRow(destinationRowNum);
        HSSFRow sourceRow = worksheet.getRow(sourceRowNum);
        short sourceRowHeight = sourceRow.getHeight();

        // If the row exist in destination, push down all rows by 1 else create a new row
        if (newRow != null) {
            worksheet.shiftRows(destinationRowNum, worksheet.getLastRowNum(), 1);
        } else {
            newRow = worksheet.createRow(destinationRowNum);
            newRow.setHeight(sourceRowHeight);
        }

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            HSSFCell oldCell = sourceRow.getCell(i);
            HSSFCell newCell = newRow.createCell(i);

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                newCell = null;
                continue;
            }

            // Copy style from old cell and apply to new cell
            HSSFCellStyle newCellStyle = workbook.createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());

            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellType());

            // Set the cell data value
            switch (oldCell.getCellType()) {
                case Cell.CELL_TYPE_BLANK:
                    newCell.setCellValue(oldCell.getStringCellValue());
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case Cell.CELL_TYPE_ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    newCell.setCellFormula(oldCell.getCellFormula());
                    break;
                case Cell.CELL_TYPE_NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case Cell.CELL_TYPE_STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
            }
        }

        // If there are are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }

    public int addCellToNewRow(XSSFSheet sheet, int newRowNum, short rowHeight, XSSFCell cell, String data) {
        XSSFRow newRow = sheet.createRow(newRowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(cell.getColumnIndex());
        newCell.setCellValue(data);
        newRowNum = newRowNum + 1;
        return newRowNum;
    }

    public int addCellToNewRow(XSSFSheet sheet, int cellNum, String data, XSSFCellStyle style) {
        int newRowNum = sheet.getLastRowNum() + 1;
        XSSFRow newRow = sheet.createRow(newRowNum);
        XSSFCell newCell = newRow.createCell(cellNum);
        newCell.setCellValue(data);
        newCell.setCellStyle(style);
        newRowNum = newRowNum + 1;
        return newRowNum;
    }

    public void addCellToLastRow(XSSFSheet sheet, int cellNum, String data, XSSFCellStyle style) {
        XSSFRow newRow = sheet.getRow(sheet.getLastRowNum());
        XSSFCell newCell = newRow.createCell(cellNum);
        newCell.setCellValue(data);
        newCell.setCellStyle(style);
    }

    public int addCellToNewRow(XSSFSheet sheet, int newRowNum, short rowHeight, XSSFCell cell, String data, XSSFCellStyle style) {
        XSSFRow newRow = sheet.createRow(newRowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(cell.getColumnIndex());
        newCell.setCellValue(data);
        newCell.setCellStyle(style);
        newRowNum = newRowNum + 1;
        return newRowNum;
    }

    public int addCellToNewRow(XSSFSheet sheet, int newRowNum, short rowHeight, int cellIndex, String data, XSSFCellStyle style) {
        XSSFRow newRow = sheet.createRow(newRowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(cellIndex);
        newCell.setCellValue(data);
        newCell.setCellStyle(style);
        newRowNum = newRowNum + 1;
        return newRowNum;
    }

    public int addCellToNewRow(XSSFSheet sheet, int newRowNum, short rowHeight, XSSFCell cell, String data, XSSFCellStyle style, int cellType) {
        XSSFRow newRow = sheet.createRow(newRowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(cell.getColumnIndex());
        newCell.setCellType(cellType);
        if (cellType == Cell.CELL_TYPE_NUMERIC) {
            newCell.setCellValue(Double.parseDouble(data));
        } else {
            newCell.setCellValue(data);
        }
        newCell.setCellStyle(style);
        newRowNum = newRowNum + 1;
        return newRowNum;
    }

    /**
     * Add cell to exist row. Return next row number.
     *
     * @param sheet
     * @param rowNum
     * @param rowHeight
     * @param colIndex
     * @param data
     * @return
     */
    public int addCellToExistRow(XSSFSheet sheet, int rowNum, short rowHeight, int colIndex, String data) {
        XSSFRow newRow = sheet.getRow(rowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(colIndex);
        newCell.setCellValue(data);
        rowNum = rowNum + 1;
        return rowNum;
    }

    /**
     * Add cell to exist row, using style. Return next row number.
     *
     * @param sheet
     * @param rowNum
     * @param rowHeight
     * @param colIndex
     * @param data
     * @param style
     * @return
     */
    public int addCellToExistRow(XSSFSheet sheet, int rowNum, short rowHeight, int colIndex, String data, XSSFCellStyle style) {
        XSSFRow newRow = sheet.getRow(rowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(colIndex);
        newCell.setCellValue(data);
        newCell.setCellStyle(style);
        rowNum = rowNum + 1;
        return rowNum;
    }

    public int addCellToExistRow(XSSFSheet sheet, int rowNum, short rowHeight, int colIndex, String data, XSSFCellStyle style, int cellType) {
        XSSFRow newRow = sheet.getRow(rowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(colIndex);
        newCell.setCellType(cellType);
        if (cellType == Cell.CELL_TYPE_NUMERIC) {
            newCell.setCellValue(Double.parseDouble(data));
        } else {
            newCell.setCellValue(data);
        }
        newCell.setCellStyle(style);
        rowNum = rowNum + 1;
        return rowNum;
    }

    /**
     * Add cell to exist row. Return next row number. Set auto size column.
     *
     * @param sheet
     * @param rowNum
     * @param rowHeight
     * @param colIndex
     * @param data
     * @return
     */
    public int addCellToExistRow(XSSFSheet sheet, int rowNum, short rowHeight, int colIndex, String data, boolean autoSizeColumn) {
        XSSFRow newRow = sheet.getRow(rowNum);
        newRow.setHeight(rowHeight);
        XSSFCell newCell = newRow.createCell(colIndex);
        newCell.setCellValue(data);
        sheet.autoSizeColumn(colIndex);
        rowNum = rowNum + 1;
        return rowNum;
    }

    public void addMergedRegionStyle(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, XSSFCellStyle style) {
        for (int row = firstRow; row <= lastRow; row++) {
            for (int col = firstCol; col <= lastCol; col++) {
                sheet.getRow(row).getCell(col).setCellStyle(style);
            }
        }
    }

    public void addPictureToXlsx(XSSFWorkbook workbook, byte[] byteArray, XSSFSheet sheet, int columnIndex, int rowIndex) {
        if (byteArray != null) {
            int pictureIdx = workbook.addPicture(byteArray, Workbook.PICTURE_TYPE_JPEG);
            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = workbook.getCreationHelper();
            //Creates the top-level drawing patriarch.
            Drawing drawing = sheet.createDrawingPatriarch();
            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();
            //set top-left corner for the image
            anchor.setCol1(columnIndex);
            anchor.setRow1(rowIndex);
            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //Reset the image to the original size
            pict.resize();
        }
    }

    public void addCellMergeToNextRow(XSSFSheet sheet, int firstCellNum, int lastCellNum, String data, XSSFCellStyle style) {
        int newRowNum =  sheet.getLastRowNum() + 1;
        XSSFRow newRow = sheet.createRow(newRowNum);
        XSSFCell newCell = newRow.createCell(firstCellNum);
        newCell.setCellValue(data);
        newCell.setCellStyle(style);
        for (int k = firstCellNum + 1; k <= lastCellNum; k++) {
            sheet.getRow(sheet.getLastRowNum()).createCell(k).setCellStyle(style);
        }
        sheet.addMergedRegion(new CellRangeAddress(newRowNum, newRowNum, firstCellNum, lastCellNum));
        //newRowNum = newRowNum + 1;
        //return newRowNum;
    }
}
