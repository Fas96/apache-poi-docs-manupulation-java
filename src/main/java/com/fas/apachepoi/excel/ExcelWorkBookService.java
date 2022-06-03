package com.fas.apachepoi.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.Calendar;
import java.util.Date;

@Service
@Slf4j
public class ExcelWorkBookService {

    public String newExcelWorkBook() throws IOException{
//        POIFSFileSystem fs = new POIFSFileSystem(new File("./output/workbook.xls"));
        Workbook wb = new HSSFWorkbook();
//        Workbook wb = new XSSFWorkbook();


        //try (OutputStream fileOut = new FileOutputStream("workbook.xlsx"))
        newSheetsCreations(wb);

        try (OutputStream fileOut = new FileOutputStream("./output/workbook.xls")) {
            wb.write(fileOut);
        }

        standardTextExtraction(wb);


        return "done";


    }

    private void standardTextExtraction(Workbook wb) throws IOException {
        try (InputStream inp = new FileInputStream("./output/workbook.xls")) {
            ExcelExtractor extractor = new ExcelExtractor((HSSFWorkbook) wb);
            extractor.setFormulasNotResults(true);
            extractor.setIncludeSheetNames(false);
            String text = extractor.getText();
            wb.close();
        }
    }

    private void newSheetsCreations(Workbook wb) {
        Sheet sheet1 = wb.createSheet("Fasnew sheet");
        Sheet sheet2 = wb.createSheet("Fas Data  sheet");
        String safeName = WorkbookUtil.createSafeSheetName("[O'Stocks's sales*?]"); // returns " O'Stocks's sales   "
        Sheet sheet3 = wb.createSheet(safeName);

        createCells(wb, sheet1);

        readingAllSheetContent(wb,sheet1);
    }

    private void readingAllSheetContent(Workbook wb, Sheet sheet1) {
        DataFormatter formatter = new DataFormatter();
         sheet1 = wb.getSheetAt(0);
        for (Row row : sheet1) {
            for (Cell cell : row) {
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
                System.out.print(" - ");
                // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
                String text = formatter.formatCellValue(cell);
                System.out.println(text);
                // Alternatively, get the value and format it yourself
//                switch (cell.getCellType()) {
//                    case CellType.STRING:
//                        System.out.println(cell.getRichStringCellValue().getString());
//                        break;
//                    case CellType.NUMERIC:
//                        if (DateUtil.isCellDateFormatted(cell)) {
//                            System.out.println(cell.getDateCellValue());
//                        } else {
//                            System.out.println(cell.getNumericCellValue());
//                        }
//                        break;
//                    case CellType.BOOLEAN:
//                        System.out.println(cell.getBooleanCellValue());
//                        break;
//                    case CellType.FORMULA:
//                        System.out.println(cell.getCellFormula());
//                        break;
//                    case CellType.BLANK:
//                        System.out.println();
//                        break;
//                    default:
//                        System.out.println();
//                }
            }
        }
    }

    private void createCells(Workbook wb, Sheet sheet1) {
        CreationHelper createHelper = wb.getCreationHelper();

// Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet1.createRow(0);
// Create a cell and put a value in it.
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
// Or do it on one line.
        row.createCell(1).setCellValue(1.2);
        row.createCell(2).setCellValue(createHelper.createRichTextString("This is a string"));
        row.createCell(3).setCellValue(true);


        creatingDateCell(wb, sheet1, createHelper);
//        Row row = sheet1.createRow(2);
        row.setHeightInPoints(30);
        creatingCells(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
        creatingCells(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
        creatingCells(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        creatingCells(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
        creatingCells(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
        creatingCells(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
        creatingCells(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);


        creatingBorders(wb, cell);
    }

    private void creatingBorders(Workbook wb, Cell cell) {
        // Create a cell and put a value in it.

        cell.setCellValue(4);
// Style the cell with borders all around.
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style.setBorderTop(BorderStyle.MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);
    }

    private void creatingCells(Workbook wb, Row row, int i, HorizontalAlignment general, VerticalAlignment center) {
        Cell cell = row.createCell(i);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(general);
        cellStyle.setVerticalAlignment(center);
        cell.setCellStyle(cellStyle);
    }



    private void creatingDateCell(Workbook wb, Sheet sheet1, CreationHelper createHelper) {
        Row row = sheet1.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(new Date());
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("m/d/yy h:mm"));
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);
//you can also set date as java.util.Calendar
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);
    }
}
