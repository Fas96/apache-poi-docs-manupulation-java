package com.fas.apachepoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

@Service
public class CreateSpreadSheetService {

    public String writeToExcel() throws IOException{
        //
        Workbook wb = new XSSFWorkbook();

        //default format
//        Workbook wb = new HSSFWorkbook();

        //sheet in workbook
        Sheet sheet = wb.createSheet("1st step to trip to Mars");
        Sheet home = wb.createSheet("Home Page");
        CreationHelper creationHelper = wb.getCreationHelper();

        //row index and column index
        int rowindex=0;
        int columnindex=0;

        Row row=sheet.createRow(rowindex++);
        Cell cell= row.createCell(columnindex++);
        cell.setCellValue("Todo  list");


        // setting the styling
        Font font= wb.createFont();
        font.setFontHeightInPoints(Short.parseShort("16"));
        font.setFontName("Ubuntu");

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cell.setCellStyle(cellStyle);

        row = sheet.createRow(rowindex++);
        columnindex = 0;
        row.createCell(columnindex++).setCellValue("Remember to bring potatoes and milk for 10 years");
        row.createCell(columnindex++).setCellValue("Bla bla bla");
        row.createCell(columnindex++).setCellValue(true);
        row.createCell(columnindex++).setCellValue(1.123123223);
        sheet.createRow(rowindex++).createCell(0).setCellValue("Bring computer with Linux!");
        sheet.createRow(rowindex++).createCell(0).setCellValue("Toiletpaper!");
        sheet.createRow(rowindex++).createCell(0).setCellValue("Fresh underwear x10!");
        try (OutputStream fileOut = new FileOutputStream("./output/XSSFWorkbook.xls")) {
            wb.write(fileOut);
        }


        return "done";
    }
}
