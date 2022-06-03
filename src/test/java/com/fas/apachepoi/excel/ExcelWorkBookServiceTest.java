package com.fas.apachepoi.excel;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class ExcelWorkBookServiceTest {

    @Test
    public void testWorkBookCreation() throws IOException{
        System.out.println(new ExcelWorkBookService().newExcelWorkBook());
    }
}