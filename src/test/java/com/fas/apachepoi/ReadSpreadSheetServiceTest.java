package com.fas.apachepoi;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class ReadSpreadSheetServiceTest {

    @Test
    public void testReadSpreadSheet() throws IOException {
        new ReadSpreadSheetService().readWorkBook();
    }
}