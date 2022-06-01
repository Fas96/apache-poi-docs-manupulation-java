package com.fas.apachepoi;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class CreateSpreadSheetServiceTest {

    @Test
    public void testPlan() throws IOException{
        new CreateSpreadSheetService().writeToExcel();
    }

}