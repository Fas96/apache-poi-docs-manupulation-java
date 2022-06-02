package com.fas.apachepoi.msword;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class MSWordWriterServiceTest {

    @Test
    public void testWriteWordFile() throws IOException{
        new MSWordWriterService().writeASmallStory();
    }
}