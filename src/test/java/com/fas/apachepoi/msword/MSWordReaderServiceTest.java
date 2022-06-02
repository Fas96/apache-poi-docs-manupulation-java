package com.fas.apachepoi.msword;

import org.junit.jupiter.api.Test;

import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class MSWordReaderServiceTest {

    @Test
    public void testReadWordFile()throws IOException{
        new MSWordReaderService().readStuff();
    }
}