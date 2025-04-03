package com.example.exceltemplate;

import static org.junit.jupiter.api.Assertions.*;
import static org.mockito.Mockito.*;
import java.io.File;
import java.nio.file.Paths;

import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

@SpringBootTest
class ReportUtilityTest {

    @BeforeEach
    void setUp() {
        MockitoAnnotations.openMocks(this);
    }
    
    @Mock
    private ReportDataFileProcessor mockProcessor;

    @Autowired
    private ReportUtility reportUtility;

    @Test
    void testOutputReport_NonExcelFormat() {
        // Arrange
        File mockFile = mock(File.class);

        // Act
        String result = reportUtility.outputReport(mockFile, "PDF");

        // Assert
        assertNull(result);
    }

    @Test
    void testOutputReport_NullOutputFormat() {
        // Arrange
        File mockFile = mock(File.class);

        // Act
        String result = reportUtility.outputReport(mockFile, null);

        // Assert
        assertNull(result);
    }

    @Test
    void testOutputReport_printDirs() {
        reportUtility.printDirs();
    }

    @Test
    void testOutputReport_outputReport01() {
        // Arrange
        File mockFile = Paths.get("./testdata/csv", "data1.csv").toFile();

        // Act
        String result = reportUtility.outputReport(mockFile, "EXCEL");

        // Assert
        assertNotNull(result);
    }
}
