package com.example.exceltemplate;

import java.io.BufferedReader;
import java.util.HashMap;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class ReportDataFileProcessor {
    private BufferedReader reportDataFileReader;
    private String xmlFormFileName;
    private int mode;
    private XSSFWorkbook workbook;
    private HashMap<String, ReportFormatField> reportFormatFieldMap;
}
