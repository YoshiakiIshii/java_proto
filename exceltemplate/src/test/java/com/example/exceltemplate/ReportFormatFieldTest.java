package com.example.exceltemplate;

import static org.assertj.core.api.Assertions.*;
import org.junit.jupiter.api.Test;

class ReportFormatFieldTest {

    @Test
    void testExecFormulaWithNullFormulaString() {
        ReportFormatField field = new ReportFormatField();
        field.setFormulaString(null);
        String result = field.execFormula("test");
        assertThat(result).isEqualTo("test");
    }

    @Test
    void testExecFormulaWithEmptyFormulaString() {
        ReportFormatField field = new ReportFormatField();
        field.setFormulaString("");
        String result = field.execFormula("test");
        assertThat(result).isEqualTo("test");
    }

    @Test
    void testExecFormulaStringWithFieldNameMatch() {
        ReportFormatField field = new ReportFormatField();
        field.setFieldName("fieldName");
        String result = (String) field.execFormulaString("fieldName", "targetValue");
        assertThat(result).isEqualTo("targetValue");
    }

    @Test
    void testExecFormulaStringWithNengoFunction() {
        ReportFormatField field = new ReportFormatField();
        field.setFieldName("作成日");
        field.setLocation("A1");
        field.setFormulaString("NENGO(CTOD(作成日))");
        String result = (String) field.execFormula("2025/04/07");
        assertThat(result).isEqualTo("令和7年");
    }

    @Test
    void testFormatNumberWithValidFormat() {
        String result = ReportFormatField.formatNumber(1234.56, "\"Z,ZZ0.00\"");
        assertThat(result).isEqualTo("1,234.56");
    }

    @Test
    void testFormatNumberWithZero() {
        String result = ReportFormatField.formatNumber(0, "\"0.00\"");
        assertThat(result).isEqualTo("0.00");
    }

    @Test
    void testFormatNumberWithNegativeNumber() {
        String result = ReportFormatField.formatNumber(-1234.56, "\"Z,ZZZ.00\"");
        assertThat(result).isEqualTo("-1,234.56");
    }

    @Test
    void testFormatNumberWithCustomFormat() {
        String result = ReportFormatField.formatNumber(1234.56, "\"00000.00\"");
        assertThat(result).isEqualTo("01234.56");
    }

    @Test
    void testFormatNumberWithRealFormula() {
        ReportFormatField field = new ReportFormatField();
        field.setFieldName("作成日");
        field.setLocation("A1");
        field.setFormulaString("FORMAT(WYEAR(CTOD(作成日)),\"Z9\")");
        String result = (String) field.execFormula("2025/04/07");
        assertThat(result).isEqualTo("7");
    }
}