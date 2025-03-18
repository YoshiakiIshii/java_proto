package com.example.exceltemplate;

import java.time.LocalDate;
import java.time.chrono.JapaneseDate;

import javax.script.Compilable;
import javax.script.CompiledScript;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class ReportFormatField {
    private String fieldName;
    private String location;
    private String formulaString;

    private Compilable compiler;
    private CompiledScript cscript;

    public void setFormulaString(String formulaString) {
        this.formulaString = formulaString;

        String scriptStr = "";
        if (formulaString == null) {
            return;
        }
        if (formulaString.contains("NENGO")) {
            scriptStr += """
                    function NENGO(date) {
                        return ReportFormatField.getJapaneseEra(date);
                    }
                        """;
        }
        if (formulaString.contains("CTOD")) {
            scriptStr += """
                    function CTOD(dateStr) {
                        return ReportFormatField.convertStringToDate(dateStr);
                    }
                        """;
        }
        if (scriptStr != "") {
            scriptStr = """
                    var ReportFormatField = Java.type("com.example.exceltemplate.ReportFormatField");
                    """ + scriptStr + formulaString;

            if (compiler == null) {
                ScriptEngine engine = new ScriptEngineManager().getEngineByName("nashorn");
                compiler = (Compilable) engine;
            }

            try {
                cscript = compiler.compile(scriptStr);
            } catch (javax.script.ScriptException e) {
                throw new IllegalArgumentException("数式が解析できないため、エラーが発生しました");
            }
        }

    }

    public static String getJapaneseEra(LocalDate date) {
        // 引数がnullの場合はエラーを投げる
        if (date == null) {
            throw new IllegalArgumentException("引数はnullであってはいけません");
        }

        // JapaneseDateを使用して和暦の元号を取得
        JapaneseDate japaneseDate = JapaneseDate.from(date);
        String eraName = japaneseDate.getEra().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.JAPAN);
        int yearInEra = japaneseDate.get(java.time.temporal.ChronoField.YEAR_OF_ERA);

        return eraName + yearInEra + "年"; // 元号と年を返す
    }

    // YYYY/MM/DD形式の文字列をLocalDate型に変換して返す
    // @param dateStr 変換対象の文字列
    // @return 変換結果
    public static LocalDate convertStringToDate(String dateStr) {
        // 引数がnullの場合はnullを返す
        if (dateStr == null) {
            return null;
        }

        // 日付のフォーマットがYYYY/MM/DDでない場合はnullを返す
        if (!dateStr.matches("\\d{4}/\\d{2}/\\d{2}")) {
            return null;
        }

        // 日付のフォーマットが正しい場合はjava.util.Date型に変換して返す
        String[] dateArray = dateStr.split("/");
        int year = Integer.parseInt(dateArray[0]);
        int month = Integer.parseInt(dateArray[1]);
        int day = Integer.parseInt(dateArray[2]);

        return LocalDate.of(year, month, day);
    }

}
