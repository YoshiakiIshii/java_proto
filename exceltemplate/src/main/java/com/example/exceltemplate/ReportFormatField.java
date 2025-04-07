package com.example.exceltemplate;

import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.chrono.JapaneseDate;
import java.util.ArrayList;

import lombok.Getter;
import lombok.Setter;

@Setter
@Getter
public class ReportFormatField {
    private String fieldName;
    private String location;
    private String formulaString;

    public String execFormula(String targetString) {
        if (formulaString == null || formulaString.isEmpty()) {
            return targetString;
        }
        return (String) execFormulaString(formulaString, targetString);
    }

    public Object execFormulaString(String formulaString, String targetString) {
        // 開くカッコが無ければ、関数ではなくフィールド名または即値と判断し、
        // フィールド名と一致すれば、targetStringを返す。それ以外はそのまま返す。
        int index = formulaString.indexOf('(');
        if (index == -1) {
            if (formulaString.equals(fieldName)) {
                return targetString;
            } else {
                return formulaString;
            }
        }

        // 関数名を取得
        String functionName = formulaString.substring(0, index);

        // 引数を取得
        String argsStr = formulaString.substring(index + 1, formulaString.length() - 1).trim();
        ArrayList<Object> argsList = new ArrayList<>();

        // 引数をカンマで分割、ただし、引数が関数である可能性を考慮し、(...)で囲まれた部分は無視する
        int parenCount = 0;
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < argsStr.length(); i++) {
            if (argsStr.charAt(i) == '(') {
                parenCount++;
            } else if (argsStr.charAt(i) == ')') {
                parenCount--;
            } else if (argsStr.charAt(i) == ',' && parenCount == 0) {
                argsList.add(execFormulaString(sb.toString().trim(), targetString));
                sb.setLength(0); // StringBuilderをクリア
                continue;
            }
            if (parenCount < 0) {
                throw new IllegalArgumentException("数式が解析できないため、エラーが発生しました");
            }
            sb.append(argsStr.charAt(i));
        }
        argsList.add(execFormulaString(sb.toString().trim(), targetString));

        // 関数を実行
        switch (functionName) {
            case "NENGO":
                if (argsList.size() != 1 || !(argsList.get(0) instanceof LocalDate)) {
                    throw new IllegalArgumentException("引数はLocalDate型でなければなりません");
                }
                return getJapaneseEra((LocalDate) argsList.get(0));
            case "CTOD":
                if (argsList.size() != 1 || !(argsList.get(0) instanceof String)) {
                    throw new IllegalArgumentException("引数はString型でなければなりません");
                }
                return convertStringToDate((String) argsList.get(0));
            case "WYEAR":
                if (argsList.size() != 1 || !(argsList.get(0) instanceof LocalDate)) {
                    throw new IllegalArgumentException("引数はLocalDate型でなければなりません");
                }
                return getJapaneseYear((LocalDate) argsList.get(0));
            case "FORMAT":
                if (argsList.size() != 2 || !(argsList.get(0) instanceof Number) || !(argsList.get(1) instanceof String)) {
                    throw new IllegalArgumentException("引数はNumber型とString型でなければなりません");
                }
                return formatNumber(((Number) argsList.get(0)).doubleValue(), (String) argsList.get(1));
            default:
                return null;
        }
    }

    /**
     * 指定された日付を和暦の元号と年に変換して文字列として返します。
     *
     * <p>例: 令和5年</p>
     *
     * @param date 和暦に変換する対象の日付。nullであってはなりません。
     * @return 和暦の元号と年を表す文字列。
     * @throws IllegalArgumentException 引数がnullの場合にスローされます。
     */
    public static String getJapaneseEra(LocalDate date) {
        // 引数がnullの場合はエラーを投げる
        if (date == null) {
            throw new IllegalArgumentException("引数はnullであってはいけません");
        }

        // JapaneseDateを使用して和暦の元号を取得
        JapaneseDate japaneseDate = JapaneseDate.from(date);
        String eraName = japaneseDate.getEra().getDisplayName(java.time.format.TextStyle.FULL, java.util.Locale.JAPAN);
        int yearInEra = japaneseDate.get(java.time.temporal.ChronoField.YEAR_OF_ERA);

        StringBuilder sb = new StringBuilder();
        sb.append(eraName).append(yearInEra).append("年");
        return sb.toString(); // 元号と年を返す
    }

    /**
     * 指定された文字列をLocalDate型に変換します。
     * 
     * <p>このメソッドは、入力された日付文字列が"YYYY/MM/DD"形式である場合に、
     * それをLocalDate型に変換して返します。それ以外の場合や、引数がnullの場合はnullを返します。</p>
     * 
     * @param dateStr 日付を表す文字列（形式: "YYYY/MM/DD"）
     * @return 変換されたLocalDateオブジェクト、または入力が無効な場合はnull
     */
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

    /**
     * 指定された日付を和暦の年（元号の年）として取得します。
     *
     * @param date 和暦の年を取得する対象の日付。nullであってはいけません。
     * @return 和暦の元号の年（例: 平成30年の場合は30を返します）。
     * @throws IllegalArgumentException 引数がnullの場合にスローされます。
     */
    public static int getJapaneseYear(LocalDate date) {
        // 引数がnullの場合はエラーを投げる
        if (date == null) {
            throw new IllegalArgumentException("引数はnullであってはいけません");
        }

        // JapaneseDateを使用して和暦の元号を取得
        JapaneseDate japaneseDate = JapaneseDate.from(date);
        return japaneseDate.get(java.time.temporal.ChronoField.YEAR_OF_ERA);
    }

    /**
     * 指定された数値を指定されたフォーマットに基づいてフォーマットします。
     *
     * <p>このメソッドは、フォーマット文字列内の特定の文字を置換してから、
     * {@link DecimalFormat} を使用して数値をフォーマットします。
     * フォーマット文字列内の以下の置換が行われます:
     * <ul>
     *   <li>ダブルクォート (") は削除されます。</li>
     *   <li>'Z' は '#' に置換されます。</li>
     *   <li>'9' は '0' に置換されます。</li>
     * </ul>
     *
     * @param number フォーマット対象の数値
     * @param format フォーマット文字列
     * @return フォーマットされた数値の文字列
     * @throws IllegalArgumentException フォーマット文字列が無効な場合
     */
    public static String formatNumber(double number, String format) {
        if (format == null || format.isEmpty()) {
            throw new IllegalArgumentException("フォーマット文字列が無効です");
        }

        String sanitizedFormat = format.replace("\"", "")
                                       .replace('Z', '#')
                                       .replace('9', '0');

        try {
            DecimalFormat decimalFormat = new DecimalFormat(sanitizedFormat);
            return decimalFormat.format(number);
        } catch (IllegalArgumentException e) {
            throw new IllegalArgumentException("フォーマット文字列が無効です: " + sanitizedFormat, e);
        }
    }

}
