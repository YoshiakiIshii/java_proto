package com.example.exceltemplate;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.HashMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.opencsv.CSVParser;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvValidationException;

@Component
public class ReportUtility {
    private static final String OUTPUT_FORMAT_EXCEL = "EXCEL";
    private static final String OUTPUT_FORMAT_EXCEL_DIRECT = "DIRECT";

    @Value("${report.output.dir:./report}")
    private String reportOutputDir;
    @Value("${report.xml.dir:./xml}")
    private String reportXmlDir;
    @Value("${report.template.dir:./template}")
    private String reportTemplateDir;

    // 引数で指定されたデータファイル、出力形式に従い、レポートを作成する
    // @param reportDataFile データファイル
    // @param outputFormat 出力形式
    // @return 作成したレポートパス
    public String outputReport(File reportDataFile, String outputFormat) {
        String reportFilePath = null;
        // 出力形式がEXCELならば、Excel形式のレポートを作成する
        // それ以外ならばnullを返す
        if (OUTPUT_FORMAT_EXCEL.equals(outputFormat)) {
            // Excel形式のレポートを作成する処理
            reportFilePath = createExcelReport(reportDataFile);
        }
        if (reportFilePath == null || OUTPUT_FORMAT_EXCEL_DIRECT.equals(reportFilePath)) {
            // Excel帳票のテンプレート出力方式以外は従来処理で行う
            reportFilePath = null;
        }
        return reportFilePath;
    }

    // Excel形式のレポートを作成する
    // XSFNコマンドでテンプレートExcelファイルが指定されなかった場合は直接出力方式を示す文字列を返す
    // @param reportDataFile データファイル
    // @return 作成したExcel形式のレポートパス
    private String createExcelReport(File reportDataFile) {
        String reportFilePath = null;

        // データファイルを関数部→データ部→…と読み込むループ
        try (BufferedReader reader = Files.newBufferedReader(reportDataFile.toPath())) {
            ReportDataFileProcessor processor = new ReportDataFileProcessor();
            processor.setReportDataFileReader(reader);
            while (true) {
                readReportDataFileFunctionSection(processor);
                if (processor.getWorkbook() == null) {
                    return OUTPUT_FORMAT_EXCEL_DIRECT;
                }
                if (!readReportDataFileDataSection(processor)) {
                    break;
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return reportFilePath;
    }

    // 関数セクションを読み込む
    // @param processor レポートデータファイル処理クラス
    private void readReportDataFileFunctionSection(ReportDataFileProcessor processor) throws IOException {
        BufferedReader reader = processor.getReportDataFileReader();

        // 関数セクションの最初の行である<start>の行を読み込む
        String line = reader.readLine();
        if (!"<start>".equals(line)) {
            throw new IllegalArgumentException("未対応のコマンド");
        }

        // 関数セクションの最後の行である<end>の行までを読み込む
        while (!(line = reader.readLine()).equals("<end>")) {
            // 1行が"関数名=パラメータ"形式のため、関数名とパラメータに分割する
            String[] functionAndParam = line.split("=");
            if (functionAndParam.length >= 2) {
                // フォーマットエラー
                throw new IllegalArgumentException("フォーマットエラー");
            }
            String functionName = functionAndParam[0];
            String param = String.join("=",
                    Arrays.copyOfRange(functionAndParam, 1, functionAndParam.length));
            XSSFWorkbook workbook = null;

            // 関数名に応じた処理を行う
            switch (functionName) {
                case "VrSetForm":
                    // paramが"XML様式ファイル名,モード"形式のため、それを分割する
                    String[] vrSetFormParam = param.split(",");
                    if (vrSetFormParam.length != 2) {
                        // フォーマットエラー
                        throw new IllegalArgumentException("フォーマットエラー");
                    }
                    processor.setXmlFormFileName(vrSetFormParam[0]);
                    processor.setMode(Integer.parseInt(vrSetFormParam[1]));
                    break;

                case "VrComout":
                    // paramが"コマンド パラメータ"形式のため、それを分割する
                    String[] vrComoutParam = param.split(" ");
                    if (vrComoutParam.length >= 2) {
                        // フォーマットエラー
                        throw new IllegalArgumentException("フォーマットエラー");
                    }
                    String command = vrComoutParam[0];
                    String[] commandParams = Arrays.copyOfRange(vrComoutParam, 1, vrComoutParam.length);
                    // コマンドに応じた処理を行う
                    switch (command) {
                        case "XSFN":
                            // commandParams[0]をテンプレートExcelファイル名として取得する
                            String templateExcelFileName = commandParams[0];
                            Path templateExcelFilePath = Paths.get(reportTemplateDir, templateExcelFileName);
                            if (!Files.exists(templateExcelFilePath)) {
                                // ファイルが存在しない
                                throw new IllegalArgumentException("ファイルが存在しない");
                            }
                            // Excelファイルを読み込む
                            try {
                                workbook = new XSSFWorkbook(templateExcelFilePath.toFile());
                            } catch (InvalidFormatException e) {
                                throw new IllegalArgumentException("Invalid Excel file format", e);
                            }
                            workbook.setActiveSheet(0);
                            processor.setWorkbook(workbook);
                            break;
                        case "XSSA":
                            workbook = processor.getWorkbook();
                            if (commandParams[0].startsWith("NAME=")) {
                                // commandParams[0]を"NAME=シート名"形式のため、分割してシート名を取得する
                                String activateSheetName = commandParams[0].substring("NAME=".length());
                                int activateSheetNo = workbook.getSheetIndex(activateSheetName);
                                if (activateSheetNo == -1) {
                                    // シートが存在しない
                                    throw new IllegalArgumentException("シートが存在しない");
                                }
                                workbook.setActiveSheet(activateSheetNo);
                            } else if (commandParams[0].startsWith("NO=")) {
                                // commandParams[0]を"NO=シート番号"形式のため、分割してシート番号を取得する
                                int activateSheetNo = Integer
                                        .parseInt(commandParams[0].substring("NO=".length()));
                                workbook.setActiveSheet(activateSheetNo);
                            } else {
                                // フォーマットエラー
                                throw new IllegalArgumentException("フォーマットエラー");
                            }
                            break;
                        case "XSSC":
                            workbook = processor.getWorkbook();
                            int fromSheetNo = -1;
                            if (commandParams[0].startsWith("NAME=")) {
                                // commandParams[0]を"NAME=シート名"形式のため、分割してシート名を取得する
                                String fromSheetName = commandParams[0].substring("NAME=".length());
                                fromSheetNo = workbook.getSheetIndex(fromSheetName);
                                if (fromSheetNo == -1) {
                                    // シートが存在しない
                                    throw new IllegalArgumentException("シートが存在しない");
                                }
                            } else if (commandParams[0].startsWith("NO=")) {
                                // commandParams[0]を"NO=シート番号"形式のため、分割してシート番号を取得する
                                fromSheetNo = Integer.parseInt(commandParams[0].substring("NO=".length()));
                            } else {
                                // フォーマットエラー
                                throw new IllegalArgumentException("フォーマットエラー");
                            }

                            if (commandParams[1].startsWith("CHANGE=")) {
                                // commandParams[1]を"CHANGE=シート名"形式のため、分割してシート名を取得する
                                String toSheetName = commandParams[1].substring("CHANGE=".length());
                                XSSFSheet sheet = workbook.cloneSheet(fromSheetNo, toSheetName);
                                workbook.setActiveSheet(workbook.getSheetIndex(sheet));
                            } else {
                                // フォーマットエラー
                                throw new IllegalArgumentException("フォーマットエラー");
                            }
                            break;

                        case "XSSD":
                            workbook = processor.getWorkbook();
                            int deleteSheetNo = -1;
                            if (commandParams[0].startsWith("NAME=")) {
                                // commandParams[0]を"NAME=シート名"形式のため、分割してシート名を取得する
                                String deleteSheetName = commandParams[0].substring("NAME=".length());
                                deleteSheetNo = workbook.getSheetIndex(deleteSheetName);
                                if (deleteSheetNo == -1) {
                                    // シートが存在しない
                                    throw new IllegalArgumentException("シートが存在しない");
                                }
                            } else if (commandParams[0].startsWith("NO=")) {
                                // commandParams[0]を"NO=シート番号"形式のため、分割してシート番号を取得する
                                deleteSheetNo = Integer
                                        .parseInt(commandParams[0].substring("NO=".length()));
                            } else {
                                // フォーマットエラー
                                throw new IllegalArgumentException("フォーマットエラー");
                            }
                            workbook.removeSheetAt(deleteSheetNo);
                            break;

                        default:
                            // 未対応のコマンド
                            throw new IllegalArgumentException("未対応のコマンド");
                    }
                    break;
                default:
                    // 未対応の関数名
                    throw new IllegalArgumentException("未対応の関数名");
            }
        }
    }

    // データセクションを読み込む
    // @param processor レポートデータファイル処理クラス
    // @return 次に関数セクションが続く場合はtrue、ファイルが終了する場合はfalse
    private boolean readReportDataFileDataSection(ReportDataFileProcessor processor) throws IOException, CsvValidationException {
        BufferedReader reader = processor.getReportDataFileReader();
        CSVParser csvParser = new CSVParser();
        XSSFWorkbook workbook = processor.getWorkbook();
        HashMap<String, ReportFormatField> reportFormatFieldMap = getReportFormatFieldMap(processor);

        // データセクションの最初の行であるCSVヘッダ行を読み込みパースする
        String[] header = csvParser.parseLine(reader.readLine());
        if (header == null) {
            // CSVヘッダ行が読み込めない場合は、ファイルの終端
            return false;
        }

        // markとresetを使って、関数セクションの開始位置に戻す
        return false;
    }

    /**
     * 指定されたプロセッサに関連付けられたレポートフォーマットフィールドのマップを取得します。
     * マップがまだ初期化されていない場合、XMLファイルを解析してマップを構築します。
     *
     * @param processor {@link ReportDataFileProcessor} のインスタンスで、XMLファイル名と
     *                  現在のレポートフォーマットフィールドマップを含みます。
     * @return フィールド名をキーとし、{@link ReportFormatField} オブジェクトを値とする
     *         {@link HashMap} を返します。
     * @throws IllegalArgumentException XMLファイルが存在しない場合、またはXMLの解析や
     *                                  設定中にエラーが発生した場合にスローされます。
     */
    private HashMap<String, ReportFormatField> getReportFormatFieldMap(ReportDataFileProcessor processor) {
        HashMap<String, ReportFormatField> reportFormatFieldMap = processor.getReportFormatFieldMap();
        if (reportFormatFieldMap == null) {
            reportFormatFieldMap = new HashMap<>();
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            // XMLファイルを読み込む
            Path xmlFilePath = Paths.get(reportXmlDir, processor.getXmlFormFileName());
            if (!Files.exists(xmlFilePath)) {
                // ファイルが存在しない
                throw new IllegalArgumentException("ファイルが存在しない");
            }
            // XMLファイルをパースして、reportFormatFieldMapに格納する
            DocumentBuilder builder;
            try {
                builder = factory.newDocumentBuilder();
            } catch (ParserConfigurationException e) {
                throw new IllegalArgumentException("Error configuring XML parser", e);
            }
            Document document;
            try {
                document = builder.parse(xmlFilePath.toFile());
            } catch (SAXException | IOException e) {
                throw new IllegalArgumentException("Error parsing XML file", e);
            }
            Element root = document.getDocumentElement();
            NodeList fieldList = root.getElementsByTagName("Field");
            for (int i = 0; i < fieldList.getLength(); i++) {
                Element fieldElement = (Element) fieldList.item(i);
                String fieldName = fieldElement.getAttribute("name");
                String fieldLocation = fieldElement.getAttribute("strComment");
                String fieldEditFormula = fieldElement.getAttribute("strEditFormula");
                ReportFormatField reportFormatField = new ReportFormatField();
                reportFormatField.setFieldName(fieldName);
                reportFormatField.setLocation(fieldLocation);
                reportFormatField.setFormulaString(fieldEditFormula);
                reportFormatFieldMap.put(fieldName, reportFormatField);
            }
            processor.setReportFormatFieldMap(reportFormatFieldMap);
        }
        return reportFormatFieldMap;
    }
}
