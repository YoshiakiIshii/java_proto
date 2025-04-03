package com.example.exceltemplate;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.Arrays;
import java.util.HashMap;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import com.opencsv.CSVParser;
import com.opencsv.exceptions.CsvValidationException;

@Component
public class ReportUtility {
    private static final String OUTPUT_FORMAT_EXCEL = "EXCEL";
    private static final String OUTPUT_FORMAT_EXCEL_DIRECT = "DIRECT";

    @Value("${report.output.dir:./testdata/report}")
    private String reportOutputDir;
    @Value("${report.xml.dir:./testdata/xml}")
    private String reportXmlDir;
    @Value("${report.template.dir:./testdata/template}")
    private String reportTemplateDir;

    public void printDirs() {
        System.out.println("reportOutputDir: " + Paths.get(reportOutputDir, "sample.txt").toAbsolutePath());
        System.out.println("reportXmlDir: " + Paths.get(reportXmlDir, "sample.xml").toAbsolutePath());
        System.out.println("reportTemplateDir: " + Paths.get(reportTemplateDir, "sample.xlsx").toAbsolutePath());
    }

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

    /**
     * 指定されたデータファイルを基にExcel帳票を作成し、保存された帳票ファイルのパスを返します。
     *
     * <p>
     * このメソッドは、データファイルを読み込み、関数部およびデータ部を処理してExcelワークブックを生成します。
     * 生成されたExcelファイルは一時ディレクトリに保存され、すべてのユーザーが読み書き可能な権限が設定されます。
     * Windows環境では、POSIXファイル権限の設定に関する例外が無視されます。
     * </p>
     *
     * @param reportDataFile レポートデータが含まれる入力ファイル
     * @return 保存されたExcelレポートファイルの絶対パス。エラーが発生した場合はnullを返します。
     * @throws IOException            入出力エラーが発生した場合
     * @throws CsvValidationException CSVデータの検証エラーが発生した場合
     */
    private String createExcelReport(File reportDataFile) {
        String reportFilePath = null;
        XSSFWorkbook workbook = null;

        // データファイルを関数部→データ部→…と読み込むループ
        try (BufferedReader reader = Files.newBufferedReader(reportDataFile.toPath())) {
            ReportDataFileProcessor processor = new ReportDataFileProcessor();
            processor.setReportDataFileReader(reader);
            while (true) {
                readReportDataFileFunctionSection(processor);
                workbook = processor.getWorkbook();
                if (workbook == null) {
                    return OUTPUT_FORMAT_EXCEL_DIRECT;
                }
                if (!readReportDataFileDataSection(processor)) {
                    break;
                }
            }
            // Excelファイルを保存する
            if (workbook != null) {
                // reportOutputDirにすべてのユーザが読み書き可能なTempFileを作成する
                Path reportFile = Files.createTempFile(Paths.get(reportOutputDir), "report_", ".xlsx");
                try {
                    Files.setPosixFilePermissions(reportFile, PosixFilePermissions.fromString("rw-rw-rw-"));
                } catch (UnsupportedOperationException e) {
                    // Windowsで発生するUnsupportedOperationExceptionを無視する
                }
                try (OutputStream outputStream = Files.newOutputStream(reportFile)) {
                    workbook.write(outputStream);
                    reportFilePath = reportFile.toAbsolutePath().toString();
                } catch (IOException e) {
                    throw new IllegalArgumentException("Error writing Excel file", e);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (CsvValidationException e) {
            e.printStackTrace();
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (IOException ex) {
                    // Excelファイルクローズ時の例外は無視する
                }
            }
        }
        return reportFilePath;
    }

    /**
     * レポートデータファイルの関数セクションを読み込み、処理を行うメソッドです。
     * 
     * <p>
     * このメソッドは、以下の形式で記述された関数セクションを解析し、対応する処理を実行します。
     * </p>
     * <ul>
     * <li>セクションの開始行: <code>&lt;start&gt;</code></li>
     * <li>セクションの終了行: <code>&lt;end&gt;</code></li>
     * <li>関数行の形式: <code>関数名=パラメータ</code></li>
     * </ul>
     * 
     * <p>
     * サポートされる関数名とその処理内容:
     * </p>
     * <ul>
     * <li><code>VrSetForm</code>: XML様式ファイル名とモードを設定します。</li>
     * <li><code>VrComout</code>: コマンドに応じてExcelテンプレートの操作を行います。</li>
     * </ul>
     * 
     * <p>
     * サポートされるコマンドとその処理内容:
     * </p>
     * <ul>
     * <li><code>XSFN</code>: テンプレートExcelファイルを読み込みます。</li>
     * <li><code>XSSA</code>: 指定されたシートをアクティブにします。</li>
     * <li><code>XSSC</code>: シートを複製し、新しい名前を設定します。</li>
     * <li><code>XSSD</code>: 指定されたシートを削除します。</li>
     * </ul>
     * 
     * @param processor レポートデータファイルの処理を行う {@link ReportDataFileProcessor} オブジェクト
     * @throws IOException              ファイルの読み込み中にエラーが発生した場合
     * @throws IllegalArgumentException 入力データの形式が不正、または未対応のコマンドや関数名が指定された場合
     */
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
            if (functionAndParam.length < 2) {
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
                    String xmlFileName = vrSetFormParam[0];
                    if (!xmlFileName.endsWith(".xml")) {
                        xmlFileName += ".xml";
                    }
                    processor.setXmlFormFileName(xmlFileName);
                    processor.setMode(Integer.parseInt(vrSetFormParam[1]));
                    break;

                case "VrComout":
                    // paramが"コマンド パラメータ"形式のため、それを分割する
                    String[] vrComoutParam = param.split(" ");
                    if (vrComoutParam.length < 2) {
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
                            XSSFWorkbookFactory workbookFactory = new XSSFWorkbookFactory();
                            try (InputStream inputStream = Files.newInputStream(templateExcelFilePath)) {
                                workbook = workbookFactory.create(inputStream);
                            } catch (IOException e) {
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

    /**
     * レポートデータファイルのデータセクションを読み込み、その内容を処理します。
     * 
     * <p>
     * このメソッドは、データセクションのCSVヘッダ行を読み込み、その後の行を
     * ファイルの終端または"<start>"行が見つかるまで解析します。"<start>"が見つかった場合、
     * リーダーをマークした位置にリセットし、{@code true} を返します。それ以外の場合は、
     * CSVデータ行を処理し、フィールド名をExcelシート内の対応する位置にマッピングして
     * 値を設定します。
     * </p>
     * 
     * @param processor {@link ReportDataFileProcessor} のインスタンスで、レポートデータ
     *                  ファイルリーダー、ワークブック、およびその他の必要なリソースに
     *                  アクセスを提供します。
     * @return "<start>" 行が見つかった場合は {@code true}、それ以外の場合は {@code false}。
     * @throws IOException            ファイルの読み込み中にI/Oエラーが発生した場合。
     * @throws CsvValidationException CSVデータの解析中にエラーが発生した場合。
     */
    private boolean readReportDataFileDataSection(ReportDataFileProcessor processor)
            throws IOException, CsvValidationException {
        BufferedReader reader = processor.getReportDataFileReader();
        CSVParser csvParser = new CSVParser();
        XSSFWorkbook workbook = processor.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(workbook.getActiveSheetIndex());
        HashMap<String, ReportFormatField> reportFormatFieldMap = getReportFormatFieldMap(processor);

        // データセクションの最初の行であるCSVヘッダ行を読み込みパースする
        String[] header = csvParser.parseLine(reader.readLine());
        if (header == null) {
            // CSVヘッダ行が読み込めない場合は、ファイルの終端
            return false;
        }

        // CSVデータ行を読み込み、ファイルが終了するか、<start>行が見つかるまでループする
        // 行を読み込む前にmarkしておき、<start>行が見つかった場合はresetする
        while (true) {
            reader.mark(1024); // 1024バイトまでmarkする
            String line = reader.readLine();
            if (line == null) {
                // ファイルの終端
                return false;
            }
            // <start>行が見つかった場合は、markした位置に戻してtrueを返す
            if ("<start>".equals(line)) {
                reader.reset();
                return true;
            }

            // CSVデータ行をパースする
            String[] data = csvParser.parseLine(line);
            if (data == null) {
                // CSVデータ行が読み込めない場合はその行を読み飛ばす
                continue;
            }

            // フィールド名と値をマップに格納する
            for (int i = 0; i < header.length; i++) {
                String fieldName = header[i];
                String fieldValue = data[i];
                ReportFormatField reportFormatField = reportFormatFieldMap.get(fieldName);
                if (reportFormatField != null) {
                    // フィールド名がマップに存在する場合、locationに指定された位置に値をセットする
                    CellReference cellRef = new CellReference(reportFormatField.getLocation());
                    Row row = sheet.getRow(cellRef.getRow());
                    if (row == null) {
                        row = sheet.createRow(cellRef.getRow());
                    }
                    Cell cell = row.getCell(cellRef.getCol());
                    if (cell == null) {
                        cell = row.createCell(cellRef.getCol());
                    }
                    cell.setCellValue(fieldValue);

                }
            }
        }
    }

    /**
     * 指定されたプロセッサに関連付けられた様式定義フィールドのマップを取得します。
     * マップがまだ初期化されていない場合、様式定義XMLファイルを解析してマップを構築します。
     *
     * @param processor {@link ReportDataFileProcessor} のインスタンスで、様式定義XMLファイル名と
     *                  様式定義フィールドマップを含みます。
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
