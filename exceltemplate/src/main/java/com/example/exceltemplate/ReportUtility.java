package com.example.exceltemplate;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

@Component
public class ReportUtility {
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
        if ("EXCEL".equals(outputFormat)) {
            // Excel形式のレポートを作成する処理
            reportFilePath = createExcelReport(reportDataFile);
        } else {
            // 何もしない
        }
        return reportFilePath;
    }

    // Excel形式のレポートを作成する
    // @param reportDataFile データファイル
    // @return 作成したExcel形式のレポートパス
    private String createExcelReport(File reportDataFile) {
        String reportFilePath = null;

        // データファイルを1行ずつ読み込むループ
        try (BufferedReader reader = Files.newBufferedReader(reportDataFile.toPath())) {
            String line;
            boolean isHeaderSection = false;
            boolean nextDataSection = true;
            String xmlFormFileName = null;
            int mode = -1;
            Workbook workbook = null;

            while ((line = reader.readLine()) != null) {
                // <start>の行から<end>の行までをヘッダセクション、
                // それ以外の行をデータセクションと識別する
                if ("<start>".equals(line)) {
                    isHeaderSection = true;
                } else if ("<end>".equals(line)) {
                    isHeaderSection = false;
                    nextDataSection = true;
                } else {
                    if (isHeaderSection) {
                        // 1行が"関数名=パラメータ"形式のため、関数名とパラメータに分割する
                        String[] functionAndParam = line.split("=");
                        if (functionAndParam.length >= 2) {
                            // フォーマットエラー
                            throw new IllegalArgumentException("フォーマットエラー");
                        }
                        String functionName = functionAndParam[0];
                        String param = String.join("=",
                                Arrays.copyOfRange(functionAndParam, 1, functionAndParam.length));
                        // 関数名に応じた処理を行う
                        switch (functionName) {
                            case "VrSetForm":
                                // paramが"XML様式ファイル名,モード"形式のため、それを分割する
                                String[] vrSetFormParam = param.split(",");
                                if (vrSetFormParam.length != 2) {
                                    // フォーマットエラー
                                    throw new IllegalArgumentException("フォーマットエラー");
                                }
                                xmlFormFileName = vrSetFormParam[0];
                                mode = Integer.parseInt(vrSetFormParam[1]);
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
                                        // Excelファイルを読み込む
                                        workbook = WorkbookFactory.create(Files
                                                .newInputStream(Paths.get(reportTemplateDir, templateExcelFileName)));
                                        break;
                                    case "XSSA":
                                        String activateSheetName = null;
                                        int activateSheetNo = -1;
                                        if (commandParams[0].startsWith("NAME=")) {
                                            // commandParams[0]を"NAME=シート名"形式のため、分割してシート名を取得する
                                            activateSheetName = commandParams[0].substring("NAME=".length());
                                        } else if (commandParams[0].startsWith("NO=")) {
                                            // commandParams[0]を"NO=シート番号"形式のため、分割してシート番号を取得する
                                            activateSheetNo = Integer
                                                    .parseInt(commandParams[0].substring("NO=".length()));
                                        } else {
                                            // フォーマットエラー
                                            throw new IllegalArgumentException("フォーマットエラー");
                                        }
                                        break;
                                    case "XSSC":
                                        String fromSheetName = null;
                                        int fromSheetNo = -1;
                                        String toSheetName = null;
                                        if (commandParams[0].startsWith("NAME=")) {
                                            // commandParams[0]を"NAME=シート名"形式のため、分割してシート名を取得する
                                            fromSheetName = commandParams[0].substring("NAME=".length());
                                        } else if (commandParams[0].startsWith("NO=")) {
                                            // commandParams[0]を"NO=シート番号"形式のため、分割してシート番号を取得する
                                            fromSheetNo = Integer.parseInt(commandParams[0].substring("NO=".length()));
                                        } else {
                                            // フォーマットエラー
                                            throw new IllegalArgumentException("フォーマットエラー");
                                        }

                                        if (commandParams[1].startsWith("CHANGE=")) {
                                            // commandParams[1]を"CHANGE=シート名"形式のため、分割してシート名を取得する
                                            toSheetName = commandParams[1].substring("CHANGE=".length());
                                        } else {
                                            // フォーマットエラー
                                            throw new IllegalArgumentException("フォーマットエラー");
                                        }
                                        break;
                                    case "XSSD":
                                        String deleteSheetName = null;
                                        int deleteSheetNo = -1;
                                        if (commandParams[0].startsWith("NAME=")) {
                                            // commandParams[0]を"NAME=シート名"形式のため、分割してシート名を取得する
                                            deleteSheetName = commandParams[0].substring("NAME=".length());
                                        } else if (commandParams[0].startsWith("NO=")) {
                                            // commandParams[0]を"NO=シート番号"形式のため、分割してシート番号を取得する
                                            deleteSheetNo = Integer
                                                    .parseInt(commandParams[0].substring("NO=".length()));
                                        } else {
                                            // フォーマットエラー
                                            throw new IllegalArgumentException("フォーマットエラー");
                                        }
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
                    } else {
                        if (nextDataSection) {
                            // データセクションのヘッダ行の処理
                            if (workbook == null) {
                                return "DIRECT";
                            }

                            nextDataSection = false;
                        } else {
                            // データセクションのデータ行の処理
                        }
                    }
                }

            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return reportFilePath;
    }
}
