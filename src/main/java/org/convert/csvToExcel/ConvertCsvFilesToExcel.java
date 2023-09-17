package org.convert.csvToExcel;

import java.util.LinkedHashMap;
import java.util.Map;

public class ConvertCsvFilesToExcel {
    private static final String sourceCsvFilesLocation = "";
    private static final String destinationExcelFileLocation = "";
    public static final char DEFAULT_CSV_FILE_DELIMITER = ',';
    public static final String XLSX_EXTENSION = ".xlsx";
    public static final String XLSX_FILE_FORMAT_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    public static final Integer DEFAULT_ROW_ACCESS_WINDOW_SIZE_OF_SXSSFWORKBOOK = -1;
    public static final Integer DEFAULT_ROW_FLUSH_TRIGGER = 1000;
    public static final Integer DEFAULT_ROWS_TO_KEEP_IN_MEMORY_WHILE_FLUSHING = 100;
    public static final String TEMP_FILE_PREFIX = "TempXlsx_";
    public static final String TEMP_FILE_SUFFIX = "";

    //convert csv files to Excel using csv files in local system
    public static void main(String[] args) {
        Map<String, String> inputCsvFilesAndSheetNames = new LinkedHashMap<>();
        inputCsvFilesAndSheetNames.put("sheet1", "test1.csv");
        inputCsvFilesAndSheetNames.put("sheet2", "test2.csv");
        inputCsvFilesAndSheetNames.put("sheet3", "test3.csv");

        buildXlsxReport(inputCsvFilesAndSheetNames);
    }

    private static void buildXlsxReport(Map<String, String> inputCsvFilesAndSheetNames) {
    }
}
