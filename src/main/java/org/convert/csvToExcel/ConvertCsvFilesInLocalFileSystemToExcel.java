package org.convert.csvToExcel;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.LinkedHashMap;
import java.util.Map;

public class ConvertCsvFilesInLocalFileSystemToExcel {

    private static final Logger LOGGER = LoggerFactory.getLogger(ConvertCsvFilesInLocalFileSystemToExcel.class);

    private static final char DEFAULT_CSV_FILE_DELIMITER = ',';
    private static final Integer DEFAULT_ROW_ACCESS_WINDOW_SIZE_OF_SXSSFWORKBOOK = -1;
    private static final Integer DEFAULT_ROW_FLUSH_TRIGGER = 1000;
    private static final Integer DEFAULT_ROWS_TO_KEEP_IN_MEMORY_WHILE_FLUSHING = 100;
    private static final String PATH_DELIMITER = "/";

    private static final String sourceCsvFilesLocation = "c:/path/to/source/csv/files";
    private static final String destinationExcelFileLocation = "c:/path/to/destination/excel/file";

    //convert csv files to Excel using csv files in local system
    public static void main(String[] args) throws Exception {
        Map<String, String> inputCsvFilesAndSheetNames = new LinkedHashMap<>();
        inputCsvFilesAndSheetNames.put("sheet1", "test1.csv");
        inputCsvFilesAndSheetNames.put("sheet2", "test2.csv");
        inputCsvFilesAndSheetNames.put("sheet3", "test3.csv");

        buildXlsxReport(inputCsvFilesAndSheetNames);
    }

    private static void buildXlsxReport(Map<String, String> inputCsvFilesAndSheetNames) throws Exception {
        try (SXSSFWorkbook workBook = new SXSSFWorkbook(DEFAULT_ROW_ACCESS_WINDOW_SIZE_OF_SXSSFWORKBOOK)) {
            for (Map.Entry<String, String> entry : inputCsvFilesAndSheetNames.entrySet()) {
                String sheetName = entry.getKey();
                String csvFilePath = sourceCsvFilesLocation + PATH_DELIMITER + entry.getValue();
                if (StringUtils.isNotEmpty(csvFilePath)) {
                    readSingleCsvFileAndAddToWorkbook(workBook, sheetName, csvFilePath);
                } else {
                    //Add blank sheet to workbook
                    workBook.createSheet(sheetName);
                }
            }
            writeWorkbookToDestinationExcelFile(workBook);
        }
    }

    private static void readSingleCsvFileAndAddToWorkbook(SXSSFWorkbook workbook, String sheetName, String sourceFile) throws Exception {
        LOGGER.info("Started reading csv file - {}", sourceFile);
        InputStreamReader inputStreamReader = new InputStreamReader(new FileInputStream(sourceFile));
        CSVParser parser = new CSVParserBuilder().withSeparator(DEFAULT_CSV_FILE_DELIMITER).build();
        CSVReader csvReader = new CSVReaderBuilder(inputStreamReader).withCSVParser(parser).build();
        String[] nextRecord;
        int rowNumber = 0;
        //create sheet in workbook
        SXSSFSheet sheet = workbook.createSheet(sheetName);
        while ((nextRecord = csvReader.readNext()) != null) {
            Row currentRecord = sheet.createRow(rowNumber++);
            for (int colNumber = 0; colNumber < nextRecord.length; colNumber++) {
                currentRecord.createCell(colNumber).setCellValue(nextRecord[colNumber]);
            }
            if (rowNumber % DEFAULT_ROW_FLUSH_TRIGGER == 0) {
                // retain 100 last rows and flush all others to disk
                sheet.flushRows(DEFAULT_ROWS_TO_KEEP_IN_MEMORY_WHILE_FLUSHING);
            }
        }
        csvReader.close();
        inputStreamReader.close();
        LOGGER.info("Completed reading csv file - {} and added to workbook", sourceFile);
    }

    private static void writeWorkbookToDestinationExcelFile(SXSSFWorkbook workBook) {
        LOGGER.info("Started writing workbook contents to xlsx file");
        String destinationExcelFilePath = destinationExcelFileLocation + PATH_DELIMITER + "generated-excel-file.xlsx";
        try(FileOutputStream outputStream = new FileOutputStream(destinationExcelFilePath)) {
            workBook.write(outputStream);
        }
        catch (Exception e) {
            e.printStackTrace();
        }
        LOGGER.info("Successfully built xlsx file- {}", destinationExcelFilePath);
    }
}
