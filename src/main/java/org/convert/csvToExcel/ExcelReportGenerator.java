package org.convert.csvToExcel;

import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3URI;
import com.amazonaws.services.s3.model.ObjectMetadata;
import com.amazonaws.services.s3.model.PutObjectRequest;
import com.amazonaws.services.s3.model.S3Object;
import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvValidationException;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalTime;
import java.util.Map;


public class ExcelReportGenerator {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelReportGenerator.class);
    public static final char DEFAULT_CSV_FILE_DELIMITER = ',';
    public static final String XLSX_EXTENSION = ".xlsx";
    public static final String XLSX_FILE_FORMAT_CONTENT_TYPE =
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    public static final Integer DEFAULT_ROW_ACCESS_WINDOW_SIZE_OF_SXSSFWORKBOOK = -1;
    public static final Integer DEFAULT_ROW_FLUSH_TRIGGER = 1000;
    public static final Integer DEFAULT_ROWS_TO_KEEP_IN_MEMORY_WHILE_FLUSHING = 100;
    public static final String TEMP_FILE_PREFIX = "TempXlsx_";
    public static final String TEMP_FILE_SUFFIX = "";
    private final AmazonS3 amazonS3Client;

    public ExcelReportGenerator(AmazonS3 amazonS3) {
        this.amazonS3Client = amazonS3;
    }

    public void buildXlsxReport(Map<String, String> inputCsvFilesAndSheetNames, String s3BasePath, String destinationFileName) {
        String destinationXlsxFilePath = s3BasePath + destinationFileName + XLSX_EXTENSION;
        LOGGER.info("Building xlsx file started for- {}, ", inputCsvFilesAndSheetNames);
        long xlsxReportBuildStartTime = System.currentTimeMillis();
        try (SXSSFWorkbook workBook = new SXSSFWorkbook(DEFAULT_ROW_ACCESS_WINDOW_SIZE_OF_SXSSFWORKBOOK)) {
            for (Map.Entry<String, String> entry : inputCsvFilesAndSheetNames.entrySet()) {
                String queryName = entry.getKey();
                String reportFilePath = entry.getValue();
                if (StringUtils.isNotEmpty(reportFilePath)) {
                    readSingleCsvFileAndAddToWorkbook(workBook, queryName, reportFilePath);
                } else {
                    //Add blank sheet to workbook
                    workBook.createSheet(queryName);
                }
            }
            writeWorkbookToDestinationFileInS3(workBook, destinationXlsxFilePath);
            String reportBuildTime = calculateTimeTakenForXlsxReportBuild(System.currentTimeMillis(), xlsxReportBuildStartTime);
            LOGGER.info("Successfully completed building xlsx file, total time taken- {}", reportBuildTime);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private void readSingleCsvFileAndAddToWorkbook(SXSSFWorkbook workbook, String sheetName, String sourceFile) throws IOException, CsvValidationException {
        LOGGER.info("Started reading csv file - {} from S3", sourceFile);
        S3Object s3Object = getS3ObjectContent(sourceFile);
        InputStreamReader inputStreamReader = new InputStreamReader(s3Object.getObjectContent());
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
        LOGGER.info("Completed reading csv file - {} from S3 and added to workbook", sourceFile);
    }

    private void writeWorkbookToDestinationFileInS3(SXSSFWorkbook workbook, String destinationFile) throws IOException {

        LOGGER.info("Started writing contents to xlsx file- {} in S3", destinationFile);
        //create temp file
        File tempFile = File.createTempFile(TEMP_FILE_PREFIX, TEMP_FILE_SUFFIX);

        try (FileInputStream fileInputStream = new FileInputStream(tempFile);
             FileOutputStream fileOutputStream = new FileOutputStream(tempFile)) {
            //write to temp file
            workbook.write(fileOutputStream);
            LOGGER.info("Intermediate operation: Written workbook to OutputStream and OutputStream to temp file- {}\n"
                    + "Started writing contents of temp file to a xlsx file in S3", tempFile.getAbsolutePath());
            //upload stream to S3
            writeByteStreamToFileInS3(destinationFile, fileInputStream);
        } finally {
            try {
                //delete the temp file used to upload to S3
                Files.delete(Paths.get(tempFile.getPath()));
                LOGGER.info("Intermediate operation: Deleted the temp file from disk- {}", tempFile.getPath());
            } catch (Exception e) {
                //Rare scenario. just log it.
                LOGGER.error("Intermediate operation: Something went wrong while deleting the temp file- {} for request id- {}",
                        tempFile.getPath(), e);
            }
            //deletes the temporary files used by SXSSFWorkbook
            if (workbook.dispose()) {
                LOGGER.info("Intermediate operation: Disposed the workbook to clean up temp resources");
            } else {
                //Rare scenario. just log it.
                LOGGER.error("Intermediate operation: Something went wrong while disposing the workbook}");
            }
        }
    }

    private String calculateTimeTakenForXlsxReportBuild(long end, long start) {
        long reportBuildTimeInSeconds = ((end - start) / 1000L);
        return LocalTime.MIN.plusSeconds(reportBuildTimeInSeconds).toString();
    }

    private S3Object getS3ObjectContent(String fileLocation) {
        String bucketName = new AmazonS3URI(fileLocation).getBucket();
        String key = new AmazonS3URI(fileLocation).getKey();
        return amazonS3Client.getObject(bucketName, key);
    }

    private void writeByteStreamToFileInS3(String fileDestination, InputStream inputStream) {
        String bucketName = new AmazonS3URI(fileDestination).getBucket();
        String key = new AmazonS3URI(fileDestination).getKey();
        ObjectMetadata metadata = new ObjectMetadata();
        metadata.setContentType(XLSX_FILE_FORMAT_CONTENT_TYPE);
        amazonS3Client.putObject(new PutObjectRequest(bucketName, key, inputStream, metadata));
    }

}
