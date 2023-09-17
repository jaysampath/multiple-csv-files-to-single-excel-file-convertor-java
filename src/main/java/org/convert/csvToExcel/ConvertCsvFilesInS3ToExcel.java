package org.convert.csvToExcel;

import com.amazonaws.auth.AWSStaticCredentialsProvider;
import com.amazonaws.auth.BasicSessionCredentials;
import com.amazonaws.services.s3.AmazonS3;
import com.amazonaws.services.s3.AmazonS3ClientBuilder;

import java.util.LinkedHashMap;
import java.util.Map;

public class ConvertCsvFilesInS3ToExcel {
    private static final String ACCESS_KEY = "";
    private static final String SECRET_ACCESS_KEY = "";
    private static final String SESSION_TOKEN = "";
    private static final String sourceCsvFilesS3Prefix = "s3://<bucket-name>/path/to/csv";
    private static final String destinationExcelFileS3Prefix = "s3://<bucket-name>/path/to/excel";

    public static void main(String[] args) {
        BasicSessionCredentials awsCreds = new BasicSessionCredentials(ACCESS_KEY, SECRET_ACCESS_KEY, SESSION_TOKEN);
        AmazonS3 amazonS3 = AmazonS3ClientBuilder.standard()
                .withCredentials(new AWSStaticCredentialsProvider(awsCreds))
                .build();

        ExcelFileGeneratorInS3 excelReportGenerator = new ExcelFileGeneratorInS3(amazonS3);

        Map<String, String> inputCsvFilesAndSheetNames = new LinkedHashMap<>();
        inputCsvFilesAndSheetNames.put("sheet1", sourceCsvFilesS3Prefix + "/" + "test1.csv");
        inputCsvFilesAndSheetNames.put("sheet2", sourceCsvFilesS3Prefix + "/" + "test2.csv");
        inputCsvFilesAndSheetNames.put("sheet3", sourceCsvFilesS3Prefix + "/" + "test3.csv");

        excelReportGenerator.buildXlsxReport(inputCsvFilesAndSheetNames, destinationExcelFileS3Prefix, "converted-excel-file");
    }
}
