# multiple-csv-files-to-single-excel-file-convertor-using-java

Convert multiple csv files to a single Excel(XLSX) file using apache-poi library

### Memory optimized approach
* Used `SXSSFWorkbook` from apache-poi library which uses temp files to load large csv files in order to preserve runtime memory.
* Any csv file/records are not loaded in Java runtime memory.
* Cleaned up resources used by `SXSSFWorkbook` using `dispose()` method

### ConvertCsvFilesInLocalFileSystemToExcel
* Run this class to convert csv files from local system and convert to an Excel file in local system

### ConvertCsvFilesInS3ToExcel
* Run this class to convert csv files from S3 folder and convert to an Excel file in S3 folder.
* Amazon S3's `PutObjectRequest` requires `InputStream` object, but `SXSSFWorkbook` can only write to an `OutputStream` object.
* Used a Temp file to convert `OutputStream` to `InputStream` and uploaded the `InputStream` to a file in S3.