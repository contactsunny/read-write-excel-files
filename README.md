# Read/Write Microsoft Excel Files

This is a simple Java POC to read and write Microsoft Excel files. 

## How to run

The program accepts three input parameters from the command line:

1. File path
2. Sheet name
3. Data to be written

An example command is as follows:

```shell
java -jar ReadWriteExcelFilePoc-1.0-SNAPSHOT.jar /home/sunny/Documents/testExcel.xls TestSheet TestingData
```

## How it works

If the file is not found at the provided path, a new file will be created and that data will be written at cell 0,0. If a file is found, the data will be appended at the end of the sheet after creating a new row.
