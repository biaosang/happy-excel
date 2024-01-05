package com.github.biaosang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;

public class ExcelFactory {
    public static Workbook createWorkbook(ExcelType excelType){
        return ExcelType.XLSX == excelType ?  new XSSFWorkbook() : new HSSFWorkbook();
    }

    public static Workbook createWorkbookFromFile(ExcelType excelType,String fileName) throws IOException {
        InputStream is = Files.newInputStream(Paths.get(fileName), StandardOpenOption.READ);
        return ExcelType.XLSX == excelType ?  new XSSFWorkbook(is) : new HSSFWorkbook(is);
    }
}
