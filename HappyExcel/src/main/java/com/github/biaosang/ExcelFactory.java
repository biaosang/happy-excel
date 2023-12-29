package com.github.biaosang;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFactory {
    public static Workbook createWorkbook(ExcelType excelType){
        return ExcelType.XLSX == excelType ?  new XSSFWorkbook() : new HSSFWorkbook();
    }
}
