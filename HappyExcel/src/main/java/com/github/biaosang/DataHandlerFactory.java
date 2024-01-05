package com.github.biaosang;

import org.apache.poi.ss.usermodel.Sheet;

public class DataHandlerFactory {

    public static DataHandler create(ExcelType excelType, Sheet sheet){
        return ExcelType.XLSX == excelType ? new XlsxDataHandler(sheet) : new XlsDataHandler(sheet);
    }
}
