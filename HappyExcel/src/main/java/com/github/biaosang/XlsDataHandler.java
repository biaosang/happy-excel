package com.github.biaosang;

import org.apache.poi.ss.usermodel.Sheet;

public class XlsDataHandler extends DataHandler {
    public XlsDataHandler(Sheet sheet) {
        super(sheet);
        super.excelType = ExcelType.XLS;
    }
}
