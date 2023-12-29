package com.github.biaosang;

public enum ExcelType {
    XLS(".xls"),
    XLSX("xlsx"),
    ;

    private String value;

    ExcelType(String value) {
        this.value = value;
    }

    public String getValue() {
        return value;
    }
}
