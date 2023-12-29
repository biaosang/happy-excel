package com.github.biaosang;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardOpenOption;
import java.util.ArrayList;
import java.util.List;

public class Excel {

    private Workbook workbook;

    private ExcelType excelType;
    private String fileName;

    private List<Sheet> sheets = new ArrayList<>();

    private Sheet currentSheet = null;

    public Excel(String fileName){
        excel(fileName,ExcelType.XLSX);
    }

    public Excel(String fileName,ExcelType excelType){
        excel(fileName,excelType);
    }

    private void excel(String fileName,ExcelType excelType){
        this.excelType = excelType;
        this.fileName = fileName;
        this.workbook = ExcelFactory.createWorkbook(excelType);
    }
    public <T> Excel addSheet(String sheetName,Class<T> clazz,List<T> data){
        return addSheet(sheetName,clazz,data,false,0);
    }
    public <T> Excel addSheet(String sheetName,Class<T> clazz,List<T> data,boolean ignoreHeader,int startRow){
        return createSheet(sheetName,clazz,data,ignoreHeader,startRow);
    }

    private <T> Excel createSheet(String sheetName,Class<T> clazz,List<T> data,boolean ignoreHeader,int startRow){
        Sheet sheet = workbook.createSheet(sheetName);
        currentSheet = sheet;
        sheets.add(sheet);
        try {
            new XlsxDataHandler(sheet).run(clazz,data,ignoreHeader,startRow);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return this;
    }

    public void export() throws IOException {
        workbook.write(Files.newOutputStream(Paths.get(fileName), StandardOpenOption.CREATE));
    }

}
