package com.github.biaosang;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class XlsxDataHandler extends DataHandler{

    public XlsxDataHandler(Sheet sheet){
        super(sheet);
        super.excelType = ExcelType.XLSX;
    }

    public <T> void run(Class<T> clazz, List<T> data, boolean ignoreHeader, int startRow) throws Exception {
        HappyExcel happyExcel = clazz.getAnnotation(HappyExcel.class);
        if(happyExcel == null)
            throw new HappyExcelException("数据类缺少'@HappyExcel'注解标识");

        currentRow = startRow;
        if(!ignoreHeader){
            addHeaderRow(clazz);
        }
        addData(clazz,data);
    }

    private <T> void addData(Class<T> clazz, List<T> data) throws Exception {
        Field[] fields =clazz.getDeclaredFields();
        for(T t : data){
            Row dataRow = sheet.createRow(currentRow++);
            for(Field field : fields) {
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                if (excelField != null) {
                    field.setAccessible(true);
                    Object object = field.get(t);
                    setCellValue(dataRow,excelField.col(),object,excelField.format());
                }
            }
        }
    }

    private <T> void addHeaderRow(Class<T> clazz) throws Exception {
        Row headerRow = sheet.createRow(currentRow++);
        //colSet 用于校验列是否重复
        Set<Integer> colSet = new HashSet<>();

        Field[] fields =clazz.getDeclaredFields();
        for(Field field : fields){
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if(excelField != null){
                if(colSet.contains(excelField.col())){
                    throw new HappyExcelException("class " + clazz.getName() +" 发现列序号配置重复");
                }
                colSet.add(excelField.col());

                String header = excelField.header();
                if(header == null){
                    header = field.getName();
                }
                headerRow.createCell(excelField.col()).setCellValue(header);
            }
        }
    }



    @Override
    protected <T> void loadData(Class<T> clazz, List<T> outputData, int startRow) throws Exception {
        super.loadData(clazz, outputData, startRow);
    }
}
