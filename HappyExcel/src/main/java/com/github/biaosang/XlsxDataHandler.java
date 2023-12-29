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

public class XlsxDataHandler {

    private Sheet sheet;

    private int currentRow;

    public XlsxDataHandler(Sheet sheet){
        this.sheet = sheet;
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

    protected void setCellValue(Row row,int col,Object value,String format){
        if(value == null){
            row.createCell(col).setCellValue("");
            return ;
        }

        if(value instanceof Integer || value instanceof Double || value instanceof Float || value instanceof Long){
            row.createCell(col).setCellValue(Double.parseDouble(String.valueOf(value)));
        }else if(value instanceof Boolean){
            row.createCell(col).setCellValue((Boolean)value);
        }else if(value instanceof Date){
            row.createCell(col).setCellValue(new SimpleDateFormat(format).format((Date)value));
        }else if(value instanceof Calendar){
            row.createCell(col).setCellValue(new SimpleDateFormat(format).format(((Calendar)value).getTime()));
        }else if(value instanceof LocalDate){
            row.createCell(col).setCellValue(((LocalDate)value).format(DateTimeFormatter.ofPattern(format)));
        }else if(value instanceof LocalDateTime){
            row.createCell(col).setCellValue(((LocalDateTime)value).format(DateTimeFormatter.ofPattern(format)));
        }else if(value instanceof RichTextString){
            row.createCell(col).setCellValue(((RichTextString)value).getString());
        }else{
            row.createCell(col).setCellValue((String)value);
        }

    }
}
