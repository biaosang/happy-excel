package com.github.biaosang;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class DataHandler {

    protected Sheet sheet;

    protected int currentRow;

    protected ExcelType excelType;

    public DataHandler(Sheet sheet) {
        this.sheet = sheet;
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

    protected  <T> void loadData(Class<T> clazz, List<T> outputData,int startRow) throws Exception {
        HappyExcel happyExcel = clazz.getAnnotation(HappyExcel.class);
        if(happyExcel == null)
            throw new HappyExcelException(clazz.getName() + " 类上没有@HappyExcel注解");
        currentRow = startRow;

        while(true){
            Row row = sheet.getRow(currentRow++);
            if(row == null)
                break;

            T t = clazz.newInstance();
            outputData.add(t);
            Field[] fields = clazz.getDeclaredFields();
            for(Field field : fields){
                ExcelField excelField = field.getAnnotation(ExcelField.class);
                if(excelField == null)
                    continue;

                field.setAccessible(true);

                setCellValue(field,row,t,excelField);
            }
        }
    }

    private <T> void setCellValue(Field field, Row row, T t, ExcelField excelField) throws Exception {

        Class<?> type = field.getType();

        if(type == Integer.class || type == int.class){
            double value = row.getCell(excelField.col()).getNumericCellValue();
            field.set(t,(int)value);
        }else if(type == Double.class || type == double.class){
            double value = row.getCell(excelField.col()).getNumericCellValue();
            field.set(t,value);
        }else if(type == Float.class || type == float.class){
            double value = row.getCell(excelField.col()).getNumericCellValue();
            field.set(t,Float.parseFloat(String.valueOf(value)));
        }else if(type == Long.class || type == long.class){
            double value = row.getCell(excelField.col()).getNumericCellValue();
            field.set(t,(long)value);
        }else if(type == Boolean.class){
            boolean value = row.getCell(excelField.col()).getBooleanCellValue();
            field.set(t,value);
        }else if(type == Date.class){
            String value = row.getCell(excelField.col()).getStringCellValue();
            field.set(t,new SimpleDateFormat(excelField.format()).parse(value));
        }else if(type == Calendar.class){
            String value = row.getCell(excelField.col()).getStringCellValue();
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(new SimpleDateFormat(excelField.format()).parse(value));
            field.set(t,calendar);
        }else if(type == LocalDate.class){
            String value = row.getCell(excelField.col()).getStringCellValue();
            field.set(t, LocalDate.parse(value,DateTimeFormatter.ofPattern(excelField.format())));
        }else if(type == LocalDateTime.class){
            String value = row.getCell(excelField.col()).getStringCellValue();
            field.set(t,LocalDateTime.parse(value,DateTimeFormatter.ofPattern(excelField.format())));
        }else if(type == RichTextString.class){
            RichTextString richTextString =  row.getCell(excelField.col()).getRichStringCellValue();
            field.set(t,richTextString);
        }else{
            field.set(t,row.getCell(excelField.col()).getStringCellValue());
        }
    }

}
