package com.github.biaosang.test;

import com.github.biaosang.Excel;
import com.github.biaosang.ExcelField;
import com.github.biaosang.ExcelType;
import com.github.biaosang.HappyExcel;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

@HappyExcel
public class TypeTest {

    @ExcelField(header = "int",col = 0)
    private int intField;

    @ExcelField(header = "Integer",col = 1)
    private Integer intPackField;

    @ExcelField(header = "string",col = 2)
    private String stringField;

    @ExcelField(header = "double",col = 3)
    private double doubleField;

    @ExcelField(header = "Double",col = 4)
    private Double doublePackField;

    @ExcelField(header = "float",col = 5)
    private float floatField;

    @ExcelField(header = "Float",col = 6)
    private Float floatPackField;

    @ExcelField(header = "Date",col = 7)
    private Date dateField;

    @ExcelField(header = "Calendar",col = 8)
    private Calendar calendarField;

    @ExcelField(header = "LocalDate",col = 9,format = "yyyy-MM-dd")
    private LocalDate localDateField;

    @ExcelField(header = "LocalDateTime",col = 10)
    private LocalDateTime localDateTimeField;

    @ExcelField(header = "RichTextString",col = 11)
    private RichTextString richTextString;

    public TypeTest() {
    }

    public TypeTest(int intField, Integer intPackField, String stringField, double doubleField, Double doublePackField, float floatField, Float floatPackField, Date dateField, Calendar calendarField, LocalDate localDateField, LocalDateTime localDateTimeField, RichTextString richTextString) {
        this.intField = intField;
        this.intPackField = intPackField;
        this.stringField = stringField;
        this.doubleField = doubleField;
        this.doublePackField = doublePackField;
        this.floatField = floatField;
        this.floatPackField = floatPackField;
        this.dateField = dateField;
        this.calendarField = calendarField;
        this.localDateField = localDateField;
        this.localDateTimeField = localDateTimeField;
        this.richTextString = richTextString;
    }


    @Override
    public String toString() {
        return "TypeTest{" +
                "intField=" + intField +
                ", intPackField=" + intPackField +
                ", stringField='" + stringField + '\'' +
                ", doubleField=" + doubleField +
                ", doublePackField=" + doublePackField +
                ", floatField=" + floatField +
                ", floatPackField=" + floatPackField +
                ", dateField=" + dateField +
                ", calendarField=" + calendarField +
                ", localDateField=" + localDateField +
                ", localDateTimeField=" + localDateTimeField +
                ", richTextString=" + richTextString +
                '}';
    }

    public static void main(String[] args) throws Exception {
        TypeTest typeTest = new TypeTest(1,2,"3",4.12,5.12,6.13f,7.13f,
                new Date(),Calendar.getInstance(),LocalDate.now(),LocalDateTime.now(),new XSSFRichTextString("123"));

        List<TypeTest> typeTestList = new ArrayList<>();
        typeTestList.add(typeTest);
        new Excel("typelist.xlsx", ExcelType.XLSX)
                .addSheet("type test", TypeTest.class,typeTestList)
                .export();
        System.out.println("导出字段类型测试完成");

        List<TypeTest> typeTestList2 = new ArrayList<>();
        new Excel("typelist.xlsx", ExcelType.XLSX)
                .importSheet(0,TypeTest.class,typeTestList2,1);
        System.out.println(typeTestList2);
    }

}
