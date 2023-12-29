package com.github.biaosang.test;

import com.github.biaosang.Excel;
import com.github.biaosang.ExcelType;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Test {

    public static void main(String[] args) throws IOException {
        List<User> users = new ArrayList<>();
        users.add(new User("张三",11,new Date()));
        users.add(new User("李四",17,new Date()));
        users.add(new User("王五",21,new Date()));

        List<Class> classes = new ArrayList<>();
        classes.add(new Class("一班","一年级"));
        classes.add(new Class("二班","一年级"));
        classes.add(new Class("一班","二年级"));
        classes.add(new Class("二班","二年级"));
        classes.add(new Class("三班","二年级"));

        new Excel("content.xlsx", ExcelType.XLSX)
                .addSheet("用户表", User.class,users)
                .addSheet("班级表", Class.class,classes)
                .addSheet("班级表无表头", Class.class,classes,true,0)
                .export();
        System.out.println("导出完成");
    }
}
