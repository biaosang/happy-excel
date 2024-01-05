package com.github.biaosang.test;

import com.github.biaosang.Excel;
import com.github.biaosang.ExcelType;

import java.util.ArrayList;
import java.util.List;

public class ImportTest {
    public static void main(String[] args) {
        List<User> users = new ArrayList<>();
        try {
            new Excel("content.xlsx", ExcelType.XLSX)
                    .importSheet(0,User.class,users,1);
            System.out.println(users);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
