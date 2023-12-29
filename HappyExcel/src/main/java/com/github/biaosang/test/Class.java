package com.github.biaosang.test;

import com.github.biaosang.ExcelField;
import com.github.biaosang.HappyExcel;

@HappyExcel
public class Class {

    @ExcelField(header = "班级",col = 0)
    private String name;

    @ExcelField(header = "年级",col = 1)
    private String level;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getLevel() {
        return level;
    }

    public void setLevel(String level) {
        this.level = level;
    }

    public Class(String name, String level) {
        this.name = name;
        this.level = level;
    }
}
