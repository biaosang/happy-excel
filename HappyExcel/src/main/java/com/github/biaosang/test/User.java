package com.github.biaosang.test;

import com.github.biaosang.Excel;
import com.github.biaosang.ExcelField;
import com.github.biaosang.ExcelType;
import com.github.biaosang.HappyExcel;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

@HappyExcel
public class User {

    @ExcelField(header = "姓名",col = 0)
    private String name;

    @ExcelField(header = "年龄",col = 1)
    private int age;

    @ExcelField(header = "生日",col = 2)
    private Date birth;

    public User(){}

    public User(String name, int age, Date birth) {
        this.name = name;
        this.age = age;
        this.birth = birth;
    }

    public User(String name, int age) {
        this.name = name;
        this.age = age;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    @Override
    public String toString() {
        return "User{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", birth=" + birth +
                '}';
    }

    public static void main(String[] args) throws IOException {
        List<User> users = new ArrayList<>();
        users.add(new User("张三",11,new Date()));
        users.add(new User("李四",17,new Date()));
        users.add(new User("王五",21,new Date()));
        new Excel("user.xlsx", ExcelType.XLSX)
                .addSheet("用户", User.class,users)
                .export();
        System.out.println("导出用户测试完成");
    }

}
