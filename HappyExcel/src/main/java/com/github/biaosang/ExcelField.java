package com.github.biaosang;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelField {

    String header () default "";

    int col();

    String format() default "yyyy-MM-dd HH:mm:ss";

}
