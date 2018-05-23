package com.github.ldzzdl.easyexcel4j.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author LDZZDL
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Excel {

    /**
     * Excel的标题名称
     * @return
     */
    String excelTitle() default "";

    /**
     * Excel的标题列号，从1开始
     * @return
     */
    int excelOrder() default 0;

}
