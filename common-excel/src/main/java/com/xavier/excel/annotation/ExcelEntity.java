package com.xavier.excel.annotation;

import java.lang.annotation.*;

/**
 * Excel导出文件名(不包含后缀名)
 *
 * @author NewGr8Player
 */
@Inherited
@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelEntity {

    /**
     * 文件名前缀
     */
    String prefix() default "";

    /**
     * 文件名后缀
     */
    String suffix() default "#{date}";
}
