package com.xavier.excel.annotation;

import java.lang.annotation.*;

/**
 * 可合并表格字段
 *
 * @author NewGr8Player
 */
@Inherited
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    /**
     * 导出字段标题
     */
    String fieldTitle() default "";

    /**
     * 单元格宽度
     */
    int cellWidth() default 0;

    /**
     * 是否参与合并部分
     */
    boolean isParent() default false;

    /**
     * 导出顺序(非单元格位置)
     */
    int index() default 0;
}
