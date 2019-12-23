package com.xavier.excel.entity;

import com.xavier.excel.mapping.Mapping;
import lombok.Getter;
import lombok.Setter;
import lombok.ToString;
import lombok.experimental.Accessors;

import java.lang.reflect.Method;

/**
 * 导出字段信息
 *
 * @author NewGr8Player
 */
@Getter
@Setter
@Accessors(chain = true)
@ToString
public class ExcelFieldModel {

    /**
     * 字段名
     */
    private String fieldName;
    /**
     * 导出字段标题
     */
    private String fieldTitle;

    /**
     * 是否参与合并部分
     */
    private boolean isParent;

    /**
     * 导出顺序
     */
    private int exportIndex;

    /**
     * 单元格宽度
     */
    private int cellWidth;

    /**
     * get方法
     */
    private Method fieldGetter;

    /**
     * set方法
     */
    private Method fieldSetter;

    /**
     * 值映射为文本信息
     */
    private Class<? extends Mapping<?>> valueMapping;
}
