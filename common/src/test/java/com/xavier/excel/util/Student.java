package com.xavier.excel.util;

import com.xavier.excel.annotation.ExcelField;
import com.xavier.excel.entity.BasicExportModel;
import com.xavier.excel.mapping.StudentSexEnumMapping;
import com.xavier.excel.mapping.StudentSexMapping;
import lombok.*;

import java.math.BigDecimal;

@Setter
@Getter
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class Student implements BasicExportModel {

    @ExcelField(fieldTitle = "序列号", isParent = true, index = -1)
    private String indexNo;

    @ExcelField(fieldTitle = "学号", isParent = true, cellWidth = 20, index = 1)
    private String stuNo;

    @ExcelField(fieldTitle = "姓名", isParent = true, cellWidth = 10, index = 2)
    private String stuName;

    @ExcelField(fieldTitle = "性别", isParent = true, index = 3,mapping = StudentSexEnumMapping.class)
    //@ExcelField(fieldTitle = "性别", isParent = true, index = 3,mapping = StudentSexMapping.class)
    private String sex;

    @ExcelField(fieldTitle = "科目成绩", index = 4)
    private BigDecimal score;

    @ExcelField(fieldTitle = "备注", isParent = true, cellWidth = 100, index = 99)
    private String comment;

    @Override
    public String getUniqueKey() {
        return stuNo;
    }
}
