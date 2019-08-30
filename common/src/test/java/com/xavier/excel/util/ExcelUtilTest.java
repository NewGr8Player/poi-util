package com.xavier.excel.util;

import com.github.jsonzou.jmockdata.DataConfig;
import com.github.jsonzou.jmockdata.JMockData;
import com.github.jsonzou.jmockdata.MockConfig;
import com.github.jsonzou.jmockdata.TypeReference;
import org.apache.commons.math3.analysis.function.Max;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.UUID;

@RunWith(JUnit4.class)
public class ExcelUtilTest {

    public static final int GEN_DATA_NUM = 100;

    private List<Student> studentList2019 = new ArrayList<>();
    private List<Student> studentList2018 = new ArrayList<>();
    private List<Student> studentList2017 = new ArrayList<>();
    private List<Student> studentList2016 = new ArrayList<>();
    private List<Student> studentList2015 = new ArrayList<>();
    private List<Student> studentList2014 = new ArrayList<>();

    @Before
    public void initTestData() {
        studentList2019 = initData("2014", GEN_DATA_NUM);
        studentList2018 = initData("2015", 66);
        studentList2017 = initData("2016", GEN_DATA_NUM);
        studentList2016 = initData("2017", 50);
        studentList2015 = initData("2018", 79);
        studentList2014 = initData("2019", 91);
    }


    @Test
    public void createExcel() {
        try {
            OutputStream outputStream = new FileOutputStream(new File("D:\\test.xlsx"));
            ExcelUtil.createXlsxExcel(studentList2019, "2019年学生表", Student.class).write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void createExcelWithMultiSheets() {
        try {
            OutputStream outputStream = new FileOutputStream(new File("D:\\test.xlsx"));
            Workbook workbook = new XSSFWorkbook();
            ExcelUtil.createExcel(studentList2019, "studentList2019", workbook, Student.class);
            ExcelUtil.createExcel(studentList2018, "studentList2018", workbook, Student.class);
            ExcelUtil.createExcel(studentList2017, "studentList2017", workbook, Student.class);
            ExcelUtil.createExcel(studentList2016, "studentList2016", workbook, Student.class);
            ExcelUtil.createExcel(studentList2015, "studentList2015", workbook, Student.class);
            ExcelUtil.createExcel(studentList2014, "studentList2014", workbook, Student.class);
            workbook.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private String genStuNo(String prefix) {
        return prefix + JMockData.mock(String.class, new MockConfig()
                .stringSeed("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")
                .sizeRange(8, 8)
        );
    }

    private List<Student> initData(String prefix, int genNum) {
        List<Student> studentList = new ArrayList<>();
        int temp = 0;
        String uuid = genStuNo(prefix);
        for (int i = 0; i < genNum; i++) {
            if (i / 10 != temp) {
                temp = i / 10;
                uuid = genStuNo(prefix);
            }
            Student student = JMockData.mock(new TypeReference<Student>() {
                                             }, new MockConfig()
                            .globalConfig()
                            .subConfig(Student.class, "stuName") /* 姓名 */
                            .stringSeed("赵", "钱", "孙", "李", "周", "吴", "郑", "王")
                            .sizeRange(2, 3)
                            .subConfig(Student.class, "comment") /* 备注 */
                            .stringSeed("优秀", "很棒", "评价", "哈哈哈哈")
                            .sizeRange(1, 1)
                            .subConfig(Student.class, "sex")
                            .stringSeed("男", "女")
                            .sizeRange(1, 1)
                            .subConfig(Student.class, "score") /* 分数 */
                            .doubleRange(0.0, 100.0)
                            .decimalScale(1)
                            .globalConfig()
                            .excludes("stuNo", "indexNo")
            );
            student.setStuNo(uuid);
            studentList.add(student);
        }
        return studentList;
    }
}