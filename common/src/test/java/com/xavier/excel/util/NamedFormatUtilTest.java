package com.xavier.excel.util;

import org.junit.Assert;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.JUnit4;

import java.util.HashMap;
import java.util.Map;

@RunWith(JUnit4.class)
public class NamedFormatUtilTest {

    @Test
    public void namedFormat() {
        final String format = "Test#{varI}_#{varI}__#{varII}";
        Map<String, String> map = new HashMap<>();
        map.put("varI", "I");
        map.put("varII", "II");
        Assert.assertEquals("TestI_I__II", NamedFormatUtil.namedFormat(format, map));
    }

    @Test
    public void capitalize() {
        Assert.assertEquals("FieldName", NamedFormatUtil.capitalize("fieldName"));
    }
}