package com.xavier.excel.mapping;

import java.util.HashMap;
import java.util.Map;

/**
 * 学生性别值映射为文本
 * 此处仅为示例，定义为枚举也可以
 *
 * @author NewGr8Player
 */
public class StudentSexMapping implements Mapping<String> {

    private static Map<String, String> sexMapping = new HashMap<>();

    private static final String UNMAPPED_TEXT = "未知";

    static {
        sexMapping.put("male", "男");
        sexMapping.put("female", "女");
    }

    @Override
    public String getText(String key) {
        return sexMapping.getOrDefault(key, UNMAPPED_TEXT);
    }
}
