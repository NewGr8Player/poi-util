package com.xavier.excel.util;

import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.util.Locale.ENGLISH;

/**
 * 格式化模板串
 * eg: namedFormat("Temp#{index}" , Map{"index","001"} )  -> {@code Temp001}
 *
 * @author NewGr8Player
 */
public class NamedFormatUtil {

    /**
     * get方法前缀
     */
    public static final String GETTER_PREFIX = "get";

    /**
     * set方法前缀
     */
    public static final String SETTER_PREFIX = "set";

    /**
     * Format regx.
     */
    private final static Pattern namedFormatPattern = Pattern.compile("#\\{(?<key>.*?)}");

    /**
     * Named-String Format By regx.
     *
     * @param format format
     * @param kvs    key-value pairs
     * @return
     */
    public static String namedFormat(final String format, Map<String, ? extends Object> kvs) {
        final StringBuffer buffer = new StringBuffer();
        final Matcher match = namedFormatPattern.matcher(format);
        while (match.find()) {
            final String key = match.group("key");
            final Object value = kvs.get(key);
            if (value != null) {
                match.appendReplacement(buffer, value.toString());
            } else if (kvs.containsKey(key)) {
                match.appendReplacement(buffer, "null");
            } else {
                match.appendReplacement(buffer, "");
            }
        }
        match.appendTail(buffer);
        return buffer.toString();
    }

    /**
     * 首字母转大写
     *
     * @param name 字段名
     * @return
     */
    public static String capitalize(String name) {
        if (name == null || name.length() == 0) {
            return name;
        }
        return name.substring(0, 1).toUpperCase(ENGLISH) + name.substring(1);
    }
}
