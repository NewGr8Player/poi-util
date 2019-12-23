package com.xavier.excel.mapping;

/**
 * 枚举值映射
 */
public interface Mapping<Key> {

    String defaultMappingMethodName = "getText";

    /**
     * 获取枚举值的 Text
     *
     * @param key 按照该只进行映射
     * @return 文本值
     */
    String getText(Key key);
}