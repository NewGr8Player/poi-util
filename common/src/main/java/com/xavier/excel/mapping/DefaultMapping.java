package com.xavier.excel.mapping;

/**
 * 仅作为默认值使用
 */
public final class DefaultMapping implements Mapping<Object> {
    @Override
    public String getText(Object s) {
        throw new UnsupportedOperationException("请自行编写实现类！");
    }
}
