package com.xavier.excel.mapping;

import lombok.extern.slf4j.Slf4j;

import java.util.Objects;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * 测试
 */
@Slf4j
public enum StudentSexEnumMapping implements Mapping<String> {
    MALE("male", "男"), FEMALE("female", "女");
    private String value;
    private String text;
    private static final String UNMAPPED_TEXT = "未知";

    StudentSexEnumMapping(String value, String text) {
        this.value = value;
        this.text = text;
    }

    @Override
    public String getText(String value) {
        try {
            Optional<StudentSexEnumMapping> result = Stream.of(StudentSexEnumMapping.values()).filter(
                    current -> Objects.equals(current.value, value)
            ).findFirst();
            return result.isPresent()
                    ? result.get().text
                    : UNMAPPED_TEXT;
        } catch (Exception ex) {
            log.warn("未能正确映射:" + value, ex);
            return UNMAPPED_TEXT;
        }
    }
}
