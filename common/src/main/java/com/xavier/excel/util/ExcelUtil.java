package com.xavier.excel.util;

import com.xavier.excel.annotation.ExcelEntity;
import com.xavier.excel.annotation.ExcelField;
import com.xavier.excel.entity.BasicExportModel;
import com.xavier.excel.entity.ExcelFieldModel;
import com.xavier.excel.mapping.DefaultMapping;
import com.xavier.excel.mapping.Mapping;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.commons.math3.exception.OutOfRangeException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @author fengzm
 * @version 2019-8-22
 */
@Slf4j
public class ExcelUtil {

    private static final String DEFAULT_FONT_NAME = "宋体";

    private static final String DEFAULT_FILE_NAME = "导出文件";

    private static final short DEFAULT_TITLE_FONT_HEIGHT_IN_POINTS = 10;
    private static final short DEFAULT_CONTENT_FONT_HEIGHT_IN_POINTS = 9;

    private static final int DEFAULT_BYTE_LENGTH_FIXED_WEIGHT = 256;

    private static final int MAX_CELL_WIDTH = 65280; /* 255 * 256 */

    private static final String DEFAULT_DATE_FIELD_NAME = "date";

    /**
     * excel 2007+ suffix
     */
    private static final String XLSX_SUFFIX = ".xlsx";

    /**
     * excel 2003 suffix
     */
    private static final String XLS_SUFFIX = ".xls";

    /**
     * 根据注解获取注解信息List
     *
     * @param clazz 类类型
     * @return 注解信息List
     */
    private static List<ExcelFieldModel> getFieldEntity(Class<?> clazz) {
        List<ExcelFieldModel> excelFieldEntities = new ArrayList<>();

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelField excelField = field.getAnnotation(ExcelField.class);
            if (Objects.nonNull(excelField)) {
                try {
                    String fieldName = field.getName();
                    excelFieldEntities.add(
                            new ExcelFieldModel().setFieldTitle(excelField.fieldTitle())
                                    .setExportIndex(excelField.index())
                                    .setParent(excelField.isParent())
                                    .setFieldName(fieldName)
                                    .setCellWidth(excelField.cellWidth())
                                    .setFieldSetter(clazz.getMethod(NamedFormatUtil.SETTER_PREFIX + NamedFormatUtil.capitalize(fieldName), field.getType()))
                                    .setFieldGetter(clazz.getMethod(NamedFormatUtil.GETTER_PREFIX + NamedFormatUtil.capitalize(fieldName)))
                                    .setValueMapping(excelField.mapping())
                    );
                } catch (NoSuchMethodException e) {
                    log.error("获取数据导出Entity出错", e);
                }
            }
        }
        excelFieldEntities.sort(
                Comparator.comparing(ExcelFieldModel::getExportIndex)
        );
        return excelFieldEntities;
    }

    /**
     * 获取文件名
     *
     * @param workBook         工作表格
     * @param excelEntityClazz 承载导出信息Entity的类类型
     * @param prefixMap        文件名前缀参数Map
     * @param suffixMap        文件名后缀参数Map
     * @return 文件名
     */
    private static String getFileNameWithSuffix(Workbook workBook, Class<?> excelEntityClazz, Map<String, Object> prefixMap, Map<String, Object> suffixMap) {
        String fileName = DEFAULT_FILE_NAME;
        try {
            final Map<String, Object> defaultParamMap = new HashMap<>();
            defaultParamMap.put(DEFAULT_DATE_FIELD_NAME, DateFormatUtils.format(new Date(), "yyyyMMddHHmmss"));

            ExcelEntity excelField = excelEntityClazz.getAnnotation(ExcelEntity.class);

            prefixMap = Optional.ofNullable(prefixMap)
                    .orElse(new HashMap<>());
            prefixMap.putAll(defaultParamMap);
            suffixMap = Optional.ofNullable(suffixMap)
                    .orElse(new HashMap<>());
            suffixMap.putAll(defaultParamMap);
            fileName = NamedFormatUtil.namedFormat(excelField.prefix(), prefixMap)
                    + NamedFormatUtil.namedFormat(excelField.suffix(), suffixMap);
            if (StringUtils.isBlank(fileName)) {
                fileName = UUID.randomUUID().toString();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        if (Objects.nonNull(workBook)) { /* 后缀名 */
            fileName += workBook instanceof XSSFWorkbook ? XLSX_SUFFIX : XLS_SUFFIX;
        }
        return fileName;
    }

    /**
     * 获取不带后缀名的文件名
     *
     * @param excelEntityClazz 注解类的类类型
     * @param prefixMap        文件名前缀参数Map
     * @param suffixMap        文件名后缀参数Map
     * @return 无后缀的文件名
     */
    private static String getFileName(Class<?> excelEntityClazz, Map<String, Object> prefixMap, Map<String, Object> suffixMap) {
        return getFileNameWithSuffix(null, excelEntityClazz, prefixMap, suffixMap);
    }

    /**
     * 创建表格(Word 2007+)
     *
     * @param dataList         表格数据
     * @param currentSheetName 当前sheet页名称
     * @param clazz            承载导出信息Entity的类类型
     * @return Workbook对象
     */
    public static Workbook createXlsxExcel(List<? extends BasicExportModel> dataList, String currentSheetName, Class<?> clazz) {
        return createExcel(dataList, currentSheetName, new XSSFWorkbook(), clazz);
    }

    /**
     * 创建表格(Word 2003)
     *
     * @param dataList         表格数据
     * @param currentSheetName 当前sheet页名称
     * @param clazz            承载导出信息Entity的类类型
     * @return Workbook对象
     */
    public static <T extends BasicExportModel> Workbook createXlsExcel(List<T> dataList, String currentSheetName, Class<?> clazz) {
        return createExcel(dataList, currentSheetName, new HSSFWorkbook(), clazz);
    }

    /**
     * 创建表格(Word 2007+)
     *
     * @param dataList         表格数据
     * @param currentSheetName 当前sheet页名称
     * @param clazz            承载导出信息Entity的类类型
     * @param outputStream     输出流
     * @return 有后缀文件名
     */
    public static <T extends BasicExportModel> String createXlsxExcel(List<T> dataList, String currentSheetName, Class<?> clazz, OutputStream outputStream) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        createExcel(dataList, currentSheetName, workbook, clazz).write(outputStream);
        return getFileNameWithSuffix(workbook, clazz, null, null);
    }

    /**
     * 创建表格(Word 2003)
     *
     * @param dataList         表格数据
     * @param currentSheetName 当前sheet页名称
     * @param clazz            承载导出信息Entity的类类型
     * @param outputStream     输出流
     * @return 有后缀文件名
     */
    public static <T extends BasicExportModel> String createXlsExcel(List<T> dataList, String currentSheetName, Class<?> clazz, OutputStream outputStream) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        createExcel(dataList, currentSheetName, workbook, clazz).write(outputStream);
        return getFileNameWithSuffix(workbook, clazz, null, null);
    }

    /**
     * 创建表格
     *
     * @param dataList         表格数据
     * @param currentSheetName 当前sheet页名称
     * @param workbook         Excel-POI实例
     * @param clazz            承载导出信息Entity的类类型
     * @return Workbook对象
     */
    public static <T extends BasicExportModel> Workbook createExcel(List<T> dataList, String currentSheetName, Workbook workbook, Class<?> clazz) {
        Sheet sheet = workbook.createSheet(currentSheetName);

        List<ExcelFieldModel> fieldEntityList = getFieldEntity(clazz);

        /* 标题行 */
        Row titleRow = sheet.createRow(0);
        CellStyle headerStyle = headerStyle(workbook);
        int titleColumnLength = fieldEntityList.size();
        for (int i = 0; i < titleColumnLength; i++) {
            String title = fieldEntityList.get(i).getFieldTitle();
            int width = Integer.min(
                    Integer.max(
                            title.length() * 2
                            , fieldEntityList.get(i).getCellWidth()
                    ) * DEFAULT_BYTE_LENGTH_FIXED_WEIGHT, MAX_CELL_WIDTH);
            sheet.setColumnWidth(i, width);
            Cell cell = titleRow.createCell(i);
            cell.setCellStyle(headerStyle);
            cell.setCellValue(title);
        }

        CellStyle contentStyle = contentStyle(workbook);
        int dataLength = dataList.size();
        Map<String, List<T>> groupedMap = dataList.stream()
                .collect(Collectors.groupingBy(T::getUniqueKey));
        for (int rowIndex = 1; rowIndex <= dataLength; ) {
            int seqNum = 0;
            for (List<T> v : groupedMap.values()) {
                int mergeSize = v.size();
                boolean mergeNeededFlag = mergeSize > 1;
                if (mergeNeededFlag || 1 == mergeSize) {
                    seqNum++; /* 序号 */
                }
                for (T tab : v) {
                    Row currentRow = sheet.createRow(rowIndex);
                    tab.setIndexNo(String.valueOf(seqNum)); /* 给序号赋值 */
                    int columnSize = fieldEntityList.size();
                    for (int j = 0; j < columnSize; j++) {
                        ExcelFieldModel fieldEntity = fieldEntityList.get(j);
                        try {
                            Object cellObject = fieldEntity.getFieldGetter().invoke(tab);
                            String cellText = Objects.equals(fieldEntity.getValueMapping(), DefaultMapping.class)
                                    ? String.valueOf(cellObject)
                                    : valueMappingHandler(fieldEntity.getValueMapping(),cellObject);
                            if (mergeNeededFlag && fieldEntity.isParent()) {
                                cellValueFiller(currentRow, j, contentStyle, cellText, mergeSize, sheet);
                            } else {
                                cellValueFiller(currentRow, j, contentStyle, cellText);
                            }
                        } catch (IllegalAccessException | InvocationTargetException e) {
                            e.printStackTrace();
                            log.error(String.format("行数据赋值出错，数据Id{}", tab.getUniqueKey()), e);
                        }
                    }
                    mergeNeededFlag = false;
                    rowIndex++;
                }
            }

        }
        return workbook;
    }

    /**
     * 单元格填入值
     *
     * @param row         行对象
     * @param columnIndex 列索引
     * @param cellStyle   单元格样式
     * @param cellValue   单元格值
     */
    private static void cellValueFiller(Row row, int columnIndex, CellStyle cellStyle, String cellValue) {
        Cell currentCell = row.createCell(columnIndex);
        currentCell.setCellStyle(cellStyle);
        currentCell.setCellValue(cellValue);
    }

    /**
     * 合并单元格并填入值
     *
     * @param row         行对象
     * @param columnIndex 列索引
     * @param cellStyle   单元格样式
     * @param cellValue   单元格值
     * @param mergeSize   合并大小
     * @param sheet       sheet对象实例
     */
    private static void cellValueFiller(Row row, int columnIndex, CellStyle cellStyle, String cellValue, int mergeSize, Sheet sheet) {
        Cell currentCell = row.createCell(columnIndex);
        currentCell.setCellStyle(cellStyle);
        currentCell.setCellValue(cellValue);
        int rowNum = row.getRowNum();
        mergeCell(rowNum, rowNum + mergeSize - 1, columnIndex, columnIndex, sheet);
    }

    /**
     * 单元格样式
     *
     * @param workbook Excel-Workbook对象
     * @return 单元格样式实例
     */
    protected static CellStyle headerStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(Font.COLOR_NORMAL);
        font.setFontName(DEFAULT_FONT_NAME);
        font.setBold(true);
        font.setFontHeightInPoints(DEFAULT_TITLE_FONT_HEIGHT_IN_POINTS);

        commonCellBorderStyle(style, BorderStyle.MEDIUM.getCode(), IndexedColors.BLACK.index);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        return style;
    }

    /**
     * 内容样式
     *
     * @param workbook Excel-Workbook对象
     * @return 单元格样式实例
     */
    protected static CellStyle contentStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setColor(Font.COLOR_NORMAL);
        font.setFontName(DEFAULT_FONT_NAME);
        font.setFontHeightInPoints(DEFAULT_CONTENT_FONT_HEIGHT_IN_POINTS);

        commonCellBorderStyle(style, BorderStyle.THIN.getCode(), IndexedColors.GREY_50_PERCENT.index);

        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);

        return style;
    }

    /**
     * 通用单元格边框样式生成
     *
     * @param style           样式对象
     * @param borderStyleCode 边框样式代码
     * @param borderColor     边框颜色代码
     */
    protected static void commonCellBorderStyle(CellStyle style, short borderStyleCode, short borderColor) {
        BorderStyle borderStyle = BorderStyle.valueOf(borderStyleCode);

        style.setBorderTop(borderStyle);
        style.setBorderBottom(borderStyle);
        style.setBorderLeft(borderStyle);
        style.setBorderRight(borderStyle);

        style.setTopBorderColor(borderColor);
        style.setBottomBorderColor(borderColor);
        style.setLeftBorderColor(borderColor);
        style.setRightBorderColor(borderColor);
    }

    /**
     * 合并单元格对象
     *
     * @param firstRow 合并起始行
     * @param lastRow  合并结束行
     * @param firstCol 合并开始列
     * @param lastCol  合并结束列
     * @param sheet    sheet页对象
     */
    protected static void mergeCell(int firstRow, int lastRow, int firstCol, int lastCol, Sheet sheet) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);

        BorderStyle borderStyle = BorderStyle.valueOf(BorderStyle.THIN.getCode());
        RegionUtil.setBorderTop(borderStyle, cellRangeAddress, sheet);
        RegionUtil.setBorderBottom(borderStyle, cellRangeAddress, sheet);
        RegionUtil.setBorderLeft(borderStyle, cellRangeAddress, sheet);
        RegionUtil.setBorderRight(borderStyle, cellRangeAddress, sheet);

        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.index, cellRangeAddress, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.index, cellRangeAddress, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.index, cellRangeAddress, sheet);
        RegionUtil.setBottomBorderColor(IndexedColors.BLACK.index, cellRangeAddress, sheet);
        sheet.addMergedRegion(cellRangeAddress);
    }

    /**
     * @param mappingClazz 映射实现类的类类型(Class Type)
     * @param key          代码值
     * @param <Key>        代码值的类型
     * @return 显示的文本
     */
    protected static <Key> String valueMappingHandler(Class<? extends Mapping<?>> mappingClazz, Key key) {
        String resultText = "";
        try {
            resultText = (String) mappingClazz.getMethod(Mapping.defaultMappingMethodName, key.getClass())
                    .invoke(mappingClazz.isEnum()
                                    ? mappingClazz.getEnumConstants()[0]
                                    : mappingClazz.newInstance()
                            , key);
        } catch (OutOfRangeException ore) {
            log.error("枚举异常，请确认实现后枚举内至少有一个枚举值。", ore);
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException | InstantiationException e) {
            log.error("值映射实现类方法获取异常", e);
        }
        return resultText;
    }
}
