# 基于注解的可跨行合并单元格导出Excel工具类

## 基于注解导出Excel

对于要导出的数据只需要实现`com.xavier.excel.entity.BasicExportModel`接口即可
+ BasicExportModel#getUniqueKey()的返回值要求全局唯一
+ BasicExportModel#setIndexNo(String)最好创建一个变量用于存储序号
+ BasicExportModel#getIndexNo()取序号用的方法

支持使用字典项并建立映射关系，只需实现`com.xavier.excel.mapping.Mapping`接口即可
+ Mapping<Key>#getText(Key)使用了泛型，但复杂判断逻辑需要在实现类中自行编写

**使用请参考单元测试**
+ `com.xavier.excel.util.ExcelUtilTest#createExcel`
+ `com.xavier.excel.util.ExcelUtilTest#createExcelWithMultiSheets`

> 有问题和意见欢迎提issue
    