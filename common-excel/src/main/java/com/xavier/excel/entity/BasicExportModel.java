package com.xavier.excel.entity;

import java.io.Serializable;

/**
 * 导出基础模型接口
 *
 * @author NewGr8Player
 */
public interface BasicExportModel extends Serializable {

    /**
     * 数据合并分组依据此方法返回值
     *
     * @return
     */
    String getUniqueKey();

    /**
     * 获取序号
     *
     * @return
     */
    String getIndexNo();

    /**
     * 设置序号
     *
     * @param indexNo 序号
     */
    void setIndexNo(String indexNo);
}
