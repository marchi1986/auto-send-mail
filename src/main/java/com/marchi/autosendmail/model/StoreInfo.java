package com.marchi.autosendmail.model;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import lombok.Data;


@Data
@ColumnWidth(30)
public class StoreInfo {

    @ExcelProperty("商户号")
    private String merchantCode;
    @ExcelProperty("商户门店号")
    private String merchantStoreCode;
    @ExcelProperty("门店号")
    private String storeCode;
    @ExcelProperty("终端名称")
    private String storeName;
    @ExcelProperty("手机号")
    private String mobile;
    @ExcelIgnore
    private String area;
}
