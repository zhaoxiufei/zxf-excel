package com.github.zhaoxiufei.excel;

import com.github.zhaoxiufei.excel.annotation.ExcelField;

import java.lang.reflect.Field;

/**
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/25 14:34
 */
public class ExcelModel {
    private Field field;
    private ExcelField excelField;

    public ExcelModel(Field field, ExcelField excelField) {
        this.field = field;
        this.excelField = excelField;
    }

    public Field getField() {
        return field;
    }

    public ExcelField getExcelField() {
        return excelField;
    }
}
