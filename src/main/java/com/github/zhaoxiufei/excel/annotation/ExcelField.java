
package com.github.zhaoxiufei.excel.annotation;

import com.github.zhaoxiufei.excel.enums.FieldType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/14 15:56
 */
@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    /**
     * 导出字段标题
     */
    String name();

    /**
     * 导出字段字段排序（升序）
     */
    int sort() default 0;

    /**
     * 列宽
     */
    int width() default 9;

    /**
     * 格式化
     */
    String format() default "";

    /**
     * 字段类型（0：导出导入；1：仅导出；2：仅导入）
     */
    FieldType type() default FieldType.ALL;

    /**
     * 导出字段对齐方式,默认:水平居中
     */
    HorizontalAlignment align() default HorizontalAlignment.CENTER;
}
