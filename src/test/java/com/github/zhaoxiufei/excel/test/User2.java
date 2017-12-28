package com.github.zhaoxiufei.excel.test;

import com.github.zhaoxiufei.excel.annotation.ExcelField;
import com.github.zhaoxiufei.excel.annotation.ExcelSheet;

/**
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/15 15:56
 */
@ExcelSheet(title = "用户数据列表2")
public class User2 {
    @ExcelField(name = "用户ID")
    private Long id;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }
}
