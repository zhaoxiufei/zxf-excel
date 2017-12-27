package com.github.zhaoxiufei.excel.test;

import com.github.zhaoxiufei.excel.annotation.ExcelSheet;
import com.github.zhaoxiufei.excel.annotation.ExcelField;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.Date;

/**
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/14 15:56
 */
@ExcelSheet(title = "用户数据列表")
public class User {
    @ExcelField(name = "用户ID", align = HorizontalAlignment.RIGHT)
    private Long id;
    @ExcelField(name = "用户名称")
    private String userName;
    @ExcelField(name = "创建时间", width = 18, format = "yyy-mm-dd hh:mm:ss")
    private Date createdTime;
    @ExcelField(name = "类型", align = HorizontalAlignment.LEFT)
    private Integer type;
    @ExcelField(name = "价格", format = "#,##0.00")
    private Float money;

    public Long getId() {
        return id;
    }

    public void setId(Long id) {
        this.id = id;
    }

    public String getUserName() {
        return userName;
    }

    public void setUserName(String userName) {
        this.userName = userName;
    }

    public Date getCreatedTime() {
        return createdTime;
    }

    public void setCreatedTime(Date createdTime) {
        this.createdTime = createdTime;
    }

    public Integer getType() {
        return type;
    }

    public void setType(Integer type) {
        this.type = type;
    }

    public Float getMoney() {
        return money;
    }

    public void setMoney(Float money) {
        this.money = money;
    }
}
