package com.github.zhaoxiufei.excel.test;

import com.github.zhaoxiufei.excel.ExcelExport;
import com.github.zhaoxiufei.excel.ExcelImport;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/19 13:36
 */
public class ExcelTest {
    @Test
    public void exportTest() {
        List<User> objects = new ArrayList<>();
        for (int i = 1; i < 100; i++) {
            User user = new User();
            user.setId((long) i);
            user.setUserName("第" + i + "个人");
            user.setCreatedTime(new Date());
            user.setType(i);
            user.setMoney(Float.parseFloat(0.11245 + ""));
            objects.add(user);
        }
        try {
            new ExcelExport(User.class).setData(objects).write("E:\\", "测试导出.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void importTest() {
        try {
            List<User> objects = new ExcelImport("E:\\测试导出.xlsx", 1).getData(User.class);
            for (User object : objects) {
                System.out.println(object.getId());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
