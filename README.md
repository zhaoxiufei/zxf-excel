# zxf-excel
基于poi实现的Excel导入导出组件,使用方便,一行代码即可搞定!

## 一、介绍
### 1.1 导出
 可直接导出到文件,InputStream或WEB端下载,格式为2007~2013及更高版本
 功能如下:
 1. 支持导出字段排序,自然排序或自定义排序
 2. 支持单元格列宽自适应和自定义
 3. 支持单元格格式化(表达式详细参见Excel设置单元格格式->自定义列表)
 4. 支持自定义单元格对齐方式
### 1.2 导入
 将Excel文件转换为List<E>,支持97~2013及更高本导入
 
### 1.3 运行环境
 JDK:1.7+
 
## 二、入门

### 2.1 引入maven依赖

```
<dependency>
    <groupId>com.github.zhaoxiufei</groupId>
    <artifactId>zxf-excel</artifactId>
    <version>1.0.0</version>
</dependency>
```

### 2.2 定义实体对象
```java
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
    getter/setter...
}
```

### 2.3 导入示例

```
//导入示例：ExcelTest

new ExcelExport(User.class).setData(objects).write("E:\\", "测试导出.xlsx");
```

### 2.4 导出示例

```
//导出示例：ExcelTest

List<User> objects = new ExcelImport("E:\\测试导出.xlsx",1).getData(User.class);
```
## 三、技术支持
 1. 使用中遇到问题,请新建[Issue](https://github.com/zhaoxiufei/zxf-excel/issues/)
 2. 如需及时帮助,请[加入QQ群](https://shang.qq.com/wpa/qunwpa?idkey=487f8cff0d1f4bce4e44c6a7626dc807cc9a1508889469778678fcd999ebf6d2)
 3. 也可以给我,zhaoxiufei@gmail.com
 
## 四、赞助
 如果你觉得有帮助到你,可以请开发者喝咖啡!<br/>
 
### 微信

![咖啡](https://github.com/zhaoxiufei/zxf-excel/blob/master/images/wx.png)
 
### 支付宝

![咖啡](https://github.com/zhaoxiufei/zxf-excel/blob/master/images/alipay.png)
