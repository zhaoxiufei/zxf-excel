
package com.github.zhaoxiufei.excel;

import com.github.zhaoxiufei.excel.enums.FieldType;
import com.github.zhaoxiufei.excel.utils.Reflections;
import com.github.zhaoxiufei.excel.utils.StringUtil;
import com.github.zhaoxiufei.excel.annotation.ExcelField;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * 导入Excel文件（支持XLS和XLSX格式）
 *
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/14 16:56
 */
public class ExcelImport {

    private static Logger log = LoggerFactory.getLogger(ExcelImport.class);

    /**
     * 工作表对象
     */
    private Sheet sheet;

    /**
     * 标题行号
     */
    private int headerNum;


    /**
     * 构造函数
     *
     * @param fileName 导入文件，读取第一个工作表,第一行
     */
    public ExcelImport(String fileName) throws IOException {
        this(fileName, 0);
    }

    /**
     * 构造函数
     *
     * @param fileName  导入文件
     * @param headerNum 标题行号
     */
    public ExcelImport(String fileName, int headerNum) throws IOException {
        this(fileName, headerNum, 0);
    }

    /**
     * 构造函数
     *
     * @param fileName   导入文件
     * @param headerNum  标题行号
     * @param sheetIndex 工作表编号
     */
    public ExcelImport(String fileName, int headerNum, int sheetIndex) throws IOException {
        this(new File(fileName), headerNum, sheetIndex);
    }

    /**
     * 构造函数
     *
     * @param file 导入文件对象，读取第一个工作表,第一行
     */
    public ExcelImport(File file) throws IOException {
        this(file, 0);
    }

    /**
     * 构造函数
     *
     * @param file      导入文件对象
     * @param headerNum 标题行号
     */
    public ExcelImport(File file, int headerNum) throws IOException {
        this(file, headerNum, 0);
    }

    /**
     * 构造函数
     *
     * @param file       导入文件对象
     * @param headerNum  标题行号
     * @param sheetIndex 工作表编号
     */
    public ExcelImport(File file, int headerNum, int sheetIndex) throws IOException {
        this(new BufferedInputStream(new FileInputStream(file)), headerNum, sheetIndex);
    }

    /**
     * 构造函数
     *
     * @param buf 导入文件对象
     */
    public ExcelImport(byte[] buf) throws IOException {
        this(buf, 0);
    }

    /**
     * 构造函数
     *
     * @param buf 导入文件对象
     */
    public ExcelImport(byte[] buf, int headerNum) throws IOException {
        this(buf, headerNum, 0);
    }

    /**
     * 构造函数
     *
     * @param buf 导入文件对象
     */
    public ExcelImport(byte[] buf, int headerNum, int sheetIndex) throws IOException {
        this(new ByteArrayInputStream(buf), headerNum, sheetIndex);
    }

    /**
     * 构造函数
     *
     * @param inputStream 导入文件对象
     */
    public ExcelImport(InputStream inputStream) throws IOException {
        this(inputStream, 0);
    }

    /**
     * 构造函数
     *
     * @param inputStream 导入文件对象
     */
    public ExcelImport(InputStream inputStream, int headerNum) throws IOException {
        this(inputStream, headerNum, 0);
    }

    /**
     * 构造函数
     *
     * @param headerNum  标题行号，数据行号=标题行号+1
     * @param sheetIndex 工作表编号
     */
    public ExcelImport(InputStream is, int headerNum, int sheetIndex) throws IOException {
        if (!is.markSupported()) {
            is = new PushbackInputStream(is, 8);
        }
        /*
      工作薄对象
     */
        Workbook wb;
        if (FileMagic.OLE2 == FileMagic.valueOf(is)) {
            wb = new HSSFWorkbook(is);
        } else if (FileMagic.OOXML == FileMagic.valueOf(is)) {
            wb = new XSSFWorkbook(is);
        } else {
            throw new RuntimeException("文档格式不正确!");
        }
        if (wb.getNumberOfSheets() < sheetIndex) {
            throw new RuntimeException("文档中没有工作表!");
        }
        this.sheet = wb.getSheetAt(sheetIndex);
        this.headerNum = headerNum;
        log.debug("Initialize success.");
    }

    /**
     * 获取行对象
     */
    private Row getRow(int rowNum) {
        return this.sheet.getRow(rowNum);
    }

    /**
     * 获取数据行号
     */
    private int getDataRowNum() {
        return headerNum + 1;
    }

    /**
     * 获取最后一个数据行号
     */
    private int getLastDataRowNum() {
        return this.sheet.getLastRowNum();
    }

    /**
     * 获取最后一个列号
     */
    public int getLastCellNum() {
        return this.getRow(headerNum).getLastCellNum();
    }

    /**
     * 获取单元格值
     *
     * @param row    获取的行
     * @param column 获取单元格列号
     * @return 单元格值
     */
    @SuppressWarnings("deprecation")
    public Object getCellValue(Row row, int column) {
        Object val = "";
        try {
            Cell cell = row.getCell(column);
            if (cell != null) {
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    val = cell.getNumericCellValue();
                } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
                    val = cell.getStringCellValue();
                } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
                    val = cell.getCellFormula();
                } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
                    val = cell.getBooleanCellValue();
                } else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
                    val = cell.getErrorCellValue();
                }
            }
        } catch (Exception e) {
            return val;
        }
        return val;
    }

    /**
     * 获取导入数据列表
     *
     * @param cls 导入对象类型
     */
    public <E> List<E> getData(Class<E> cls) {
        List<ExcelModel> excelModels = new ArrayList<>();
        Field[] fs = cls.getDeclaredFields();
        for (Field f : fs) {
            ExcelField ef = f.getAnnotation(ExcelField.class);
            if (ef != null && (FieldType.ALL == ef.type() || FieldType.IMPORT == ef.type())) {
                excelModels.add(new ExcelModel(f, ef));
            }
        }
        Collections.sort(excelModels, new Comparator<ExcelModel>() {
            @Override
            public int compare(ExcelModel o1, ExcelModel o2) {
                return o1.getExcelField().sort() - o2.getExcelField().sort();
            }
        });
        List<E> dataList = new ArrayList<>();
        for (int i = this.getDataRowNum(); i <= this.getLastDataRowNum(); i++) {
            E e;
            try {
                e = cls.newInstance();
            } catch (InstantiationException | IllegalAccessException ex) {
                throw new RuntimeException(ex);
            }
            int column = 0;
            Row row = this.getRow(i);
            for (ExcelModel em : excelModels) {
                Object val = this.getCellValue(row, column++);
                if (val != null) {
                    Class<?> valType = em.getField().getType();
                    try {
                        if (valType == String.class) {
                            String s = String.valueOf(val.toString());
                            if (s.endsWith(".0")) {
                                val = StringUtil.substringBefore(s, ".0");
                            } else {
                                val = String.valueOf(val.toString());
                            }
                        } else if (valType == Integer.class) {
                            val = Double.valueOf(val.toString()).intValue();
                        } else if (valType == Long.class) {
                            val = Double.valueOf(val.toString()).longValue();
                        } else if (valType == Double.class) {
                            val = Double.valueOf(val.toString());
                        } else if (valType == Float.class) {
                            val = Float.valueOf(val.toString());
                        } else if (valType == Date.class) {
                            val = DateUtil.getJavaDate((Double) val);
                        }
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        log.info("Get cell value [" + i + "," + column + "] error: {}", ex.getMessage());
                        val = null;
                    }
                    Reflections.invokeSetter(e, em.getField().getName(), val);
                }
            }
            dataList.add(e);
        }
        return dataList;
    }
}
