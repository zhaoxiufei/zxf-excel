package com.github.zhaoxiufei.excel;

import com.github.zhaoxiufei.excel.annotation.ExcelField;
import com.github.zhaoxiufei.excel.annotation.ExcelSheet;
import com.github.zhaoxiufei.excel.enums.FieldType;
import com.github.zhaoxiufei.excel.utils.Reflections;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

/**
 * 导出Excel文件（导出“XLSX”格式，支持大数据量导出   @see org.apache.poi.ss.SpreadsheetVersion）
 *
 * @author 赵秀非 E-mail:zhaoxiufei@gmail.com
 * @version 创建时间：2017/12/14 15:56
 */
public class ExcelExport {

    private static Logger logger = LoggerFactory.getLogger(ExcelExport.class);
    private static String TITLE = "title";
    private static String HEADER = "header";
    private static String FILE_SUFFIX = ".xlsx";


    /**
     * 工作薄对象
     */
    private SXSSFWorkbook wb;

    /**
     * 工作表对象
     */
    private SXSSFSheet sheet;

    /**
     * 当前行号
     */
    private int rowNum;

    /**
     * excelSheet
     */
    private ExcelSheet excelSheet;
    /**
     * 注解列表
     */
    private List<ExcelModel> excelModels = new ArrayList<>();
    /**
     * 样式缓存
     */
    private Map<Field, CellStyle> cacheStyles = new HashMap<>();

    /**
     * 构造函数
     *
     * @param cls 实体对象
     */
    public ExcelExport(Class<?> cls) {
        excelSheet = cls.getAnnotation(ExcelSheet.class);
        if (excelSheet == null) {
            throw new RuntimeException("无效的ExcelSheet类");
        }
        //字段
        Field[] fs = cls.getDeclaredFields();
        for (Field f : fs) {
            ExcelField ef = f.getAnnotation(ExcelField.class);
            if (ef != null && (FieldType.ALL == ef.type() || FieldType.EXPORT == ef.type())) {
                excelModels.add(new ExcelModel(f, ef));
            }
        }
        //排序
        Collections.sort(excelModels, new Comparator<ExcelModel>() {
            @Override
            public int compare(ExcelModel o1, ExcelModel o2) {
                return o1.getExcelField().sort() - o2.getExcelField().sort();
            }
        });
        initialize();
    }

    /**
     * 初始化函数
     */
    private void initialize() {
        wb = new SXSSFWorkbook(500);//内存行数
        sheet = wb.createSheet();

        Map<String, CellStyle> styles = createStyles(wb);
        //构建表格标题
        if (!"".equals(excelSheet.title())) {
            Row titleRow = addRow();
            titleRow.setHeightInPoints(30);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellStyle(styles.get(TITLE));
            titleCell.setCellValue(excelSheet.title());
            if (excelModels.size() > 1) {
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, excelModels.size() - 1));
            }
        }
        //构建表头
        Row headerRow = addRow();
        headerRow.setHeightInPoints(20);
        for (int i = 0; i < excelModels.size(); i++) {
            ExcelField excelField = excelModels.get(i).getExcelField();
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(styles.get(HEADER));
            cell.setCellValue(excelField.name());
            if (excelField.width() == 9) {
                //自动列宽
                sheet.trackAllColumnsForAutoSizing();
                sheet.autoSizeColumn(i);
            } else {
                sheet.setColumnWidth(i, 256 * excelField.width() + 184);
            }
        }
    }

    /**
     * 创建表格样式
     *
     * @param wb 工作薄对象
     * @return 样式列表
     */
    private Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();

        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);//垂直居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);//水平居中
        Font titleFont = wb.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(Boolean.TRUE);
        style.setFont(titleFont);
        styles.put(TITLE, style);

        style = wb.createCellStyle();
        style.cloneStyleFrom(styles.get(TITLE));
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);
        headerFont.setBold(Boolean.TRUE);
        style.setFont(headerFont);
        styles.put(HEADER, style);

        return styles;
    }

    /**
     * 添加一行
     *
     * @return Row
     */
    private Row addRow() {
        return sheet.createRow(rowNum++);
    }

    /**
     * 添加一个单元格
     *
     * @param row    添加的行
     * @param column 添加列号
     * @param val    添加值
     * @return Cell 单元格对象
     */
    private Cell addCell(Row row, int column, Object val, ExcelModel excelModel) {
        Cell cell = row.createCell(column);
        ExcelField excelField = excelModel.getExcelField();
        CellStyle style = cacheStyles.get(excelModel.getField());
        if (style == null) {
            style = createStyle(excelField.align());
            cacheStyles.put(excelModel.getField(), style);
        }
        cell.setCellStyle(style);
        if (val == null) {
            cell.setCellValue("");
            return cell;
        }
        if (!"".equals(excelField.format())) {
            DataFormat format = wb.createDataFormat();
            style.setDataFormat(format.getFormat(excelField.format()));
        }
        if (val instanceof String) {
            cell.setCellValue((String) val);
        } else if (val instanceof Integer) {
            cell.setCellValue((Integer) val);
        } else if (val instanceof Long) {
            cell.setCellValue((Long) val);
        } else if (val instanceof Double) {
            cell.setCellValue((Double) val);
        } else if (val instanceof Float) {
            cell.setCellValue((Float) val);
        } else if (val instanceof Date) {
            cell.setCellValue((Date) val);
        }
        return cell;
    }

    /**
     * 样式
     *
     * @param align 对齐方式（1：靠左；2：居中；3：靠右）
     * @return CellStyle CellStyle
     */
    private CellStyle createStyle(HorizontalAlignment align) {
        CellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);//垂直居中
        style.setAlignment(align);
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        style.setFont(dataFont);
        return style;
    }

    /**
     * 添加数据
     *
     * @param data 数据
     * @return ExcelExport ExcelExport
     */
    public ExcelExport setData(List<?> data) {
        for (Object o : data) {
            Row row = addRow();
            for (int j = 0; j < excelModels.size(); ) {
                ExcelModel excelModel = excelModels.get(j);
                Object val = Reflections.invokeGetter(o, excelModel.getField().getName());
                //添加表格
                addCell(row, j++, val, excelModel);
            }
        }
        return this;
    }

    /**
     * 输出数据流
     *
     * @param os 输出数据流
     * @throws IOException IOException
     */
    private void write(OutputStream os) throws IOException {
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        BufferedOutputStream bos = null;
        BufferedInputStream bis = null;
        try {
            wb.write(baos);
            bos = new BufferedOutputStream(os);
            bis = new BufferedInputStream(new ByteArrayInputStream(baos.toByteArray()));
            byte[] buff = new byte[2048];
            int bytesRead;
            while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
                bos.write(buff, 0, bytesRead);
            }
        } catch (Exception e) {
            logger.error("write error,msg={}", e.getMessage());
            throw new RuntimeException(e);
        } finally {
            try {
                if (bis != null) {
                    bis.close();
                }
            } finally {
                if (bos != null) {
                    bos.close();
                }
                dispose();
            }
        }
    }

    /**
     * 输出到客户端,文件名为:ExcelSheet.title()
     *
     * @param response HttpServletResponse
     * @throws IOException IOException
     */
    public void write(HttpServletResponse response) throws IOException {
        write(response, null);
    }

    /**
     * 输出到客户端
     *
     * @param fileName 输出文件名
     * @param response HttpServletResponse
     * @throws IOException IOException
     */
    public void write(HttpServletResponse response, String fileName) throws IOException {
        response.reset();
        response.setContentType("application/octet-stream; charset=utf-8");
        response.setHeader("Content-Disposition", "attachment; filename=" +
                URLEncoder.encode(fileName == null ? excelSheet.title() + FILE_SUFFIX : fileName, "UTF-8"));
        write(response.getOutputStream());
    }

    /**
     * 输出到文件中,文件名为:ExcelSheet.title()
     *
     * @param path 输出路径
     * @throws IOException IOException
     */
    public void write(String path) throws IOException {
        String fileName = (path.endsWith("/") ? path + excelSheet.title() : path + "/" + excelSheet.title()) + FILE_SUFFIX;
        FileOutputStream os = new FileOutputStream(fileName);
        write(os);
    }

    /**
     * 输出到文件,
     *
     * @param path     输出路径
     * @param fileName 输出文件名
     * @throws IOException IOException
     */
    public void write(String path, String fileName) throws IOException {
        FileOutputStream os = new FileOutputStream(path + (fileName.contains(".") ? fileName : fileName + FILE_SUFFIX));
        write(os);
    }

    /**
     * 清理临时文件
     */
    private void dispose() {
        wb.dispose();
    }
}
