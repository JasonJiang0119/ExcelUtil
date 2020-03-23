package com.jason.mylog.utils;

import com.alibaba.fastjson.JSON;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @author Jq
 * @Title ExcelUtil
 * @Description
 * @date 2020/3/19 17:01:24
 */
@Slf4j
public class ExcelUtil {
    private static final String APPLICATION_HEADER = "application/vnd.ms-excel;charset=UTF-8";

    private static final String CONTENT_DISPOSITION = "Content-disposition";

    private static final String SUFFIX = ".xlsx";

    private static final String FILE_NAME_EXCHANGE_ORDER = "attachment;filename=";

    private static final String STRING_CONSTANT = "String";

    private static final String INTEGER_CONSTANT = "Integer";

    private static final String INT_CONSTANT = "int";

    private static final String LONG_CONSTANT = "Long";

    private static final String LONG_CONSTANT_SMALL = "long";

    private static final String DOUBLE_CONSTANT = "Double";

    private static final String DOUBLE_CONSTANT_SMALL = "double";

    private static final String DATE_CONSTANT = "Date";

    private static final String LIST_CONSTANT = "List";

    private static final String COLLECTION_CONSTANT = "Collection";

    private static final String SET_CONSTANT = "HashSet";

    private ExcelUtil(){

    }

    public static <T> void export(final String filenNamePrefix,
                                  final String[] rowName,
                                  final List<T> data,
                                  final String sheetName,
                                  final HttpServletResponse response) {
        BufferedOutputStream bos = null;
        try {
            final SimpleDateFormat sdf=new SimpleDateFormat("yyyyMMddHHmmss");
            final String fileName = filenNamePrefix + sdf.format(new Date()) + String.format("%04d", new Random().nextInt(10000)) + SUFFIX;
            bos = getBufferedOutputStream(fileName, response);
            doExport(rowName,sheetName,data,bos);
        }catch (final Exception e){
            log.error("Export Excel error : {}",e.getMessage());
        }finally {
            try {
                if (bos != null) {
                    bos.close();
                }
            } catch (IOException e) {
                log.error("Export Excel error : {}",e.getMessage());
            }
        }

    }

    /**
     * 从excel中读内容
     */
    public static <T> List<T> readExcel(final MultipartFile file,final Class<T> cls) {
        XSSFWorkbook workBook = null;
        List<T> list = new ArrayList<>();
        try (InputStream inputStream = file.getInputStream()){
            //读取工作簿
            workBook = new XSSFWorkbook(inputStream);
            list = doRead(workBook,cls);
        }catch (final Exception e){
            log.error("Read MultipartFile Excel error : {}",e.getMessage());
        }finally {
            if(workBook != null){
                //关闭工作簿
                try {
                    workBook.close();
                } catch (IOException e) {
                    log.error("Read MultipartFile Excel error : {}",e.getMessage());
                }
            }
        }
        return list;
    }

    /**
     * 从excel中读内容
     */
    public static <T> List<T> readExcel(final File file,final Class<T> cls) {
        XSSFWorkbook workBook = null;
        List<T> list = new ArrayList<>();
        try (InputStream inputStream = new FileInputStream(file)){
            //读取工作簿
            workBook = new XSSFWorkbook(inputStream);
            list = doRead(workBook,cls);
        }catch (final Exception e){
            log.error("Read File Excel error : {}",e.getMessage());
        }finally {
            if(workBook != null){
                //关闭工作簿
                try {
                    workBook.close();
                } catch (IOException e) {
                    log.error("Read File Excel error : {}",e.getMessage());
                }
            }
        }
        return list;
    }


    private static <T> List<T> doRead(final XSSFWorkbook wb,final Class<T> cls){
        final XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;
        final int lastRowNum = sheet.getLastRowNum();
        // 循环读取
        final List<T> lists = new ArrayList<>();
        Field[] fields = cls.getDeclaredFields();
        for (int i = 1; i <= lastRowNum; i++) {
            row = sheet.getRow(i);
            if(null != row){
                final Map<String,Object> map = new HashMap<>(fields.length);
                for(int y= 0; y < fields.length;y++){
                    map.put(fields[y].getName(),getCellValue(row.getCell(y)));
                }
                lists.add(JSON.parseObject(JSON.toJSON(map).toString(),cls));
            }
        }
        return lists;
    }

    private static String getCellValue(final Cell cell)
    {
        if (cell == null)
        {
            log.info("cell is null,return null");
            return "";
        }
        String value = null;

        switch (cell.getCellType()) {
            case _NONE:
            case BLANK:
                value = "";
                break;
            case NUMERIC:
                value = StringUtils.trim(new DecimalFormat("##").format(cell.getNumericCellValue()));
                break;
            case BOOLEAN:
                value = (cell.getBooleanCellValue() ? "TRUE":"FALSE");
                break;
            case STRING:
                value = StringUtils.trim(cell.getStringCellValue());
                break;
            default:
                break;
        }

        return value;
    }

    private static BufferedOutputStream getBufferedOutputStream(final String fileName, final HttpServletResponse response) throws IOException {
        response.setContentType(APPLICATION_HEADER);
        response.setHeader(CONTENT_DISPOSITION, FILE_NAME_EXCHANGE_ORDER + new String(fileName.getBytes("gb2312"), StandardCharsets.ISO_8859_1));
        return new BufferedOutputStream(response.getOutputStream());
    }

    private static <T> void doExport(final String[] headers,
                                     final String sheetName,
                                     final List<T> data,
                                     final OutputStream outputStream) throws IOException, NoSuchFieldException, IllegalAccessException {
        final SXSSFWorkbook workbook = new SXSSFWorkbook();
        createSheet(workbook,headers,data,sheetName);
        if (outputStream != null) {
            workbook.write(outputStream);
        }

    }

    private static <T> void createSheet(final SXSSFWorkbook wb,
                                        final String[] headers,
                                        final List<T> dataList,
                                        final String sheetName) throws NoSuchFieldException, IllegalAccessException {

        // 创建一张工作表
        final SXSSFSheet sheet = wb.createSheet(sheetName);

        final CellStyle style = wb.createCellStyle();
        final CellStyle style2 = wb.createCellStyle();
        //创建表头
        final Font font = wb.createFont();
        font.setFontName("微软雅黑");
        //设置字体大小
        font.setFontHeightInPoints((short) 11);
        //选择需要用到的字体格式
        style.setFont(font);

        // 设置背景色
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.PALE_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 居中
        style.setAlignment(HorizontalAlignment.CENTER);
        //下边框
        style.setBorderBottom(BorderStyle.THIN);
        //右边框
        style.setBorderRight(BorderStyle.THIN);

        //选择需要用到的字体格式
        style2.setFont(font);

        // 设置背景色
        style2.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //垂直居中
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        // 水平向下居中
        style2.setAlignment(HorizontalAlignment.CENTER);
        //下边框
        style2.setBorderBottom(BorderStyle.THIN);
        //右边框
        style2.setBorderRight(BorderStyle.THIN);
        //左边框
        style2.setBorderLeft(BorderStyle.THIN);
        //上边框
        style2.setBorderTop(BorderStyle.THIN);

        //表头
        final Row headerRow = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellStyle(style);
            sheet.setColumnWidth(i, 4000);
            cell.setCellValue(headers[i]);
        }

        int rownum = 0;
        for(final T data : dataList){
            final Row row = sheet.createRow(rownum + 1);

            final Field[] fields = getExportFields(data.getClass());
            for(int cellnum = 0;cellnum < fields.length;cellnum++){
                final Field field = fields[cellnum];
                final Cell cell = row.createCell(cellnum);
                cell.setCellStyle(style2);
                setData(field, data, field.getName(), cell,row);
            }
            rownum = sheet.getLastRowNum();
        }
    }

    private static Field[] getExportFields(final Class<?> targetClass) {
        return targetClass.getDeclaredFields();
    }

    /**
     * 根据属性设置对应的属性值
     *
     * @param dataField 属性
     * @param object    数据对象
     * @param property  表头的属性映射
     * @param cell      单元格
     */
    private static <T> void setData(final Field dataField,
                                    final T object,
                                    final String property,
                                    final Cell cell,
                                    final Row row)
            throws IllegalAccessException, NoSuchFieldException {
        //允许访问private属性
        dataField.setAccessible(true);
        //获取属性值
        Object val = dataField.get(object);
        //获取单元格样式
        final CellStyle style = cell.getCellStyle();
        int cellnum = cell.getColumnIndex();
        if (null != val) {
            dataSet(dataField,val,cell,row,cellnum,style,property);
        }
    }

    private static void dataSet(final Field dataField,
                                final Object val,
                                final Cell cell,
                                final Row row,
                                int cellnum,
                                final CellStyle style,
                                final String property) throws NoSuchFieldException, IllegalAccessException {
        if (dataField.getType().toString().endsWith(STRING_CONSTANT)
                || dataField.getType().toString().endsWith(INTEGER_CONSTANT)
                || dataField.getType().toString().endsWith(INT_CONSTANT)
                || dataField.getType().toString().endsWith(LONG_CONSTANT)
                || dataField.getType().toString().endsWith(LONG_CONSTANT_SMALL)
                || dataField.getType().toString().endsWith(DOUBLE_CONSTANT)
                || dataField.getType().toString().endsWith(DOUBLE_CONSTANT_SMALL)) {
            cell.setCellValue(String.valueOf(val));
        } else if (dataField.getType().toString().endsWith(DATE_CONSTANT)) {
            final DateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            cell.setCellValue(format.format((Date) val));
        } else if (dataField.getType().toString().endsWith(LIST_CONSTANT) || dataField.getType().toString().endsWith(COLLECTION_CONSTANT) || dataField.getType().toString().endsWith(SET_CONSTANT)) {
            listSet(cellnum,val,row,style,cell);
        } else {
            final String str = ".";
            if (property.contains(str)) {
                final String p = property.substring(property.indexOf(str) + 1);
                final Field field = getDataField(val, p);
                setData(field, val, p, cell,row);
            } else {
                cell.setCellValue(val.toString());
            }
        }
    }

    @SuppressWarnings("unchecked")
    private static <T> void listSet(final int cellnum,
                                    final Object val,
                                    final Row row,
                                    final CellStyle style,
                                    final Cell cell) throws NoSuchFieldException, IllegalAccessException {
        //适用于list平铺模板
        int listCell = cellnum;
        final Collection<T> list = (Collection<T>) val;
        for (Object o : list) {
            Field[] listFields = getExportFields(o.getClass());
            for (final Field listField : listFields) {
                Cell cellList = row.createCell(listCell);
                cellList.setCellStyle(style);
                cell.setCellStyle(style);
                setData(listField, o, listField.getName(), cellList, row);
                listCell = listCell + 1;
            }
        }
    }

    /**
     * 获取单条数据的属性
     */
    private static <T> Field getDataField(final T object, final String property) throws NoSuchFieldException {
        Field dataField;
        final String str = ".";
        if (property.contains(str)) {
            final String p = property.substring(0, property.indexOf(str));
            dataField = object.getClass().getDeclaredField(p);
            return dataField;
        } else {
            dataField = object.getClass().getDeclaredField(property);
        }
        return dataField;
    }

}
