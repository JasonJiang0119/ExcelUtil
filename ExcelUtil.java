package com.xc.common.util;

import com.alibaba.fastjson.JSON;
import lombok.extern.log4j.Log4j2;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
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
@Log4j2
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

    public static final String BOLD = "bold";

    private ExcelUtil(){

    }

    public static <T> void export(final String fileNamePrefix,
                                  final String[] rowName,
                                  final List<T> data,
                                  final String sheetName,
                                  final HttpServletResponse response,
                                  final String titleName,
                                  final List<CellRangeAddress> cellRangeAddressList) {
        BufferedOutputStream bos = null;
        try {
            final SimpleDateFormat sdf=new SimpleDateFormat("yyyyMMddHHmmss");
            final String fileName = fileNamePrefix + sdf.format(new Date()) + String.format("%04d", new Random().nextInt(10000)) + SUFFIX;
            bos = getBufferedOutputStream(fileName, response);
            doExport(rowName,sheetName,data,bos,titleName,cellRangeAddressList);
        }catch (final Exception e){
            log.warn("Export Excel error : {}",e.getMessage());
        }finally {
            try {
                if (bos != null) {
                    bos.close();
                }
            } catch (IOException e) {
                log.warn("Export Excel error :{} ",e.getMessage());
            }
        }

    }

    public static <T> void exportManySheet(final String[] rowName,
                                           final List<T> data,
                                           final String sheetName,
                                           final String titleName,
                                           final SXSSFWorkbook workbook) throws IllegalAccessException, NoSuchFieldException, IOException {
        doManySheetExport(rowName,sheetName,data,titleName,workbook);

    }

    /**
     * 从excel中读内容
     */
    public static <T> List<T> readExcel(final MultipartFile file, final Class<T> cls) {
        XSSFWorkbook workBook = null;
        List<T> list = new ArrayList<>();
        try (InputStream inputStream = file.getInputStream()){
            //读取工作簿
            workBook = new XSSFWorkbook(inputStream);
            list = doRead(workBook,cls);
        }catch (final Exception e){
            log.warn("Read MultipartFile Excel error :{}",e.getMessage());
        }finally {
            if(workBook != null){
                //关闭工作簿
                try {
                    workBook.close();
                } catch (IOException e) {
                    log.warn("Read MultipartFile Excel error : {}",e.getMessage());
                }
            }
        }
        return list;
    }

    /**
     * 从excel中读获取头内容
     */
    public static <T> T readExcelTitle(final MultipartFile file, final Class<T> cls) {
        XSSFWorkbook workBook = null;
        List<T> list = new ArrayList<>();
        try (InputStream inputStream = file.getInputStream()){
            //读取工作簿
            workBook = new XSSFWorkbook(inputStream);
            list = doReadHeader(workBook,cls);
        }catch (final Exception e){
            log.warn("Read MultipartFile Excel error :{}",e.getMessage());
        }finally {
            if(workBook != null){
                //关闭工作簿
                try {
                    workBook.close();
                } catch (IOException e) {
                    log.warn("Read MultipartFile Excel error : {}",e.getMessage());
                }
            }
        }
        if(CollectionUtils.isNotEmpty(list)){
            return list.get(0);
        }
        return null;
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
            log.warn("Read File Excel error :{}",e.getMessage());
        }finally {
            if(workBook != null){
                //关闭工作簿
                try {
                    workBook.close();
                } catch (IOException e) {
                    log.warn("Read File Excel error :{}",e.getMessage(),e);
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
            if(null != row && !isRowEmpty(row)){
                final Map<String,Object> map = new HashMap<>(fields.length);
                for(int y= 0; y < fields.length;y++){
                    map.put(fields[y].getName(),getCellValue(row.getCell(y)));
                }
                lists.add(JSON.parseObject(JSON.toJSON(map).toString(),cls));
            }
        }
        return lists;
    }

    private static <T> List<T> doReadHeader(final XSSFWorkbook wb,final Class<T> cls){
        final XSSFSheet sheet = wb.getSheetAt(0);
        XSSFRow row;
        final List<T> lists = new ArrayList<>();
        Field[] fields = cls.getDeclaredFields();
        row = sheet.getRow(0);
        if(null != row && !isRowEmpty(row)){
            final Map<String,Object> map = new HashMap<>(fields.length);
            for(int y= 0; y < fields.length;y++){
                map.put(fields[y].getName(),getCellValue(row.getCell(y)));
            }
            lists.add(JSON.parseObject(JSON.toJSON(map).toString(),cls));
        }
        return lists;
    }

    private static boolean isRowEmpty(Row row){
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && !cell.getCellType().equals(CellType.BLANK)
                    && !cell.getCellType().equals(CellType._NONE)){
                return false;
            }
        }
        return true;
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

    public static BufferedOutputStream getBufferedOutputStream(final String fileName, final HttpServletResponse response) throws IOException {
        response.setContentType(APPLICATION_HEADER);
        response.setHeader(CONTENT_DISPOSITION, FILE_NAME_EXCHANGE_ORDER + new String(fileName.getBytes("gb2312"), StandardCharsets.ISO_8859_1));
        return new BufferedOutputStream(response.getOutputStream());
    }

    private static <T> void doExport(final String[] headers,
                                     final String sheetName,
                                     final List<T> data,
                                     final OutputStream outputStream,
                                     final String titleName,
                                     final List<CellRangeAddress> cellRangeAddressList) throws IOException, NoSuchFieldException, IllegalAccessException {
        final SXSSFWorkbook workbook = new SXSSFWorkbook();
        createSheet(workbook,headers,data,sheetName,titleName,cellRangeAddressList);
        if (outputStream != null) {
            workbook.write(outputStream);
        }

    }

    private static <T> void doManySheetExport(final String[] headers,
                                              final String sheetName,
                                              final List<T> data,
                                              final String titleName,
                                              final SXSSFWorkbook workbook) throws NoSuchFieldException, IllegalAccessException {
        createSheet(workbook,headers,data,sheetName,titleName,null);

    }

    private static <T> void createSheet(final SXSSFWorkbook wb,
                                        final String[] headers,
                                        final List<T> dataList,
                                        final String sheetName,
                                        final String titleName,
                                        final List<CellRangeAddress> cellRangeAddressList) throws NoSuchFieldException, IllegalAccessException {

        // 创建一张工作表
        final SXSSFSheet sheet = wb.createSheet(sheetName);
        final CellStyle style = wb.createCellStyle();
        final CellStyle style2 = wb.createCellStyle();
        final CellStyle style3 = wb.createCellStyle();
        final CellStyle style4 = wb.createCellStyle();

        //创建表头
        final Font font = wb.createFont();
        font.setFontName("宋体");
        font.setBold(true);
        //设置字体大小
        font.setFontHeightInPoints((short) 11);
        //选择需要用到的字体格式
        style.setFont(font);

        // 设置背景色
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 居中
        style.setAlignment(HorizontalAlignment.CENTER);
        //下边框
        style.setBorderBottom(BorderStyle.THIN);
        //右边框
        style.setBorderRight(BorderStyle.THIN);
        //上边框
        style.setBorderTop(BorderStyle.THIN);

        //创建文本内容
        final Font font2 = wb.createFont();
        font2.setFontName("宋体");
        //设置字体大小
        font2.setFontHeightInPoints((short) 10);
        //选择需要用到的字体格式
        style2.setFont(font2);

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

        //创建合并title
        final Font font3 = wb.createFont();
        font3.setFontName("宋体");
        font3.setBold(true);
        //设置字体大小
        font3.setFontHeightInPoints((short) 18);
        //选择需要用到的字体格式
        style3.setFont(font3);

        // 设置背景色
        style3.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        style3.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 居中
        style3.setAlignment(HorizontalAlignment.CENTER);
        //下边框
        style3.setBorderBottom(BorderStyle.THIN);
        //右边框
        style3.setBorderRight(BorderStyle.THIN);
        //左边框
        style3.setBorderLeft(BorderStyle.THIN);
        //上边框
        style3.setBorderTop(BorderStyle.THIN);

        //创建加粗文本内容
        final Font font4 = wb.createFont();
        font4.setFontName("宋体");
        font4.setBold(true);
        //设置字体大小
        font4.setFontHeightInPoints((short) 10);
        //选择需要用到的字体格式
        style4.setFont(font4);

        // 设置背景色
        style4.setFillForegroundColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        style4.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //垂直居中
        style4.setVerticalAlignment(VerticalAlignment.CENTER);
        // 水平向下居中
        style4.setAlignment(HorizontalAlignment.CENTER);
        //下边框
        style4.setBorderBottom(BorderStyle.THIN);
        //右边框
        style4.setBorderRight(BorderStyle.THIN);
        //左边框
        style4.setBorderLeft(BorderStyle.THIN);
        //上边框
        style4.setBorderTop(BorderStyle.THIN);

        if(StringUtils.isNotBlank(titleName)){
            //合并单元格
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, headers.length-1);
            sheet.addMergedRegion(region);
            if(CollectionUtils.isNotEmpty(cellRangeAddressList)){
                for(final CellRangeAddress merge : cellRangeAddressList){
                    sheet.addMergedRegion(merge);
                }
            }

            final Row titleRow = sheet.createRow(0);
            Cell cellTitle = titleRow.createCell(0);
            cellTitle.setCellStyle(style3);
            // 设置标题内容
            cellTitle.setCellValue(titleName);

            //表头
            final Row headerRow = sheet.createRow(1);

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellStyle(style);
                cell.setCellValue(headers[i]);
            }

            int rowNum = 1;
            for(final T data : dataList){
                final Row row = sheet.createRow(rowNum + 1);

                final Field[] fields = getExportFields(data.getClass());
                for (final Field field : fields) {
                    setData(field, data, field.getName(), style2, row,sheet,style4);
                }
                rowNum = sheet.getLastRowNum();
            }
        }else {
            //表头
            final Row headerRow = sheet.createRow(0);

            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellStyle(style);
                cell.setCellValue(headers[i]);
            }

            int rowNum = 0;
            for(final T data : dataList){
                final Row row = sheet.createRow(rowNum + 1);

                final Field[] fields = getExportFields(data.getClass());
                for (final Field field : fields) {
                    setData(field, data, field.getName(), style2, row,sheet,null);
                }
                rowNum = sheet.getLastRowNum();
            }
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
     * @param style     样式
     */
    private static <T> void setData(final Field dataField,
                                    final T object,
                                    final String property,
                                    final CellStyle style,
                                    final Row row,
                                    final SXSSFSheet sheet,
                                    final CellStyle specialStyle)
            throws IllegalAccessException, NoSuchFieldException {
        //允许访问private属性
        dataField.setAccessible(true);
        //获取属性值
        Object val = dataField.get(object);

        int cellnum = row.getLastCellNum() < 0 ? 0 : row.getLastCellNum();
        if (null != val) {
            dataSet(dataField,val,row,cellnum,style,property,sheet,specialStyle);
        }else{
            final Cell cell = row.createCell(cellnum);
            cell.setCellStyle(style);
        }
    }

    private static void dataSet(final Field dataField,
                                final Object val,
                                final Row row,
                                int cellnum,
                                final CellStyle style,
                                final String property,
                                final SXSSFSheet sheet,
                                final CellStyle specialStyle) throws NoSuchFieldException, IllegalAccessException {
        if (dataField.getType().toString().endsWith(STRING_CONSTANT)
                || dataField.getType().toString().endsWith(INTEGER_CONSTANT)
                || dataField.getType().toString().endsWith(INT_CONSTANT)
                || dataField.getType().toString().endsWith(LONG_CONSTANT)
                || dataField.getType().toString().endsWith(LONG_CONSTANT_SMALL)
                || dataField.getType().toString().endsWith(DOUBLE_CONSTANT)
                || dataField.getType().toString().endsWith(DOUBLE_CONSTANT_SMALL)) {
            final Cell cell = row.createCell(cellnum);
            int columnWidth = sheet.getColumnWidth(cellnum) / 256;
            final String value = String.valueOf(val);
            cell.setCellStyle(style);
            if(ObjectUtils.isNotEmpty(specialStyle)) {
                if (value.contains(BOLD)) {
                    cell.setCellStyle(specialStyle);
                }
            }
            int valueLength = value.getBytes().length;
            if(columnWidth < valueLength){
                sheet.setColumnWidth(cellnum, Math.min(valueLength * 256, 20000));
            }
            cell.setCellValue(String.valueOf(val).replace(BOLD,""));
        } else if (dataField.getType().toString().endsWith(DATE_CONSTANT)) {
            final DateFormat format = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
            final Cell cell = row.createCell(cellnum);
            cell.setCellStyle(style);
            cell.setCellValue(format.format((Date) val));
        } else if (dataField.getType().toString().endsWith(LIST_CONSTANT) || dataField.getType().toString().endsWith(COLLECTION_CONSTANT) || dataField.getType().toString().endsWith(SET_CONSTANT)) {
            listSet(cellnum,val,row,style,sheet,specialStyle);
        } else {
            final String str = ".";
            if (property.contains(str)) {
                final String p = property.substring(property.indexOf(str) + 1);
                final Field field = getDataField(val, p);
                setData(field, val, p, style,row,sheet,specialStyle);
            } else {
                final Cell cell = row.createCell(cellnum);
                cell.setCellStyle(style);
                cell.setCellValue(val.toString());
            }
        }
    }

    @SuppressWarnings("unchecked")
    private static <T> void listSet(final int cellnum,
                                    final Object val,
                                    final Row row,
                                    final CellStyle style,
                                    final SXSSFSheet sheet,
                                    final CellStyle specialStyle) throws NoSuchFieldException, IllegalAccessException {
        //适用于list平铺模板
        int listCell =cellnum;
        final Collection<T> list = (Collection<T>) val;
        for (Object o : list) {
            Field[] listFields = getExportFields(o.getClass());
            for (final Field listField : listFields) {
                setData(listField, o, listField.getName(), style, row,sheet,specialStyle);
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
