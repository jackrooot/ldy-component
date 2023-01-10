package cn.lindianyu.component.excel.down;

import cn.lindianyu.component.enmus.NumEnums;
import cn.lindianyu.component.excel.annotation.Excel;
import cn.lindianyu.component.excel.vo.CellStyleParams;
import cn.lindianyu.component.excel.vo.ExcelDownVo;
import cn.lindianyu.component.excel.vo.FieldNameAndOrder;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/9/29
 */
public final class ExcelDownUtil {
    private ExcelDownUtil() {
    }

    private static String format = "yyyy-MM-dd";

    private static List<String> stringToType = new ArrayList<>();

    private static List<String> dateToType = new ArrayList<>();

    static {
        stringToType.add("String");
        stringToType.add("Integer");
        stringToType.add("float");
        dateToType.add("Date");
    }

    private static void cellStyleSet(Workbook xssfWorkbook, Cell cell, CellStyleParams cellStyleParams) {
        XSSFColor xssfColor = new XSSFColor(
                new Color(cellStyleParams.getBackGroundColorRgbRed(), cellStyleParams.getBackGroundColorRgbGreen(),
                        cellStyleParams.getBackGroundColorRgbBlue()), new DefaultIndexedColorMap()
        );
        CellStyle newCellStyle = xssfWorkbook.createCellStyle();
        Font font = xssfWorkbook.createFont();
        font.setFontHeightInPoints(Short.parseShort("11"));
        //font.setBold(true);
        font.setColor(cellStyleParams.getFontColor());
        newCellStyle.setFont(font);
        newCellStyle.setWrapText(true);
        newCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        newCellStyle.setAlignment(HorizontalAlignment.CENTER);
        newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        newCellStyle.setFillForegroundColor(xssfColor.getIndexed());
        cell.setCellStyle(newCellStyle);
    }

    static void installField(Workbook workbook, Sheet sheet, Row row, List<FieldNameAndOrder> fieldNameList, CellStyleParams cellStyleParams) {
        for (int i = 0; i < fieldNameList.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(fieldNameList.get(i).getAnnotationsName());
            sheet.setColumnWidth(i, fieldNameList.get(i)
                    .getAnnotationsName()
                    .getBytes(StandardCharsets.UTF_8)
                    .length * NumEnums.INT255.num);
            if (cellStyleParams != null) {
                Boolean isFilter = cellStyleParams.getIsFilter();
                if (isFilter) {
                    CellRangeAddress cellRangeAddress = new CellRangeAddress(row.getRowNum(),
                            row.getRowNum(), NumEnums.INT0.num,
                            fieldNameList.size() - NumEnums.INT1.num);
                    sheet.setAutoFilter(cellRangeAddress);
                }
                cellStyleSet(workbook, cell, cellStyleParams);
            }
        }
    }

    public static <T> Workbook installOther(XSSFWorkbook xssfWorkbook, ExcelDownVo<T> excelDownVo, CellStyleParams cellStyleParams) {
        List<T> list = excelDownVo.getList();
        Class<T> clazz = excelDownVo.getClazz();
        Integer startRow = excelDownVo.getStartRow();

        Sheet sheet = xssfWorkbook.createSheet();

        List<FieldNameAndOrder> fieldNameAndOrders = new ArrayList<>();

        Arrays.stream(clazz.getDeclaredFields()).forEach(field -> {
            Excel annotation = field.getAnnotation(Excel.class);
            addAnnotation(annotation, field, fieldNameAndOrders);
        });
        fieldNameAndOrders.sort(Comparator.comparing(FieldNameAndOrder::getOrderNum));

        List<String> filedNameList = fieldNameAndOrders.stream().map(FieldNameAndOrder::getFiledName).collect(Collectors.toList());

        Row row = sheet.createRow(startRow);
        installField(xssfWorkbook, sheet, row, fieldNameAndOrders, cellStyleParams);
        if (list.size() == 0) {
            return xssfWorkbook;
        }
        try {
            for (T t1 : list) {
                Row row1 = sheet.createRow(++startRow);
                int column = 0;
                Field[] declaredFields = t1.getClass().getDeclaredFields();
                for (String filedName : filedNameList) {
                    for (Field field : declaredFields) {
                        if (filedName.equals(field.getName())) {
                            field.setAccessible(true);
                            Cell cell = row1.createCell(column);
                            if (column == 0) {
                                cell.setCellValue(column);
                            }
                            column++;
                            String type = field.getType().toString();
                            type = type.substring(type.lastIndexOf(".") + 1);
                            //字符串类型处理
                            if (stringToType.contains(type) && !Objects.isNull(field.get(t1))) {
                                cell.setCellValue(field.get(t1).toString());
                            }
                            //日期类型处理
                            if (dateToType.contains(type) && !Objects.isNull(field.get(t1))) {
                                cell.setCellValue(DateFormatUtils.format((Date) field.get(t1), format));
                            }
                        }
                    }
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return xssfWorkbook;
    }

    public static void addAnnotation(Excel annotation, Field field, List<FieldNameAndOrder> fieldNameAndOrders) {
        FieldNameAndOrder fieldNameAndOrder = new FieldNameAndOrder();
        if (!Objects.isNull(annotation)) {
            fieldNameAndOrder.setAnnotationsName(annotation.name());
            fieldNameAndOrder.setOrderNum(Integer.parseInt(StringUtils.isEmpty(annotation.orderNum()) ? "0" : annotation.orderNum()));
            fieldNameAndOrder.setFiledName(field.getName());
            fieldNameAndOrders.add(fieldNameAndOrder);
        }
    }
}
