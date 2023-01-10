package cn.lindianyu.component.excel.down;

import cn.lindianyu.component.enmus.NumEnums;
import cn.lindianyu.component.excel.annotation.Excel;
import cn.lindianyu.component.excel.down.base.BaseDownBuild;
import cn.lindianyu.component.excel.vo.CellStyleParams;
import cn.lindianyu.component.excel.vo.FieldNameAndOrder;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2023/1/5
 */
public class ExcelDownBuild<T> extends BaseDownBuild<T> {

    protected String format = "yyyy-MM-dd";

    private Field[] declaredFields;

    private HashMap<String, Integer> map = new HashMap<>();
    //字段名字列表
    private List<FieldNameAndOrder> fieldNameList = new ArrayList<>();

    private List<String> nameList;
    //数据量
    private Integer pageNum;

    @Override
    public BaseDownBuild<T> setDateFormat(String dateFormat) {
        this.format = dateFormat;
        return this;
    }

    @Override
    public BaseDownBuild<T> buildSheetEndNum(Integer sheetEndNum) {
        super.sheetEndNum = sheetEndNum;
        return this;
    }

    @Override
    public Workbook buildExcel(List<T> t, Class<T> clazz, CellStyleParams cellStyleParams) {
        if (t.size() == NumEnums.INT0.num) {
            return super.workbook;
        }
        if (super.sheetEndNum != 0) {
            pageNum = Double.valueOf(Math.ceil((double) t.size() / super.sheetEndNum)).intValue();
        } else {
            pageNum = 1;
        }
        //初始化
        initFields(clazz);
        for (int i = 0; i < pageNum; i++) {
            int startRow = 0;
            int curNum = i;
            Sheet sheet = super.workbook.createSheet(String.valueOf(curNum + 1));
            Row row = sheet.createRow(startRow);
            List<T> list;
            if (sheetEndNum != 0) {
                list = t.subList(curNum * super.sheetEndNum, (curNum + 1) * super.sheetEndNum);
            } else {
                list = t;
            }
            ExcelDownUtil.installField(super.workbook, sheet, row, fieldNameList, cellStyleParams);
            excelBeanSet(sheet, startRow, list);
        }
        return super.workbook;
    }

    private void initFields(Class<T> clazz) {
        declaredFields = clazz.getDeclaredFields();
        for (int i = 0; i < declaredFields.length; i++) {
            Field field = declaredFields[i];
            Excel annotatiaon = field.getAnnotation(Excel.class);
            addAnnotation(annotatiaon, field);
            map.put(field.getName(), i);
        }
        fieldNameList.sort(Comparator.comparing(FieldNameAndOrder::getOrderNum));
        nameList = fieldNameList.stream().map(FieldNameAndOrder::getFiledName).collect(Collectors.toList());
    }

    private void addAnnotation(Excel annotation, Field field) {
        FieldNameAndOrder fieldNameAndOrder = new FieldNameAndOrder();
        if (!Objects.isNull(annotation)) {
            fieldNameAndOrder.setAnnotationsName(annotation.name());
            fieldNameAndOrder.setOrderNum(Integer.parseInt(StringUtils.isEmpty(annotation.orderNum()) ? "0" : annotation.orderNum()));
            fieldNameAndOrder.setFiledName(field.getName());
            this.fieldNameList.add(fieldNameAndOrder);
        }
    }

    private <T> void excelBeanSet(Sheet sheet, Integer startRow, List<T> dataList) {
        try {
            int size = dataList.size();
            for (int i = 0; i < size; i++) {
                T t1 = dataList.get(i);
                Row row1 = sheet.createRow(++startRow);
                int column = 0;
                for (String filedName : nameList) {
                    Field field = declaredFields[map.get(filedName)];
                    field.setAccessible(true);
                    Cell cell = row1.createCell(column);
                    if (column == 0) {
                        cell.setCellValue(startRow);
                    }
                    column++;
                    String type = field.getType().toString();
                    type = type.substring(type.lastIndexOf(".") + 1);
                    //字符串类型处理
                    if (stringToType.contains(type) && !Objects.isNull(field.get(t1))) {
                        XSSFRichTextString xssfRichTextString = new XSSFRichTextString();
                        xssfRichTextString.setString(field.get(t1).toString());
                        cell.setCellValue(xssfRichTextString);
                    }
                    //日期类型处理
                    if (dateToType.contains(type) && !Objects.isNull(field.get(t1))) {
                        cell.setCellValue(DateFormatUtils.format((Date) field.get(t1), format));
                    }
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }
}
