package cn.lindianyu.component.excel.sax.xlsx;


import cn.lindianyu.component.excel.annotation.Excel;
import cn.lindianyu.component.excel.strategy.StringToFieldValueStrategy;
import cn.lindianyu.component.excel.vo.FieldNameAndOrder;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/9
 */
public class ExcelXlsxReadOfSax<T> implements XSSFSheetXMLHandler.SheetContentsHandler {

    private T t;

    private HashMap<String,Field> map = new HashMap<>();

    private List<String> filedNameList;

    List<T> dataList = new ArrayList<>();

    private int cellNum = 0;

    private int rowNum =0;

    int startRow = 1;

    boolean blankFlag = true;

    private ExcelXlsxReadOfSax(){

    }
    ExcelXlsxReadOfSax(T t) {
        this.t = t;
        Field[] declaredFields = t.getClass().getDeclaredFields();
        List<FieldNameAndOrder> fieldNameAndOrderList = new ArrayList<>();
        for (Field field : declaredFields) {
            Excel annotation = field.getAnnotation(Excel.class);
            if (!Objects.isNull(annotation)) {
                cellNum++;
                FieldNameAndOrder fieldNameAndOrder = new FieldNameAndOrder();
                map.put(field.getName(), field);
                fieldNameAndOrder.setAnnotationsName(annotation.name());
                fieldNameAndOrder.setOrderNum(Integer.parseInt(StringUtils.isEmpty(annotation.orderNum()) ? "0" : annotation.orderNum()));
                fieldNameAndOrder.setFiledName(field.getName());
                fieldNameAndOrderList.add(fieldNameAndOrder);
            }
        }
        cellNum = 0;
        fieldNameAndOrderList.sort(Comparator.comparing(FieldNameAndOrder::getOrderNum));
        filedNameList = fieldNameAndOrderList.stream().map(FieldNameAndOrder::getFiledName).collect(Collectors.toList());
    }
    ExcelXlsxReadOfSax(T t, Integer startRow) {
        this.startRow = startRow;
        this.t = t;
        Field[] declaredFields = t.getClass().getDeclaredFields();
        List<FieldNameAndOrder> fieldNameAndOrderList = new ArrayList<>();
        for (Field field : declaredFields) {
            Excel annotation = field.getAnnotation(Excel.class);
            if (!Objects.isNull(annotation)) {
                cellNum++;
                FieldNameAndOrder fieldNameAndOrder = new FieldNameAndOrder();
                map.put(field.getName(), field);
                fieldNameAndOrder.setAnnotationsName(annotation.name());
                fieldNameAndOrder.setOrderNum(Integer.parseInt(StringUtils.isEmpty(annotation.orderNum()) ? "0" : annotation.orderNum()));
                fieldNameAndOrder.setFiledName(field.getName());
                fieldNameAndOrderList.add(fieldNameAndOrder);
            }
        }
        cellNum = 0;
        fieldNameAndOrderList.sort(Comparator.comparing(FieldNameAndOrder::getOrderNum));
        filedNameList = fieldNameAndOrderList.stream().map(FieldNameAndOrder::getFiledName).collect(Collectors.toList());
    }

    @Override
    public void startRow(int i) {
        this.rowNum = i;
    }

    @Override
    public void endRow(int i) {
        try {
            if(i>=startRow && !blankFlag){
                dataList.add(t);
            }
            this.t= (T)t.getClass().newInstance();
        } catch (InstantiationException | IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void cell(String s, String s1 , XSSFComment xssfComment) {
        try {
            cellNum++;
            int size = filedNameList.size();
            if (StringUtils.isNotEmpty(s1)) {
                blankFlag = false;
            }
            if(rowNum  < 1) {
                return;
            }
            int mod=(cellNum - 1) % size;
            String fieldName = filedNameList.get(mod);
            Field field = map.get(fieldName);
            field.setAccessible(true);
            StringToFieldValueStrategy stringToFieldValueStrategy = new StringToFieldValueStrategy();
            field.set(t,stringToFieldValueStrategy.valueOf(field,s1));
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
    }

    @Override
    public void headerFooter(String s, boolean b, String s1) {

    }
    List<T> getDataList() {
        return dataList;
    }
}
