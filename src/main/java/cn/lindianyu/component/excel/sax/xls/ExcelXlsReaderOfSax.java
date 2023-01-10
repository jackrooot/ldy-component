package cn.lindianyu.component.excel.sax.xls;

import cn.lindianyu.component.excel.annotation.Excel;
import cn.lindianyu.component.excel.down.ExcelDownUtil;
import cn.lindianyu.component.excel.strategy.StringToFieldValueStrategy;
import cn.lindianyu.component.excel.vo.FieldNameAndOrder;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.record.*;

import java.lang.reflect.Field;
import java.util.*;
import java.util.function.Supplier;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/13
 */
public class ExcelXlsReaderOfSax<T> implements HSSFListener {

    ExcelXlsReaderOfSax(T t) {
        this.t = t;
    }

    private SSTRecord sstRecord;

    List<T> dataList = new ArrayList<>();
    //指定sheet读取
    Integer sheetNum = 0;
    //表索引
    private int sheetIndex = 0;
    //读取表索引
    int sheetStartIndex = 0;
    //行索引
    private int rowIndex = 0;
    //读取行索引
    Integer rowStartIndex = 2;
    //列索引
    private int cellIndex = 0;

    private T t;

    private int rowCellNum = 0;

    Boolean blankFlag = true;

    @Override
    public void processRecord(Record record) {
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                break;
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    Map<String, Supplier<?>> genericProperties = br.getGenericProperties();
                    //sheet页新增
                    sheetIndex++;
                    //行索引重置
                    rowIndex = 0;
                }
                break;
            case StringRecord.sid:
                StringRecord stringRecord = (StringRecord) record;
                break;
            case LabelSSTRecord.sid:
                if (rowIndex >= rowStartIndex) {
                    LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                    String value = sstRecord.getString(labelSSTRecord.getSSTIndex()).toString();
                    if (StringUtils.isNotEmpty(value)) {
                        blankFlag = false;
                    }
                    ArrayList<FieldNameAndOrder> fieldNameList = new ArrayList<>();
                    Field[] declaredFields = t.getClass().getDeclaredFields();
                    for (int i = 0; i < declaredFields.length; i++) {
                        Field field = declaredFields[i];
                        Excel annotatiaon = field.getAnnotation(Excel.class);
                        ExcelDownUtil.addAnnotation(annotatiaon, field, fieldNameList);
                    }
                    fieldNameList.sort(Comparator.comparing(FieldNameAndOrder::getOrderNum));
                    try {
                        StringToFieldValueStrategy stringToFieldValueStrategy = new StringToFieldValueStrategy();
                        FieldNameAndOrder fieldNameAndOrder = fieldNameList.get(cellIndex);
                        Field declaredField = t.getClass().getDeclaredField(fieldNameAndOrder.getFiledName());
                        Object o = stringToFieldValueStrategy.valueOf(declaredField, value);
                        declaredField.setAccessible(true);
                        declaredField.set(t, o);
                    } catch (NoSuchFieldException e) {
                        e.printStackTrace();
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    cellIndex++;
                }
                break;
        }
        //行结束
        if (record instanceof LastCellOfRowDummyRecord) {
            rowIndex++;
            //列索引重置
            cellIndex = 0;

            int row = ((LastCellOfRowDummyRecord) record).getRow();

            boolean sheetFlag = sheetNum != 0 && sheetIndex != sheetNum;
            boolean rowFlag = row < rowStartIndex;
            if (sheetFlag || rowFlag || blankFlag) {
                return;
            }
            dataList.add(t);
            try {
                this.t = (T) t.getClass().newInstance();
                if (rowCellNum == 0) {
                    rowCellNum = t.getClass().getDeclaredFields().length;
                }
            } catch (InstantiationException | IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }
}
