package cn.lindianyu.component.excel.down.base;

import cn.lindianyu.component.excel.vo.CellStyleParams;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2023/1/9
 */
public abstract class BaseDownBuild<T> {

    protected static List<String> stringToType = new ArrayList<>();

    protected static List<String> dateToType = new ArrayList<>();

    protected Workbook workbook = new SXSSFWorkbook();

    protected Integer sheetEndNum;

    public BaseDownBuild() {
        sheetEndNum = 0;
    }

    static {
        stringToType.add("String");
        stringToType.add("Integer");
        stringToType.add("float");
        dateToType.add("Date");
    }

    public abstract BaseDownBuild<T> setDateFormat(String dateFormat);

    public abstract BaseDownBuild<T> buildSheetEndNum(Integer sheetEndNum);

    public abstract Workbook buildExcel(List<T> t, Class<T> clazz, CellStyleParams cellStyleParams);

}
