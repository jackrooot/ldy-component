package cn.lindianyu.component.excel.vo;

import lombok.Data;

import java.util.List;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/8
 */
@Data
public class ExcelDownVo<T> {

    private List<T> list;

    private Class<T> clazz;

    private String templatePath;

    Integer startRow;
}
