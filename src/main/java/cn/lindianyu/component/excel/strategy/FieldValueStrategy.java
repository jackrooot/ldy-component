package cn.lindianyu.component.excel.strategy;

import java.lang.reflect.Field;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/10
 */
public interface FieldValueStrategy {
    Object valueOf(Field field,String s);
}
