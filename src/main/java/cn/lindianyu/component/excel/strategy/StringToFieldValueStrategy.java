package cn.lindianyu.component.excel.strategy;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/10
 */
public class StringToFieldValueStrategy implements FieldValueStrategy {

    String method = "valueOf";
    @Override
    public Object valueOf(Field field, String s) {
        try {
            String type = field.getType().toString();
            type = type.replaceFirst("class ","");
            Class<?> aClass = Class.forName(type);
            Object invoke = null;
            if (aClass.getTypeName().contains("String")) {
                return s;
            } else if(aClass.getTypeName().contains("Integer")||
                    aClass.getTypeName().contains("Long")||
                    aClass.getTypeName().contains("Short")||
                    aClass.getTypeName().contains("Double")||
                    aClass.getTypeName().contains("Float")
            ){
                Method method = aClass.getMethod(this.method,String.class);
                invoke = method.invoke(aClass, s);
            } else if (aClass.getTypeName().contains("Date")) {
                String format = "yyyy-MM-dd";
                invoke = new SimpleDateFormat(format).parse(s);
            } else{
                invoke = s;
            }

            return invoke;
        } catch (ClassNotFoundException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (ParseException e) {
            e.printStackTrace();
        }
        return null;
    }
}
