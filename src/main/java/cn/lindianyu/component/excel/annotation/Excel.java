package cn.lindianyu.component.excel.annotation;

import java.lang.annotation.*;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/8
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(value = ElementType.FIELD)
@Documented
public @interface Excel {

    String orderNum() default "0";

    String name();
}
