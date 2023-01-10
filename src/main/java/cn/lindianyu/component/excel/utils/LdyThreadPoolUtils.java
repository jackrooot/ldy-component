package cn.lindianyu.component.excel.utils;

import cn.lindianyu.component.enmus.NumEnums;

import java.util.concurrent.ArrayBlockingQueue;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2023/1/9
 */
public class LdyThreadPoolUtils {
    public static ThreadPoolExecutor getThreadPoolExecutor(){
        return new ThreadPoolExecutor(
                NumEnums.INT3.num,
                NumEnums.INT4.num,
                NumEnums.INT5000.num.longValue(),
                TimeUnit.MILLISECONDS,
                new ArrayBlockingQueue<>(NumEnums.INT20.num));
    }
}
