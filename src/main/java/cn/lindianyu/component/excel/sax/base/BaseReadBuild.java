package cn.lindianyu.component.excel.sax.base;

import java.io.InputStream;
import java.util.List;

/**
 * @Author: ya
 * @param <T>
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/11/17
 */
public abstract class BaseReadBuild<T> {

    protected abstract BaseReadBuild<T> buildFile(InputStream inputStream, Class<? extends T> clazz);

    protected abstract BaseReadBuild<T> buildStartRow(Integer rowNum);

    protected abstract BaseReadBuild<T> buildSheetNum(Integer rowNum);

    protected abstract BaseReadBuild<T> buildJumpBlankRow(Boolean b);

    protected abstract List<T> build();


}
