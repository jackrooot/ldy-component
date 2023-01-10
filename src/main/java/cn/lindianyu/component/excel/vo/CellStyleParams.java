package cn.lindianyu.component.excel.vo;

import lombok.Data;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/8
 */
@Data
public class CellStyleParams {
    //红
    private Integer backGroundColorRgbRed;
    //绿
    private Integer backGroundColorRgbGreen;
    //蓝
    private Integer backGroundColorRgbBlue;
    //是否筛选
    private Boolean isFilter;
    //字体颜色
    private Short fontColor;

}
