package cn.lindianyu.component.enmus;

/**
 * @Author: ya
 * @Email: 1002019117@qq.com
 * @Date:create in  2022/10/8
 */
public enum NumEnums {
    /**
    *数字枚举
    */
    INT0(0),
    INT1(1),
    INT2(2),
    INT3(3),
    INT4(4),
    INT20(20),
    INT5000(5000),
    INT255(255);
    public Integer num;

    NumEnums(Integer num) {
        this.num = num;
    }
}
