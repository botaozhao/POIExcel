package org.base;

import java.io.Serializable;

/**
 * 导入状态码
 *
 * @author zhaobotao
 * @date 2019/6/17
 */
public enum ReaderEnum implements Serializable {

    /**
     * 成功
     */
    SUCCESS("200", "成功"),

    /**
     * 其他异常
     */
    OTHER("2020", "其他异常"),

    /**
     * 数据重复
     */
    REPEAT("2021", "数据重复");

    private String code;
    private String msg;

    ReaderEnum() {
    }

    ReaderEnum(String code, String msg) {
        this.code = code;
        this.msg = msg;
    }

    public String getCode() {
        return code;
    }

    public String getMsg() {
        return msg;
    }
}
