package org.base;

import java.util.List;

/**
 * 导入结果
 *
 * @author zhaobotao
 * @date 2019/6/14
 */
public class ReaderResult {
    /**
     * 是否成功
     */
    private boolean isSuccess;
    /**
     * 行内cell集合
     */
    private List<String> cellList;
    /**
     * code
     */
    private String code;
    /**
     * 备注
     */
    private String node;

    public ReaderResult() {
    }

    public ReaderResult(boolean isSuccess) {
        this.isSuccess = isSuccess;
        this.node = ReaderEnum.SUCCESS.getMsg();
        this.code = ReaderEnum.SUCCESS.getCode();
    }

    public ReaderResult(ReaderEnum readerEnum) {
        this.isSuccess = false;
        this.node = readerEnum.getMsg();
        this.code = readerEnum.getCode();
    }

    public boolean isSuccess() {
        return isSuccess;
    }

    public void setSuccess(boolean success) {
        isSuccess = success;
    }

    public List<String> getCellList() {
        return cellList;
    }

    public void setCellList(List<String> cellList) {
        this.cellList = cellList;
    }

    public String getNode() {
        return node;
    }

    public void setNode(String node) {
        this.node = node;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }
}
