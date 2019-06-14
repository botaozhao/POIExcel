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
     * 备注
     */
    private String node;

    public ReaderResult() {
    }

    public ReaderResult(boolean isSuccess) {
        this.isSuccess = isSuccess;
    }

    public ReaderResult(boolean isSuccess, String node) {
        this.isSuccess = isSuccess;
        this.node = node;
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

}
