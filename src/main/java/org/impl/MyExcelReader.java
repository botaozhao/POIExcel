package org.impl;

import org.base.ReaderResult;
import org.interfaces.IExcelReader;

import java.util.List;
import java.util.Map;

/**
 * excel导入-核心数据处理
 *
 * @author zhaobotao
 * @date 2019/6/14
 */
public class MyExcelReader implements IExcelReader {


    /**
     * excel导出数据发送接口
     *
     * @param sheetIndex sheetIndex
     * @param sheetName  sheetName
     * @param curRow     当前行
     * @param cellList   当前行锁包含的列
     * @param parameters 附加参数
     * @return ReaderResult
     */
    @Override
    public ReaderResult sendRows(int sheetIndex, String sheetName, int curRow, List<String> cellList, Map<String, Object> parameters) {
        boolean result = true;
        System.out.print(curRow + "\t");
        for (int i = 0; i < cellList.size(); i++) {
            System.out.print(cellList.get(i) + "\t");
        }

        Map<Object, Object> dafTree = (Map<Object, Object>) parameters.get("dafTree");

        // 数据处理，模拟数据重复验重以及其他异常导致的导入失败
        if (cellList.get(0).charAt(0) == '0') {
            dafTree.put(cellList.get(0), "123");
            result = false;
            System.out.println("\t" + result);
            return new ReaderResult(result, "数据重复");
        }
        if (cellList.get(0).charAt(0) == '1') {
            result = false;
            System.out.println("\t" + result);
            return new ReaderResult(result, "其他异常");
        }

        System.out.println("\t" + result);
        return new ReaderResult(result);
    }
}
