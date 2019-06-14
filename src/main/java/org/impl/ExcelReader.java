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
public class ExcelReader implements IExcelReader {

    @Override
    public ReaderResult sendRows(int sheetIndex, String sheetName, int curRow, List<String> cellList, Map<String, Object> other) {
        boolean result = true;
        System.out.print(curRow + "\t");
        for (int i = 0; i < cellList.size(); i++) {
            System.out.print(cellList.get(i) + "\t");
        }
        System.out.println("\t" + result);

        if (cellList.get(0).charAt(0) == '0') {
            return new ReaderResult(false, "数据重复");
        }
        if (cellList.get(0).charAt(0) == '1') {
            return new ReaderResult(false, "其他异常");
        }

        return new ReaderResult(true);
    }
}
