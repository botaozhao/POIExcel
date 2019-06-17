package org.interfaces;

import org.base.ReaderResult;

import java.util.List;
import java.util.Map;

/**
 * excel导入-核心数据处理
 *
 * @author zhaobotao
 * @date 2019/6/14
 */
public interface IExcelReader {

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
    ReaderResult sendRows(int sheetIndex, String sheetName, int curRow, List<String> cellList, Map<String, Object> parameters);
}
