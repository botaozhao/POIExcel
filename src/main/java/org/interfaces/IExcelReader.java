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
    ReaderResult sendRows(int sheetIndex, String sheetName, int curRow, List<String> cellList, Map<String, Object> parameters);
}
