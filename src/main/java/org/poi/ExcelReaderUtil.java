package org.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.base.ReaderConstant;
import org.base.ReaderResult;
import org.impl.ExcelReader;
import org.interfaces.IExcelReader;

import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 *
 * 导出工具类
 *
 * @author zhaobotao
 * @date 2019/6/14
 **/
public class ExcelReaderUtil {

    public static void readExcel(IExcelReader reader, String fileName) throws Exception {
        Map<String, Object> result;

        Map<String, Object> parameters = new HashMap<>();
        Map<Object, Object> dafTree = new HashMap<>();
        parameters.put("dafTree", dafTree);

        if (fileName.endsWith(ReaderConstant.EXCEL03_EXTENSION)) {
            //处理excel2003文件
            ExcelXlsReader excelXls = new ExcelXlsReader();
            excelXls.setExcelReader(reader);
            excelXls.setRecord(true);
            excelXls.setParameters(parameters);
            result = excelXls.process(fileName);
        } else if (fileName.endsWith(ReaderConstant.EXCEL07_EXTENSION)) {
            //处理excel2007文件
            ExcelXlsxReader excelXlsxReader = new ExcelXlsxReader();
            excelXlsxReader.setExcelReader(reader);
            excelXlsxReader.setRecord(true);
            excelXlsxReader.setParameters(parameters);
            result = excelXlsxReader.process(fileName);
        } else {
            throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
        }
        System.out.println("发送的总行数：" + result.get(ReaderConstant.TOTAL_ROWS));
        System.out.println("成功的总行数：" + result.get(ReaderConstant.SUCCESS_ROWS));
        List<ReaderResult> readerResult = (List<ReaderResult>) result.get(ReaderConstant.FAIL_INFO);
        System.out.println("失败的总行数：" + readerResult.size());

        // 导出失败记录

        String toPath = "C:\\Users\\botao\\Desktop\\测试错误导出.xlsx";

        // keep 100 rows in memory, exceeding rows will be flush ed to disk
        SXSSFWorkbook wb = new SXSSFWorkbook(100);
        Sheet sh = wb.createSheet();
        Row row1 = sh.createRow(0);
        Cell cell00 = row1.createCell(0);
        cell00.setCellValue("敏感词");
        Cell cell01 = row1.createCell(1);
        cell01.setCellValue("原因");
        for (int i = 0; i < readerResult.size(); i++) {
            Row row = sh.createRow(i + 1);
            Cell cell0 = row.createCell(0);
            cell0.setCellValue(readerResult.get(i).getCellList().get(0));
            Cell cell1 = row.createCell(1);
            cell1.setCellValue(readerResult.get(i).getNode());
        }

        FileOutputStream out = new FileOutputStream(toPath);
        wb.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }

    public static void main(String[] args) throws Exception {
        String path = "C:\\Users\\botao\\Desktop\\测试导入1.xlsx";
        IExcelReader rowReader = new ExcelReader();
        ExcelReaderUtil.readExcel(rowReader, path);
    }
}
