package org.poi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;

/**
 * 导出测试类
 * http://poi.apache.org/components/spreadsheet/how-to.html#sxssf
 *
 * @author zhaobotao
 * @date 2019/6/14
 */
public class ExcelExport {
    public static void main(String[] args) throws Throwable {
        String toPath = "C:\\Users\\botao\\Desktop\\测试导出.xlsx";

        // keep 100 rows in memory, exceeding rows will be flush ed to disk
        SXSSFWorkbook wb = new SXSSFWorkbook(100);
        Sheet sh = wb.createSheet();
        Row row1 = sh.createRow(0);
        Cell cell0 = row1.createCell(0);
        cell0.setCellValue("敏感词");
        Cell cell1 = row1.createCell(1);
        cell1.setCellValue("原因");

        for (int rownum = 1; rownum < 300; rownum++) {
            Row row = sh.createRow(rownum);
            for (int cellnum = 0; cellnum < 2; cellnum++) {
                Cell cell = row.createCell(cellnum);
                String address = new CellReference(cell).formatAsString();
                cell.setCellValue(address);
            }
        }

        FileOutputStream out = new FileOutputStream(toPath);
        wb.write(out);
        out.close();

        // dispose of temporary files backing this workbook on disk
        wb.dispose();
    }
}
