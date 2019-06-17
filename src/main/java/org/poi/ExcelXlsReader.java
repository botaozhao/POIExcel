package org.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.base.ReaderConstant;
import org.base.ReaderResult;
import org.interfaces.IExcelReader;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * xls2003版本大数据量问题
 *
 * @author zhaobotao
 * @create 2019-6-14
 * @desc 用于解决.xls2003版本大数据量问题
 **/
public class ExcelXlsReader implements HSSFListener {

    private IExcelReader excelReader;

    public void setExcelReader(IExcelReader excelReader) {
        this.excelReader = excelReader;
    }

    Map<String, Object> parameters;

    public void setParameters(Map<String, Object> parameters) {
        this.parameters = parameters;
    }

    private int minColumns = -1;

    private POIFSFileSystem fs;

    /**
     * 总行数
     */
    private int totalRows = 0;

    /**
     * 成功行数
     */
    private int successRows = 0;

    /**
     * 上一行row的序号
     */
    private int lastRowNumber;

    /**
     * 上一单元格的序号
     */
    private int lastColumnNumber;

    /**
     * 是否输出formula，还是它对应的值
     */
    private boolean outputFormulaValues = true;

    /**
     * 用于转换formulas
     */
    private SheetRecordCollectingListener workbookBuildingListener;

    //excel2003工作簿
    private HSSFWorkbook stubWorkbook;

    private SSTRecord sstRecord;

    private FormatTrackingHSSFListener formatListener;

    private final HSSFDataFormatter formatter = new HSSFDataFormatter();

    /**
     * 文件的绝对路径
     */
    private String filePath = "";

    //表索引
    private int sheetIndex = 0;

    private BoundSheetRecord[] orderedBSRs;

    @SuppressWarnings("unchecked")
    private ArrayList boundSheetRecords = new ArrayList();

    private int nextRow;

    private int nextColumn;

    private boolean outputNextStringRecord;

    /**
     * 当前行
     */
    private int curRow = 0;

    /**
     * 存储一行记录所有单元格的容器
     */
    private List<String> cellList = new ArrayList<String>();

    /**
     * 失败是否需要记录
     */
    private boolean isRecord = false;

    public void setRecord(boolean record) {
        isRecord = record;
    }

    /**
     * 存储失败的记录
     */
    private List<ReaderResult> failInfo = new ArrayList<>();

    /**
     * 判断整行是否为空行的标记
     */
    private boolean flag = false;

    @SuppressWarnings("unused")
    private String sheetName;

    /**
     * 遍历excel下所有的sheet
     *
     * @param fileName
     * @throws Exception
     */
    public Map<String, Object> process(String fileName) throws Exception {
        Map<String, Object> result = new HashMap<String, Object>();
        filePath = fileName;
        this.fs = new POIFSFileSystem(new FileInputStream(fileName));
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }
        factory.processWorkbookEvents(request, fs);

        result.put(ReaderConstant.TOTAL_ROWS, totalRows);
        result.put(ReaderConstant.SUCCESS_ROWS, successRows);
        result.put(ReaderConstant.FAIL_INFO, failInfo);
        return result;
    }

    /**
     * HSSFListener 监听方法，处理Record
     * 处理每个单元格
     *
     * @param record <pre>
     *                   关键字	用途
     *               all	to suppress all warnings
     *               boxing 	to suppress warnings relative to boxing/unboxing operations
     *               cast	to suppress warnings relative to cast operations
     *               dep-ann	to suppress warnings relative to deprecated annotation
     *               deprecation	to suppress warnings relative to deprecation
     *               fallthrough	 to suppress warnings relative to missing breaks in switch statements
     *               finally 	to suppress warnings relative to finally block that don’t return
     *               hiding	to suppress warnings relative to locals that hide variable
     *               incomplete-switch	 to suppress warnings relative to missing entries in a switch statement (enum case)
     *               nls	 to suppress warnings relative to non-nls string literals
     *               null	to suppress warnings relative to null analysis
     *               rawtypes	to suppress warnings relative to un-specific types when using generics on class params
     *               restriction	to suppress warnings relative to usage of discouraged or forbidden references
     *               serial	to suppress warnings relative to missing serialversionuid field for a serializable class
     *               static-access	o suppress warnings relative to incorrect static access
     *               synthetic-access 	 to suppress warnings relative to unoptimized access from inner classes
     *               unchecked	 to suppress warnings relative to unchecked operations
     *               unqualified-field-access	to suppress warnings relative to field access unqualified
     *               unused	to suppress warnings relative to unused code
     *               </pre>
     */
    @SuppressWarnings({"unchecked"})
    @Override
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = null;
        String value = null;
        switch (record.getSid()) {
            case BoundSheetRecord.sid:
                boundSheetRecords.add(record);
                break;
            //开始处理每个sheet
            case BOFRecord.sid:
                BOFRecord br = (BOFRecord) record;
                if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                    //如果有需要，则建立子工作簿
                    if (workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                    }

                    if (orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                    sheetName = orderedBSRs[sheetIndex].getSheetname();
                    sheetIndex++;
                }
                break;
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            //单元格为空白
            case BlankRecord.sid:
                BlankRecord brec = (BlankRecord) record;
                thisRow = brec.getRow();
                thisColumn = brec.getColumn();
                thisStr = "";
                cellList.add(thisColumn, thisStr);
                break;
            //单元格为布尔类型
            case BoolErrRecord.sid:
                BoolErrRecord berec = (BoolErrRecord) record;
                thisRow = berec.getRow();
                thisColumn = berec.getColumn();
                thisStr = berec.getBooleanValue() + "";
                cellList.add(thisColumn, thisStr);
                //如果里面某个单元格含有值，则标识该行不为空行
                this.checkRowIsNull(value);
                break;
            //单元格为公式类型
            case FormulaRecord.sid:
                FormulaRecord frec = (FormulaRecord) record;
                thisRow = frec.getRow();
                thisColumn = frec.getColumn();
                if (outputFormulaValues) {
                    if (Double.isNaN(frec.getValue())) {
                        outputNextStringRecord = true;
                        nextRow = frec.getRow();
                        nextColumn = frec.getColumn();
                    } else {
                        thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
                    }
                } else {
                    thisStr = '"' + HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression()) + '"';
                }
                cellList.add(thisColumn, thisStr);
                //如果里面某个单元格含有值，则标识该行不为空行
                this.checkRowIsNull(value);
                break;
            case StringRecord.sid:
                //单元格中公式的字符串
                if (outputNextStringRecord) {
                    StringRecord srec = (StringRecord) record;
                    thisStr = srec.getString();
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    outputNextStringRecord = false;
                }
                break;
            case LabelRecord.sid:
                LabelRecord lrec = (LabelRecord) record;
                curRow = thisRow = lrec.getRow();
                thisColumn = lrec.getColumn();
                value = lrec.getValue().trim();
                value = value.equals("") ? "" : value;
                cellList.add(thisColumn, value);
                //如果里面某个单元格含有值，则标识该行不为空行
                this.checkRowIsNull(value);
                break;
            case LabelSSTRecord.sid:
                //单元格为字符串类型
                LabelSSTRecord lsrec = (LabelSSTRecord) record;
                curRow = thisRow = lsrec.getRow();
                thisColumn = lsrec.getColumn();
                if (sstRecord == null) {
                    cellList.add(thisColumn, "");
                } else {
                    value = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
                    value = value.equals("") ? "" : value;
                    cellList.add(thisColumn, value);
                    //如果里面某个单元格含有值，则标识该行不为空行
                    this.checkRowIsNull(value);
                }
                break;
            case NumberRecord.sid:
                //单元格为数字类型
                NumberRecord numrec = (NumberRecord) record;
                curRow = thisRow = numrec.getRow();
                thisColumn = numrec.getColumn();

                //第一种方式
                //value = formatListener.formatNumberDateCell(numrec).trim();//这个被写死，采用的m/d/yy h:mm格式，不符合要求

                //第二种方式，参照formatNumberDateCell里面的实现方法编写
                Double valueDouble = ((NumberRecord) numrec).getValue();
                String formatString = formatListener.getFormatString(numrec);
                if (formatString.contains("m/d/yy") || formatString.contains("yyyy/mm/dd") || formatString.contains("yyyy/m/d")) {
                    formatString = "yyyy-MM-dd hh:mm:ss";
                }
                int formatIndex = formatListener.getFormatIndex(numrec);
                value = formatter.formatRawCellContents(valueDouble, formatIndex, formatString).trim();

                value = value.equals("") ? "" : value;
                //向容器加入列值
                cellList.add(thisColumn, value);
                //如果里面某个单元格含有值，则标识该行不为空行
                this.checkRowIsNull(value);
                break;
            default:
                break;
        }

        //遇到新行的操作
        if (thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }

        //空值的操作
        if (record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord) record;
            curRow = thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            cellList.add(thisColumn, "");
        }

        //更新行和列的值
        if (thisRow > -1) {
            lastRowNumber = thisRow;
        }
        if (thisColumn > -1) {
            lastColumnNumber = thisColumn;
        }

        //行结束时的操作
        if (record instanceof LastCellOfRowDummyRecord) {
            if (minColumns > 0) {
                //列值重新置空
                if (lastColumnNumber == -1) {
                    lastColumnNumber = 0;
                }
            }
            lastColumnNumber = -1;

            //该行不为空 发送数据
            if (flag) {
                ReaderResult result = this.excelReader.sendRows(sheetIndex, sheetName, curRow + 1, cellList, parameters);
                if (result.isSuccess()) {
                    successRows++;
                } else {
                    if (isRecord) {
                        result.setCellList(cellList);
                        this.failInfo.add(result);
                        cellList = new ArrayList<String>();
                    }
                }
                totalRows++;
            }
            //清空容器
            if (!cellList.isEmpty()) {
                cellList.clear();
            }
            flag = false;
        }
    }

    /**
     * 如果里面某个单元格含有值，则标识该行不为空行
     *
     * @param value value
     */
    public void checkRowIsNull(String value) {
        if (StringUtils.isNotBlank(value)) {
            flag = true;
        }
    }
}