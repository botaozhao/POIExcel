package org.poi;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.base.ReaderConstant;
import org.base.ReaderResult;
import org.interfaces.IExcelReader;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * xlsx2007版本大数据量问题
 *
 * @author zhaobotao
 * @create 2018-01-18 14:28
 * @desc POI读取excel有两种模式，一种是用户模式，一种是事件驱动模式
 * 采用SAX事件驱动模式解决XLSX文件，可以有效解决用户模式内存溢出的问题，
 * 该模式是POI官方推荐的读取大数据的模式，
 * 在用户模式下，数据量较大，Sheet较多，或者是有很多无用的空行的情况下，容易出现内存溢出
 * <p>
 * 用于解决.xlsx2007版本大数据量问题
 **/
public class ExcelXlsxReader extends DefaultHandler {

	private IExcelReader excelReader;

	public void setExcelReader(IExcelReader excelReader) {
		this.excelReader = excelReader;
	}

	Map<String, Object> parameters;

	public void setParameters(Map<String, Object> parameters) {
		this.parameters = parameters;
	}

	/**
	 * 单元格中的数据可能的数据类型
	 */
	enum CellDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
	}

	/**
	 * 共享字符串表
	 */
	private SharedStringsTable sst;

	/**
	 * 上一次的索引值
	 */
	private String lastIndex;

	/**
	 * 文件的绝对路径
	 */
	private String filePath = "";

	/**
	 * 工作表索引
	 */
	private int sheetIndex = 0;

	/**
	 * sheet名
	 */
	private String sheetName = "";

	/**
	 * 总行数
	 */
	private int totalRows=0;

	/**
	 * 成功行数
	 */
	private int successRows = 0;

	/**
	 * 一行内cell集合
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

	/**
	 * 当前行
	 */
	private int curRow = 1;

	/**
	 * 当前列
	 */
	private int curCol = 0;

	/**
	 * T元素标识
	 */
	private boolean isTElement;

	/**
	 * 判断上一单元格是否为文本空单元格
	 */
	private boolean endElementFlag = false;
	private boolean charactersFlag = false;

	/**
	 * 异常信息，如果为空则表示没有异常
	 */
	private String exceptionMessage;

	/**
	 * 单元格数据类型，默认为字符串类型
	 */
	private CellDataType nextDataType = CellDataType.SSTINDEX;

	private final DataFormatter formatter = new DataFormatter();

	/**
	 * 单元格日期格式的索引
	 */
	private short formatIndex;

	/**
	 * 日期格式字符串
	 */
	private String formatString;

	//定义前一个元素和当前元素的位置，用来计算其中空的单元格数量，如A6和A8等
	private String preRef = null, ref = null;

	//定义该文档一行最大的单元格数，用来补全一行最后可能缺失的单元格
	private String maxRef = null;

	/**
	 * 单元格
	 */
	private StylesTable stylesTable;

	/**
	 * 遍历工作簿中所有的电子表格
	 * 并缓存在mySheetList中
	 *
	 * @param filename
	 * @throws Exception
	 */
	public Map<String, Object> process(String filename) throws Exception {
		Map<String, Object> result = new HashMap<String, Object>();
		filePath = filename;
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader(pkg);
		stylesTable = xssfReader.getStylesTable();
		SharedStringsTable sst = xssfReader.getSharedStringsTable();
		XMLReader parser = this.fetchSheetParser(sst);
		XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
		//遍历sheet
		while (sheets.hasNext()) {
			//标记初始行为第一行
			curRow = 0;
			sheetIndex++;
			//sheets.next()和sheets.getSheetName()不能换位置，否则sheetName报错
			InputStream sheet = sheets.next();
			sheetName = sheets.getSheetName();
			InputSource sheetSource = new InputSource(sheet);
			//解析excel的每条记录，在这个过程中startElement()、characters()、endElement()这三个函数会依次执行
			parser.parse(sheetSource);
			sheet.close();
		}
		result.put(ReaderConstant.TOTAL_ROWS, totalRows);
		result.put(ReaderConstant.SUCCESS_ROWS, successRows);
		result.put(ReaderConstant.FAIL_INFO, failInfo);
		return result;
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader("org.apache.xerces.parsers.SAXParser");
		this.sst = sst;
		parser.setContentHandler(this);
		return parser;
	}

	/**
	 * 第一个执行
	 *
	 * @param uri
	 * @param localName
	 * @param name
	 * @param attributes
	 * @throws SAXException
	 */
	@Override
	public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
		//c => 单元格
		if ("c".equals(name)) {


			//前一个单元格的位置
			if (preRef == null) {
				preRef = attributes.getValue("r");
			} else {
				//判断前一次是否为文本空字符串，true则表明不是文本空字符串，false表明是文本空字符串跳过把空字符串的位置赋予preRef
				if (endElementFlag){
					preRef = ref;
				}
			}

			//当前单元格的位置
			ref = attributes.getValue("r");
			//设定单元格类型
			this.setNextDataType(attributes);
			endElementFlag = false;
			charactersFlag = false;
		}

		//当元素为t时
		if ("t".equals(name)) {
			isTElement = true;
		} else {
			isTElement = false;
		}

		//置空
		lastIndex = "";
	}



	/**
	 * 第二个执行
	 * 得到单元格对应的索引值或是内容值
	 * 如果单元格类型是字符串、INLINESTR、数字、日期，lastIndex则是索引值
	 * 如果单元格类型是布尔值、错误、公式，lastIndex则是内容值
	 * @param ch
	 * @param start
	 * @param length
	 * @throws SAXException
	 */
	@Override
	public void characters(char[] ch, int start, int length) throws SAXException {
		charactersFlag = true;
		lastIndex += new String(ch, start, length);
	}

	/**
	 * 第三个执行
	 *
	 * @param uri
	 * @param localName
	 * @param name
	 * @throws SAXException
	 */
	@Override
	public void endElement(String uri, String localName, String name) throws SAXException {
		//t元素也包含字符串 这个程序没经过
		if (isTElement) {
			//将单元格内容加入rowlist中，在这之前先去掉字符串前后的空白符
			String value = lastIndex.trim();
			cellList.add(curCol, value);
			endElementFlag = true;
			curCol++;
			isTElement = false;
			//如果里面某个单元格含有值，则标识该行不为空行
			this.checkRowIsNull(value);
		} else if ("v".equals(name)) {
			//v => 单元格的值，如果单元格是字符串，则v标签的值为该字符串在SST中的索引
			// 根据索引值获取对应的单元格值
			String value = this.getDataValue(lastIndex.trim(), "");
			//补全单元格之间的空单元格
			if (!ref.equals(preRef)) {
				int len = countNullCell(ref, preRef);
				for (int i = 0; i < len; i++) {
					cellList.add(curCol, "");
					curCol++;
				}
			}
			cellList.add(curCol, value);
			curCol++;
			endElementFlag = true;
			//如果里面某个单元格含有值，则标识该行不为空行
			this.checkRowIsNull(value);
		} else {
			//如果标签名称为row，这说明已到行尾，调用optRows()方法
			if ("row".equals(name)) {
				//默认第一行为表头，以该行单元格数目为最大数目
				if (curRow == 1) {
					maxRef = ref;
				}
				//补全一行尾部可能缺失的单元格
				if (maxRef != null) {
					int len = -1;
					//前一单元格，true则不是文本空字符串，false则是文本空字符串
					if (charactersFlag){
						len = countNullCell(maxRef, ref);
					}else {
						len = countNullCell(maxRef, preRef);
					}
					for (int i = 0; i <= len; i++) {
						cellList.add(curCol, "");
						curCol++;
					}
				}

				//该行不为空行且该行不是第一行，则发送 //每行结束时，调用sendRows()方法
				if (flag){
					ReaderResult result = this.excelReader.sendRows(sheetIndex, sheetName, curRow + 1, cellList, parameters);
					if (result.isSuccess()) {
						successRows++;
					}else{
						if (isRecord) {
							result.setCellList(cellList);
							this.failInfo.add(result);
							cellList = new ArrayList<String>();
						}
					}
					totalRows++;
				}
				if (!cellList.isEmpty()) {
					cellList.clear();
				}
				curRow++;
				curCol = 0;
				preRef = null;
				ref = null;
				flag=false;
			}
		}
	}

	/**
	 * 处理数据类型
	 *
	 * @param attributes
	 */
	public void setNextDataType(Attributes attributes) {
		//cellType为空，则表示该单元格类型为数字
		nextDataType = CellDataType.NUMBER;
		formatIndex = -1;
		formatString = null;
		//单元格类型
		String cellType = attributes.getValue("t");
		String cellStyleStr = attributes.getValue("s");
		//获取单元格的位置，如A1,B1
		String columnData = attributes.getValue("r");

		//处理布尔值
		if ("b".equals(cellType)) {
			nextDataType = CellDataType.BOOL;
		} else if ("e".equals(cellType)) {
			//处理错误
			nextDataType = CellDataType.ERROR;
		} else if ("inlineStr".equals(cellType)) {
			nextDataType = CellDataType.INLINESTR;
		} else if ("s".equals(cellType)) {
			//处理字符串
			nextDataType = CellDataType.SSTINDEX;
		} else if ("str".equals(cellType)) {
			nextDataType = CellDataType.FORMULA;
		}

		if (cellStyleStr != null) {
			//处理日期
			int styleIndex = Integer.parseInt(cellStyleStr);
			XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
			formatIndex = style.getDataFormat();
			formatString = style.getDataFormatString();
			if (formatString.contains("m/d/yyyy") || formatString.contains("yyyy/mm/dd")|| formatString.contains("yyyy/m/d") ) {
				nextDataType = CellDataType.DATE;
				formatString = "yyyy-MM-dd hh:mm:ss";
			}

			if (formatString == null) {
				nextDataType = CellDataType.NULL;
				formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
			}
		}
	}

	/**
	 * 对解析出来的数据进行类型处理
	 * @param value   单元格的值，
	 *                value代表解析：BOOL的为0或1， ERROR的为内容值，FORMULA的为内容值，INLINESTR的为索引值需转换为内容值，
	 *                SSTINDEX的为索引值需转换为内容值， NUMBER为内容值，DATE为内容值
	 * @param thisStr 一个空字符串
	 * @return
	 */
	@SuppressWarnings("deprecation")
	public String getDataValue(String value, String thisStr) {
		switch (nextDataType) {
			// 这几个的顺序不能随便交换，交换了很可能会导致数据错误
			//布尔值
			case BOOL:
				char first = value.charAt(0);
				thisStr = first == '0' ? "FALSE" : "TRUE";
				break;
			//错误
			case ERROR:
				thisStr = "\"ERROR:" + value.toString() + '"';
				break;
			//公式
			case FORMULA:
				thisStr = '"' + value.toString() + '"';
				break;
			case INLINESTR:
				XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
				thisStr = rtsi.toString();
				rtsi = null;
				break;
			//字符串
			case SSTINDEX:
				String sstIndex = value.toString();
				try {
					int idx = Integer.parseInt(sstIndex);
					//根据idx索引值获取内容值
					XSSFRichTextString rtss = new XSSFRichTextString(sst.getEntryAt(idx));
					thisStr = rtss.toString();
					//有些字符串是文本格式的，但内容却是日期

					rtss = null;
				} catch (NumberFormatException ex) {
					thisStr = value.toString();
				}
				break;
			//数字
			case NUMBER:
				if (formatString != null) {
					thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString).trim();
				} else {
					thisStr = value;
				}
				thisStr = thisStr.replace("_", "").trim();
				break;
			//日期
			case DATE:
				thisStr = formatter.formatRawCellContents(Double.parseDouble(value), formatIndex, formatString);
				// 对日期字符串作特殊处理，去掉T
				thisStr = thisStr.replace("T", " ");
				break;
			default:
				thisStr = " ";
				break;
		}
		return thisStr;
	}

	public int countNullCell(String ref, String preRef) {
		//excel2007最大行数是1048576，最大列数是16384，最后一列列名是XFD
		String xfd = ref.replaceAll("\\d+", "");
		String xfd_1 = preRef.replaceAll("\\d+", "");

		xfd = fillChar(xfd, 3, '@', true);
		xfd_1 = fillChar(xfd_1, 3, '@', true);

		char[] letter = xfd.toCharArray();
		char[] letter_1 = xfd_1.toCharArray();
		int res = (letter[0] - letter_1[0]) * 26 * 26 + (letter[1] - letter_1[1]) * 26 + (letter[2] - letter_1[2]);
		return res - 1;
	}

	public String fillChar(String str, int len, char let, boolean isPre) {
		int len_1 = str.length();
		if (len_1 < len) {
			if (isPre) {
				for (int i = 0; i < (len - len_1); i++) {
					str = let + str;
				}
			} else {
				for (int i = 0; i < (len - len_1); i++) {
					str = str + let;
				}
			}
		}
		return str;
	}

	/**
	 * 如果里面某个单元格含有值，则标识该行不为空行
	 * @param value value
	 */
	public void checkRowIsNull(String value){
		if (StringUtils.isNotBlank(value)) {
			flag = true;
		}
	}
}
