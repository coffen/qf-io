package com.qf.io.excel.writer;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.zip.ZipEntry;
import java.util.zip.ZipException;
import java.util.zip.ZipFile;
import java.util.zip.ZipOutputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qf.io.excel.DataFetcher;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Excel大数据导出工具
 * <br>
 * File Name: BigDataXlsxExport.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:15:41 
 * @version: v1.0
 *
 */
public class BigDataXlsxExport {
	
	private final static String TEMP_FILE_PATH = "E:/Temp";
    
	/**
	 * 导出大数据xlsx文档
	 * 
	 * @param fetcher 数据获取接口
	 * @param stream  输出流
	 * 
	 */
	public static void export(DataFetcher fetcher, OutputStream stream) throws ZipException, IOException {
		if (fetcher == null || fetcher.getTitles() == null || fetcher.getTitles().size() == 0 || stream == null) {
			return;
		}
		XSSFWorkbook book = new XSSFWorkbook();
		List<String> sheetRefList = new ArrayList<String>();
		List<File> sheetFileList = new ArrayList<File>();
		// 标题
		LinkedHashMap<String, String> titles = fetcher.getTitles();
		List<Map<String, Object>> data = null;
		File tmpXmlFile = null;
		int i = 0;
		while (fetcher.more()) {
			data = fetcher.getData();
			if (data == null || data.size() == 0) {
				continue;
			}
			XSSFSheet sheet = book.createSheet("sheet" + (i + 1));
			String sheetRef = sheet.getPackagePart().getPartName().getName();
			sheetRefList.add(sheetRef.substring(1));
			tmpXmlFile = new File(TEMP_FILE_PATH + "/" + ("sheet" + (i + 1)) + ".xml");
			generateXml(titles, data, tmpXmlFile, true);
			sheetFileList.add(tmpXmlFile);
			i++;
		}
		String tmpFileName = UUID.randomUUID().toString().replace("-", "");
		File tmpXlsx = new File(TEMP_FILE_PATH + "/file_" + tmpFileName + ".xlsx");;
		FileOutputStream os = new FileOutputStream(tmpXlsx);
		book.write(os);
		book.close();
		os.close();
		substitute(tmpXlsx, sheetFileList, sheetRefList, stream);
		stream.close();
		// 删除临时文件
		tmpXlsx.delete();
		for (File file : sheetFileList) {
			file.delete();
		}
	}
	
	private static void generateXml(LinkedHashMap<String, String> titles, List<Map<String, Object>> data, File file, boolean renderHeader) throws IOException {
		if (titles == null || titles.size() == 0) {
			return;
		}
		Writer write = new OutputStreamWriter(new FileOutputStream(file));
		SpreadsheetWriter sw = new SpreadsheetWriter(write);
		sw.beginSheet();
		int _row = 0;
		int _column = 0;
		String _key = null;
		Object _value = null;
		if (renderHeader) {
			Iterator<String> it = titles.keySet().iterator();
			sw.insertRow(0);
			while (it.hasNext()) {
				_key = it.next();
				if (StringUtils.isNotBlank(_key)) {
					_value = titles.get(_key);
					sw.createCell(_column, _value == null ? "" : (String)_value);
				}
				_column++;
			}
			sw.endRow();
		}
		_row = 1;
		if (data != null && data.size() > 0) {
			for (Map<String, Object> info : data) {
				Iterator<String> it = titles.keySet().iterator();
				sw.insertRow(_row);
				_column = 0;
				while (it.hasNext()) {
					_key = it.next();
					if (StringUtils.isNotBlank(_key)) {
						_value = info.get(_key);
						if (_value == null) {
							sw.createCell(_column, "");
						}
						else if (_value instanceof String) {
							sw.createCell(_column, (String)_value);
						}
						else if (_value instanceof Number) {
							sw.createCell(_column, Double.valueOf(_value.toString()));
						}
						else if (_value instanceof Date) {
							Calendar c = Calendar.getInstance();
							c.setTime((Date)_value);
							sw.createCell(_column, c, -1);
						}
					}
					_column++;
				}
				sw.endRow();
				_row++;
			}
		}		
		sw.endSheet();
		write.close();
	}
	
	private static void substitute(File zipfile, List<File> tmpfileList, List<String> entryList, OutputStream out) throws ZipException, IOException {
		ZipFile zip = new ZipFile(zipfile);
		ZipOutputStream zos = new ZipOutputStream(out);
		Enumeration<? extends ZipEntry> en = zip.entries();
		while (en.hasMoreElements()) {
			ZipEntry ze = (ZipEntry) en.nextElement();
			if (!entryList.contains(ze.getName())) {
				zos.putNextEntry(new ZipEntry(ze.getName()));
				InputStream is = zip.getInputStream(ze);
				copyStream(is, zos);
				is.close();
			}
		}
		InputStream is = null;
		for (int i = 0, len = entryList.size(); i < len; i++) {
			zos.putNextEntry(new ZipEntry(entryList.get(i)));
			is = new FileInputStream(tmpfileList.get(i));
			copyStream(is, zos);
			is.close();
		}
		zos.close();
		zip.close();
	}
	
	// 复制数据流
	public static void copyStream(InputStream in, OutputStream out) throws IOException {
		byte[] chunk = new byte[1024 * 10];
		int count;
		while ((count = in.read(chunk)) >= 0) {
			out.write(chunk, 0, count);
		}
	}
	
	/**
	 *  写入数据到XML
	 */
	public static class SpreadsheetWriter {
		
		private final Writer _out;
		private int _rownum;

		public SpreadsheetWriter(Writer out) {
			_out = out;
		}

		public void beginSheet() throws IOException {
			_out.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
			_out.write("<sheetData>" + System.lineSeparator());
		}

		public void endSheet() throws IOException {
			_out.write("</sheetData>" + System.lineSeparator());
			_out.write("</worksheet>");
		}

		// 插入新行
		public void insertRow(int rownum) throws IOException {
			_out.write("<row r=\"" + (rownum + 1) + "\">" + System.lineSeparator());
			this._rownum = rownum;
		}

		// 插入行结束符
		public void endRow() throws IOException {
			_out.write("</row>" + System.lineSeparator());
		}

		public void createCell(int columnIndex, String value, int styleIndex) throws IOException {
			String ref = new CellReference(_rownum, columnIndex).formatAsString();
			_out.write("<c r=\"" + ref + "\" t=\"inlineStr\"");
			if (styleIndex != -1) {
				_out.write(" s=\"" + styleIndex + "\"");
			}
			_out.write(">");
			_out.write("<is><t>" + value + "</t></is>");
			_out.write("</c>");
		}

		public void createCell(int columnIndex, String value) throws IOException {
			createCell(columnIndex, value, -1);
		}

		public void createCell(int columnIndex, double value, int styleIndex) throws IOException {
			String ref = new CellReference(_rownum, columnIndex).formatAsString();
			_out.write("<c r=\"" + ref + "\" t=\"n\"");
			if (styleIndex != -1) {
				_out.write(" s=\"" + styleIndex + "\"");
			}
			_out.write(">");
			_out.write("<v>" + value + "</v>");
			_out.write("</c>");
		}

		public void createCell(int columnIndex, double value) throws IOException {
			createCell(columnIndex, value, -1);
		}

		public void createCell(int columnIndex, Calendar value, int styleIndex) throws IOException {
			createCell(columnIndex, DateUtil.getExcelDate(value, false), styleIndex);
		}
		
		public void createCell(int columnIndex, Calendar value) throws IOException {
			createCell(columnIndex, DateUtil.getExcelDate(value, false), -1);
		}
		
	}
	
}
