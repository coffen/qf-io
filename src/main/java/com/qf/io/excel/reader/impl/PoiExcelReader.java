package com.qf.io.excel.reader.impl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.qf.io.FileErrorException;
import com.qf.io.excel.PoiUtils;
import com.qf.io.excel.reader.ExcelReader;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Excel数据读取工具类接口POI实现（非线程安全）
 * <br>
 * File Name: PoiExcelReader.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月27日 上午9:38:57 
 * @version: v1.0
 *
 */
public class PoiExcelReader implements ExcelReader {
	
	private final static Logger log = LoggerFactory.getLogger(PoiExcelReader.class);
	
	private Workbook workbook;
	
	private AtomicInteger mark = new AtomicInteger(0);	// 当前读取的位置（zero-based）
	private int maxLength = 0; 							// 最大行（zero-based）
	
	public PoiExcelReader(File file) throws FileErrorException, IOException {
		if (!(file.exists() && file.isFile())) {
			throw new FileErrorException(FileErrorException.FILE_NOT_FOUND);
		}
		String suffix = FilenameUtils.getExtension(file.getName());
		if (!("xls".equalsIgnoreCase(suffix) || "xlsx".equalsIgnoreCase(suffix))) {
			throw new FileErrorException("Excel", FileErrorException.FILE_EXTENSION_ERROR);
		}		
		FileInputStream fis = new FileInputStream(file);
		workbook = "xls".equalsIgnoreCase(suffix) ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
		fis.close();
		if (workbook.getSheetAt(0) != null) {
			maxLength = workbook.getSheetAt(0).getLastRowNum();
		}		
	}
	
	@Override
	public boolean hasMore() {
		return mark.get() < maxLength;
	}

	@Override
	public List<Object> read() {
		List<Object> data = null;
		List<List<Object>> dataList = read(1);
		if (dataList != null && dataList.size() > 0) {
			data = dataList.get(0);
		}
		return data;
	}
	
	@Override
	public List<List<Object>> read(int length) {
		if (length <= 0 || !hasMore()) {
			return null;  // 目前不支持重复读取
		}
		List<List<Object>> records = new ArrayList<List<Object>>();
		Sheet _sheet = workbook.getSheetAt(0);
		List<Object> tmp = null;
		Row _row = null;
		for (int i = 0; i < length; i++) {
			if (mark.get() > maxLength) {
				break;
			}
			_row = _sheet.getRow(mark.getAndIncrement());
			if (_row != null) {
				tmp = new ArrayList<Object>();
				for (int j = 0; j < _row.getLastCellNum(); j++) {
					// 空单元格也算作一个属性
					tmp.add(PoiUtils.getValue(_row.getCell(j)));
				}
			}
			records.add(tmp);
		}
		return records;
	}

	@Override
	public Map<String, Object> read(List<String> fields) {
		Map<String, Object> map = null;
		List<Map<String, Object>> records = read(fields, 1);
		if (records != null && records.size() > 0) {
			map = records.get(0);
		}
		return map;
	}
	
	@Override
	public List<Map<String, Object>> read(List<String> fields, int length) {
		if (fields == null || fields.size() == 0 || length <= 0 || !hasMore()) {
			log.error("ExcelReader.read@param[fields]不能为空!");
			return null;
		}
		// 检查fields的非空属性
		for (String field : fields) {
			if (StringUtils.isBlank(field)) {
				log.error("ExcelReader@param[fields]不能有空元素!");
				return null;
			}
		}
		List<Map<String, Object>> records = new ArrayList<Map<String, Object>>();
		Sheet _sheet = workbook.getSheetAt(0);
		Map<String, Object> tmp = null;
		Row _row = null;
		for (int i = 0; i < length; i++) {
			if (mark.get() > maxLength) {
				break;
			}
			_row = _sheet.getRow(mark.getAndIncrement());
			if (_row != null) {
				tmp = new HashMap<String, Object>();
				for (int j = 0; j < _row.getLastCellNum(); j++) {
					if (j >= fields.size()) {
						break;
					}
					tmp.put(fields.get(j), PoiUtils.getValue(_row.getCell(j)));
				}
			}
			records.add(tmp);
		}
		return records;
	}

}
