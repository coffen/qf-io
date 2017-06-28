package com.qf.io.csv;

import java.io.IOException;
import java.io.InputStream;
import java.io.Reader;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.csvreader.CsvReader;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Csv数据读取工具类
 * <br>
 * File Name: SimpleCsvReader.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017-06-24 09:21:25 
 * @version: v1.0
 *
 */
public class SimpleCsvReader {
	
	private final static Logger log = LoggerFactory.getLogger(SimpleCsvReader.class);
	
	private CsvReader reader;
	
	public SimpleCsvReader(Reader r) {
		reader = new CsvReader(r);
		reader.setUseTextQualifier(false);
	}
	
	public SimpleCsvReader(InputStream inputStream, char delimiter, Charset charSet) {
		reader = new CsvReader(inputStream, delimiter, charSet);
		reader.setUseTextQualifier(false);
	}
	
	/**
	 * 获取文本分割符, 默认为“,”
	 */
	public char getDelimiter() {
		return reader.getDelimiter();
	}
	
	/**
	 * 设置文本分割符
	 */
	public void setDelimiter(char delimiter) {
		reader.setDelimiter(delimiter);
	}
	
	/**
	 * 获取文本引用符, 默认为“"”
	 */
	public char getTextQualifier() {
		return reader.getTextQualifier();
	}

	/**
	 * 设置文本引用符
	 */
	public void setTextQualifier(char textQualifier) {
		reader.setUseTextQualifier(true);
		reader.setTextQualifier(textQualifier);
	}
	
	/**
	 * 是否还有数据待读取
	 * 
	 * @return true: 是, false：否.
	 * 
	 */
	public boolean hasMore() throws IOException {
		return reader.readRecord();
	}
	
	/**
	 * <p>读取单行并解析返回数据列表</p>
	 *
	 * <p>返回的数据按列顺序排列</p>
	 * 
	 * @return List<String[]> 数据列表
	 */
	public String[] read() throws IOException {
		String[] data = null;
		List<String[]> dataList = read(1);
		if (dataList != null && dataList.size() > 0) {
			data = dataList.get(0);
		}
		return data;
	}
	
	/**
	 * <p>按指定行数读取并解析返回数据列表</p>
	 * 
	 * <p>返回的数据按列的顺序排列</p>
	 * 
	 * @param length 读取的行数, 如超出剩余的行数, 则返回剩余的行数据, 如已经是最后一行, 返回空列表
	 * 
	 * @return List<String[]> 数据列表
	 */
	public List<String[]> read(int length) throws IOException {
		List<String[]> records = new ArrayList<String[]>();
		if (length <= 0) {
			log.error("参数不正确: length={}", length);
			return records;
		}
		String[] _data = null; 
		for (int i = 0; i < length; i++) {
			if (!reader.readRecord()) {
				break;
			}
			_data = reader.getValues();
			if (_data != null && _data.length > 0) {
				records.add(_data);		
			}
		}
		return records;
	}
	
	/**
	 * <p>读取单行并解析返回数据列表</p>
	 * 
	 * @param fields 列名按顺序排列的列表, 返回的Map数据以field属性作为Key值
	 * 
	 * @return List<Map<String, String>> 数据列表
	 */
	public Map<String, String> read(List<String> fields) throws IOException {
		Map<String, String> map = null;
		List<Map<String, String>> records = read(fields, 1);
		if (records != null && records.size() > 0) {
			map = records.get(0);
		}
		return map;
	}
	
	/**
	 * <p>按指定行数读取并解析返回数据列表</p>
	 * 
	 * @param fields 列名按顺序排列的列表, 返回的Map数据以field属性作为Key值
	 * @param length 读取的行数, 如超出剩余的行数, 则返回剩余的行数据, 如已经是最后一行, 返回空列表
	 * 
	 * @return List<Map<String, String>> 数据流
	 */
	public List<Map<String, String>> read(List<String> fields, int length) throws IOException {
		List<Map<String, String>> records = new ArrayList<Map<String, String>>();
		if (fields == null || fields.size() == 0 || length <= 0) {
			log.error("参数不正确: fields={}, length={}", fields, length);
			return records;
		}
		// 检查fields的非空属性
		for (String field : fields) {
			if (StringUtils.isBlank(field)) {
				log.error("SimpleCsvReader.read@param[fields]不能有空元素!");
				return records;
			}
		}
		Map<String, String> tmp = null;
		String[] _data = null; 
		for (int i = 0; i < length; i++) {
			if (!reader.readRecord()) {
				break;
			}
			_data = reader.getValues();
			if (_data != null && _data.length > 0) {
				tmp = new HashMap<String, String>();
				for (int j = 0; j < _data.length; j++) {
					if (j >= fields.size()) {
						break;
					}
					tmp.put(fields.get(j), _data[j]);
				}
				records.add(tmp);		
			}
		}
		return records;
	}
	
	/**
	 * 关闭流并清空资源
	 */
	public void close() {
		reader.close();
	}
	
}
