package com.qf.io.excel.reader;

import java.util.List;
import java.util.Map;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Excel数据读取工具类接口
 * <br>
 * File Name: ExcelReader.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月27日 上午9:27:33 
 * @version: v1.0
 *
 */
public interface ExcelReader {
	
	/**
	 * 是否还有数据待读取
	 * 
	 * @return true: 是, false：否.
	 * 
	 */
	public boolean hasMore();
	
	/**
	 * 读取单行并解析返回数据列表
	 * <p>
	 * 返回的数据按列顺序排列
	 * </p>
	 * 
	 * @return List<Object> 数据列表
	 */
	public List<Object> read();
	
	/**
	 * 读取指定行数并解析返回数据列表
	 * <p>
	 * 返回的数据按列顺序排列
	 * </p>
	 * 
	 * @param int length 读取的行数, 如超出剩余的行数, 则返回剩余的行数据, 如已经是最后一行, 返回Null
	 * 
	 * @return List<List<Object>> 数据列表
	 */
	public List<List<Object>> read(int length);
	
	/**
	 * 读取单行并解析返回数据列表
	 * 
	 * @param List<String> fields 列名按顺序排列的列表, 返回的Map数据以field属性作为Key值
	 * 
	 * @return Map<String, Object> 数据列表
	 */
	public Map<String, Object> read(List<String> fields);
	
	/**
	 * 读取指定行数并解析返回数据列表
	 * 
	 * @param List<String> fields 列名按顺序排列的列表, 返回的Map数据以field属性作为Key值
	 * @param int length 读取的行数, 如超出剩余的行数, 则返回剩余的行数据, 如已经是最后一行, 返回Null
	 * 
	 * @return List<Map<String, Object>> 数据列表
	 */
	public List<Map<String, Object>> read(List<String> fields, int length);
	
	/**
	 * 关闭Reader
	 */
	public void close();
	
}
