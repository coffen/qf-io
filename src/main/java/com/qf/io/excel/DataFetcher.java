package com.qf.io.excel;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 数据获取接口
 * <br>
 * 封装了数据来源的处理
 * <br>
 * File Name: DataFetcher.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:20:37 
 * @version: v1.0
 *
 */
public interface DataFetcher {
	
	/**
	 * 判断是否还有数据待处理
	 * 
	 * @return boolean true: 是, false: 否
	 */
	public boolean more();
	
	/**
	 * 获取标题（标题-属性的映射关系）
	 * 
	 * @return LinkedHashMap<String, String> 列标题组成的集合（格式：LinkedHashMap<Property, Name>）
	 */
	public LinkedHashMap<String, String> getTitles();
	
	/**
	 * 获取数据
	 * <P>
	 * 本方法可以重复调用, 直到无法获取数据为止, 这种调用方式一般用于数据库访问读取数据
	 * 也可以和<b>more</b>组合调用, 根据more方法的返回值判断是否还可以获取更多数据
	 * </p>
	 * 
	 * @return List<Map<String, Object>> 行数据组成的列表
	 */
	public List<Map<String, Object>> getData();
	
}
