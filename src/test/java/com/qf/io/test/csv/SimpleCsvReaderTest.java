package com.qf.io.test.csv;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections4.CollectionUtils;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSON;
import com.qf.io.csv.SimpleCsvReader;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: SimpleCsvReader测试
 * <br>
 * File Name: SimpleCsvReaderTest.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017-06-24 10:55:39 
 * @version: v1.0
 *
 */
public class SimpleCsvReaderTest {
	
	private final static Logger log = LoggerFactory.getLogger(SimpleCsvReader.class);
	
	private final String path = Thread.currentThread().getContextClassLoader().getResource("").getPath();
	
	@Test
	public void testReader() throws FileNotFoundException, IOException {
		String filename = "test.csv";
		FileInputStream fis = new FileInputStream(path + filename);
		SimpleCsvReader csvReader = new SimpleCsvReader(fis, ',', Charset.forName("GB2312"));
		List<String[]> list = csvReader.read(1);
		while (CollectionUtils.isNotEmpty(list)) {
			list = csvReader.read(1);			
			log.error(JSON.toJSONString(list));
		}
	}
	
	@Test
	public void testReaderByField() throws FileNotFoundException, IOException {
		String filename = "test.csv";
		String[] fields = new String[] {"id", "shopName", "shopUrl", "point", "efficiency", "credit", "mainPageSales", "productNum", "sales", "sellerNick", "createTime"};
		List<String> fieldList = Arrays.asList(fields);
		FileInputStream fis = new FileInputStream(path + filename);		
		SimpleCsvReader csvReader = new SimpleCsvReader(fis, ',', Charset.forName("GB2312"));
		List<Map<String,String>> list = csvReader.read(fieldList, 1);
		while (CollectionUtils.isNotEmpty(list)) {
			list = csvReader.read(fieldList, 1);			
			log.error(JSON.toJSONString(list));
		}
	}
	
}
