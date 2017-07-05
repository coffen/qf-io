package com.qf.io.test.excel;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSON;
import com.qf.io.csv.SimpleCsvReader;
import com.qf.io.excel.reader.impl.PoiExcelReader;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: PoiExcelReader测试类
 * <br>
 * File Name: PoiExcelReaderTest.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年7月5日 上午11:15:34 
 * @version: v1.0
 *
 */
public class PoiExcelReaderTest {
	
	private final static Logger log = LoggerFactory.getLogger(SimpleCsvReader.class);
	
	// @Test
	public void readExcel() throws Exception {
		String filePath = "C:\\Users\\Richard\\Desktop\\导出商品20170629173201.xlsx";
		PoiExcelReader reader = new PoiExcelReader(new File(filePath));	
		log.error("**直接读取----------------------------");
		log.error("**读取标题---------");
		log.error(JSON.toJSONString(reader.read()));
		log.error("**读取内容---------");
		int count = 0;
		while (reader.hasMore()) {
			log.error(JSON.toJSONString(reader.read(), 10));
			count++;
		}
		log.error("**一共读取{}条记录---------", count);
		reader.close();
	}
	
	@Test
	public void readExcelByField() throws Exception {
		String filePath = "C:\\Users\\Richard\\Desktop\\导出商品20170629173201.xlsx";
		PoiExcelReader reader = new PoiExcelReader(new File(filePath));	
		log.error("**按属性读取----------------------------");
		log.error("**读取标题---------");
		log.error(JSON.toJSONString(reader.read()));
		log.error("**读取内容---------");
		String[] properties = new String[] { "id", "title", "url", "shopName", "shopUrl", "salePrice", "commissionRate", "couponPrice", "couponUrl", "couponAmount", "couponDeadline", "applyTime", "auditTime", "arrangeTime", "status" };
		List<String> fields = Arrays.asList(properties);		
		int count = 0;
		while (reader.hasMore()) {
			log.error(JSON.toJSONString(reader.read(fields), 10));
			count++;
		}
		log.error("**一共读取{}条记录---------", count);
		reader.close();
	}
	
}
