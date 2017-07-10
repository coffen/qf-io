package com.qf.io.test.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.qf.io.FileErrorException;
import com.qf.io.excel.writer.module.PoiListDataModule;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: PoiExcel列表数据导出测试
 * <br>
 * File Name: PoiListDataExporterTest.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年7月6日 下午3:01:37 
 * @version: v1.0
 *
 */
public class PoiListDataExporterTest {
	
	private final static Logger log = LoggerFactory.getLogger(PoiListDataExporterTest.class);
	
	private final String path = Thread.currentThread().getContextClassLoader().getResource("").getPath();
	
	@Test
	public void export() throws IOException, FileErrorException {
		String filename = "listDataModule.xlsx";
		PoiListDataModule module = new PoiListDataModule(path + filename);
		
		String exportFile = "listDataExportFile.xlsx";
		OutputStream os = new FileOutputStream(new File(path + exportFile));
		
		LinkedHashMap<String, String> titleMap = new LinkedHashMap<String, String>();
		String[] titles = new String[] { "商品ID", "标题", "商品链接", "价格", "库存", "备注" };
		String[] fields = new String[] { "id", "title", "link", "price", "stock", "remark" };
		for (int i = 0; i < fields.length; i++) {
			titleMap.put(fields[i], titles[i]);
		}
		List<Map<String, ?>> data = new ArrayList<Map<String, ?>>();
		
		log.error("导出开始...");
		
		module.export(titleMap, data, os);
		
		log.error("导出完成.");
	}
	
}
