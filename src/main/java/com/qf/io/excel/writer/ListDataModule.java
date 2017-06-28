package com.qf.io.excel.writer;

import java.io.IOException;
import java.io.OutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description:  List数据导出模板接口（用于导出List类型数据）
 * <br>
 * File Name: ListDataModule.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:17:17 
 * @version: v1.0
 *
 */
public interface ListDataModule extends ExportModule {
	
	/**
	 * <p>按模板设定导出List数据到指定输出流</p>
	 * 
	 * @param titles 表头标题, 标题全为空时隐藏表头; 格式：Map<属性名: 表头标题>
	 * @param data   导出数据; 格式：List<Map<属性名: 数据对象>>
	 * @param stream 输出流
	 */
	public void export(LinkedHashMap<String, String> titles, List<Map<String, ?>> data, OutputStream stream) throws IOException;
	
}
