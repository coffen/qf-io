package com.qf.io.excel.writer;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Excel导出模板接口
 * <br>
 * File Name: ExportModule.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:17:09 
 * @version: v1.0
 *
 */
public interface ExportModule {
	
	public final static int FORMAT_XLS = 1;
	public final static int FORMAT_XLSX = 1 >> 1;
	
	/**
	 * 解析模板文件
	 */
	public void parse();
	
	/**
	 * 将模板转换为HTML表格Css样式（部分模板色调上可能有偏差）
	 * 
	 * @return String css字符串
	 */
	public String transformToCss();
	
}
