package com.qf.io.excel.writer;

import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 表达式语言导出模板接口
 * <br>
 * 专用于导出EL表达式的模板， 单元格中数据、图片、公式均可通过表达式语言设置
 * <br>
 * File Name: ELModule.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:15:57 
 * @version: v1.0
 *
 */
public interface ELModule extends ExportModule {
	
	/**
	 * 按模板设定导出Bean对象到指定输出流
	 * 
	 * @param stream	输出流
	 * @param beans 	封装各项数据的Bean对象数组, 对应每一个Sheet的数据
	 */
	public void export(OutputStream stream, Serializable... beans) throws IOException;
	
}