package com.qf.io.excel.writer;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.qf.io.FileErrorException;
import com.qf.io.excel.DataFetcher;
import com.qf.io.excel.ExcelFileFormat;
import com.qf.io.excel.UnsupportedExcelDataException;
import com.qf.io.excel.writer.module.PoiELModule;
import com.qf.io.excel.writer.module.PoiListDataModule;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Excel数据导出工具
 * <br>
 * File Name: ExcelWriter.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:11:51 
 * @version: v1.0
 *
 */
public class ExcelWriter {
	
	private final static Logger log = LoggerFactory.getLogger(ExcelWriter.class);
	
	public static final int EXCEL_XLS_ROW_COUNT_MAX = 65535; // xls最大输出行数
	public static final int EXCEL_XLS_COLUMN_COUNT_MAX = 256; // xls最大输出列数
	
	public static final int EXCEL_XLSX_ROW_COUNT_MAX = 65535; // xlsx最大输出行数
	public static final int EXCEL_XLSX_COLUMN_COUNT_MAX = 16384; // xlsx最大输出列数
	
	/**
	 * 按指定的模板导出List数据集, 包括单行表头, 只能输出单个Sheet, 不包含合并单元格, 可以有图片、计算公式
	 * <P>
	 * 适合于小数据集的导出（Rowcount <= 65535）
	 * </p>
	 * 
	 * @param titles  		表格标题栏
	 * @param data    		数据集
	 * @param modulePath  	导出模板
	 * @param stream  		输出流
	 */
	public static void write(LinkedHashMap<String, String> titles, List<Map<String, ?>> data, String modulePath, OutputStream stream) throws UnsupportedExcelDataException, FileErrorException, FileNotFoundException, IOException {
		if (titles == null || titles.size() == 0 || data == null || StringUtils.isBlank(modulePath) || stream == null) {
			throw new UnsupportedExcelDataException(UnsupportedExcelDataException.ILLEGAL_PARAMS);
		}
		ListDataModule mod = getListDataModule(modulePath);
		boolean isXls = mod.getFormat() == ExcelFileFormat.XLS;
		if ((isXls && titles.size() > EXCEL_XLS_COLUMN_COUNT_MAX) || (!isXls && titles.size() > EXCEL_XLSX_COLUMN_COUNT_MAX)) {
			throw new UnsupportedExcelDataException(UnsupportedExcelDataException.EXCEED_COLUMN_COUNT_LIMIT_ERROR);
		}
		if ((isXls && data.size() > EXCEL_XLS_ROW_COUNT_MAX) || (!isXls && data.size() > EXCEL_XLSX_ROW_COUNT_MAX)) {
			throw new UnsupportedExcelDataException(UnsupportedExcelDataException.EXCEED_ROW_COUNT_LIMIT_ERROR);
		}
		mod.export(titles, data, stream);
	}
	
	/**
	 * 
	 * 按指定的Xls模板导出不规则数据集, 可以输出多个Sheet, 可以包含合并单元格、图片、计算公式
	 * <p>
	 * bean通常是一个统计Bean对象, 私有属性对应导出模板中的每一个输出项
	 * bean对应的模板是以Bean类名命名, 有分别对应xls和xlsx格式的模板
	 * </p>
	 * 
	 * @param bean			封装各项数据的Bean对象
	 * @param modulePath	模板路径
	 * @param stream		输出流
	 * @throws UnsupportedExcelDataException
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void write(Serializable bean, String modulePath, OutputStream stream) throws UnsupportedExcelDataException, FileErrorException, FileNotFoundException, IOException {
		if (bean == null || StringUtils.isBlank(modulePath) || stream == null) {
			throw new UnsupportedExcelDataException(UnsupportedExcelDataException.ILLEGAL_PARAMS);
		}
		ELModule mod = getELModule(modulePath);
		mod.export(bean, stream);
	}
	
	/**
	 * 导出大数据记录（仅支持xlsx[2007+]）
	 * 
	 * @param DateHandler 数据处理方式
	 * 
	 * @param stream 输出流
	 */
	public static void write(DataFetcher fetcher, OutputStream stream) throws UnsupportedExcelDataException, FileNotFoundException, IOException {
		if (fetcher == null || fetcher.getTitles() == null || fetcher.getTitles().size() == 0 || stream == null) {
			throw new UnsupportedExcelDataException(UnsupportedExcelDataException.ILLEGAL_PARAMS);
		}
		BigDataXlsxExport.export(fetcher, stream);
		log.info("Excel数据导出完成");
	}
	
	/**
	 * 获取列表模板对象
	 * 
	 * @param modulePath	导出模板
	 * @return
	 * @throws FileNotFoundException
	 * @throws FileErrorException
	 * @throws IOException
	 */
	private static ListDataModule getListDataModule(String modulePath) throws FileNotFoundException, FileErrorException, IOException {
		return new PoiListDataModule(modulePath);
	}
	
	/**
	 * 获取支持EL表达式的模板对象
	 * 
	 * @param modulePath	导出模板
	 * @return
	 * @throws FileNotFoundException
	 * @throws FileErrorException
	 * @throws IOException
	 */
	private static ELModule getELModule(String modulePath) throws FileNotFoundException, FileErrorException, IOException {
		return new PoiELModule(modulePath);
	}
	
}
