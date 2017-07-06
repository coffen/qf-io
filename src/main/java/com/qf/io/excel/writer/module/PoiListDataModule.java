package com.qf.io.excel.writer.module;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qf.io.FileErrorException;
import com.qf.io.excel.ExcelFileFormat;
import com.qf.io.excel.PoiUtils;
import com.qf.io.excel.writer.ListDataModule;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 通过POI处理List类型数据导出的模板
 * <br>
 * File Name: PoiListDataModule.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:15:32 
 * @version: v1.0
 *
 */
public class PoiListDataModule implements ListDataModule {
	
	private Workbook workbook;
	private ExcelFileFormat format;
	
	private String modulePath;
	
	private Short headerHeight;
	private Short rowHeight;
	
	private CellStyle headerStyle;
	private CellStyle[][] bodyStyle;
	
	public PoiListDataModule(String path) throws FileNotFoundException, FileErrorException, IOException {
		this.modulePath = path;
		
		init();
	}
	
	public String getModulePath() {
		return modulePath;
	}
	
	public ExcelFileFormat getFormat() {
		return format;
	}
	
	public Short getHeaderHeight() {
		return headerHeight;
	}
	
	public Short getRowHeight() {
		return rowHeight;
	}
	
	public CellStyle getHeaderStyle() {
		return headerStyle;
	}
	
	public CellStyle[][] getBodyStyle() {
		return bodyStyle;
	}
	
	private void init() throws FileErrorException, IOException {
		parse();
	}
	
	@Override
	public void parse() throws FileErrorException, IOException {
		loadModule();	// 加载模板
		
		Sheet sheet = workbook.getSheetAt(0);
		if (sheet.getPhysicalNumberOfRows() == 0) {
			return;
		}
		int lastRowIndex = sheet.getLastRowNum();
		int i = 0;
		Row _row = null;
		for (; i <= lastRowIndex; i++) {
			_row = sheet.getRow(i);
			if (_row.getPhysicalNumberOfCells() == 0) {
				continue;
			}
			headerHeight = _row.getHeight();
			Cell tmp = PoiUtils.getFirstNotNullCell(_row);
			if (tmp != null) {
				headerStyle = tmp.getCellStyle();
				break;
			}
		}
		if (i < lastRowIndex) {
			int[] region = PoiUtils.getMaxPhysicalRegion(sheet, 1, lastRowIndex);
			if (region != null && region.length == 2 && region[0] > 0 && region[1] > 0) {
				// 表体列样式的设置（0：无样式表体, 1: 单一设置的表体, 2: 包括主栏的表体, 3: 设置尾列的表体）
				bodyStyle = new CellStyle[region[0]][region[1]];
				int x = 0;
				Row row = null;
				for (i = i + 1; i <= lastRowIndex; i++) {
					row = sheet.getRow(i);
					int y = 0;
					if (row.getPhysicalNumberOfCells() == 0) {
						continue;
					}
					if (x == 0) {
						rowHeight = row.getHeight();
					}
					for (int j = 0; j < row.getLastCellNum(); j++) {
						if (StringUtils.isNotBlank(row.getCell(j).getStringCellValue())) {
							bodyStyle[x][y] = row.getCell(j).getCellStyle();
							y++;
						}
					}
					x++;
				}
			}
		}
	}

	@Override
	public void export(LinkedHashMap<String, String> titles, List<Map<String, ?>> data, OutputStream stream) throws IOException {
		Sheet sheet = workbook.createSheet();
		// 表头渲染
		Row header = sheet.createRow(0);
		int idx = 0;
		String _title = null;
		Cell _headerCell = null;
		boolean hideHeader = true;
		for (Iterator<String> iterator = titles.values().iterator(); iterator.hasNext(); idx++) {
			_headerCell = header.createCell(idx);
			_title = iterator.next();
			_headerCell.setCellValue(_title);
			if (headerStyle != null) {
				_headerCell.setCellStyle(headerStyle);
			}
			hideHeader = hideHeader && StringUtils.isBlank(_title);
		}
		if (hideHeader) {
			header.setZeroHeight(true); // 隐藏表头
		} 
		else if (headerHeight != null) {
			header.setHeight(headerHeight);
		}
		
		// 表体渲染
		if (data != null && data.size() > 0) {
			Row _row = null;
			Map<String, ?> _rowData = null;
			CellStyle[] _rowStyles = null;  // 当前行应用的样式（单元格样式数组）
			Cell _bodyCell = null;
			String _key;
			for (int i = 0; i < data.size(); i++) {
				_rowData = data.get(i);
				_row = sheet.createRow(i + 1);
				if (bodyStyle != null && bodyStyle.length > 0) {
					int _styleIndex = (i + 1) % bodyStyle.length - 1;
					if (_styleIndex == -1) {
						_styleIndex = bodyStyle.length - 1;
					}
					_rowStyles = bodyStyle[_styleIndex];
				}
				int j = 0;
				for (Iterator<String> iterator = titles.keySet().iterator(); iterator.hasNext(); j++) {
					_bodyCell = _row.createCell(j);
					_key = iterator.next();
					PoiUtils.assignValue(_bodyCell, _rowData.get(_key));
					if (_rowStyles != null && _rowStyles.length > 0) {
						if (_rowStyles.length == 1) {
							if (_rowStyles[0] != null) {
								_bodyCell.setCellStyle(_rowStyles[0]);
							}
						} 
						else if (_rowStyles.length == 2) {
							if (j == 0) {
								if (_rowStyles[0] != null) {
									_bodyCell.setCellStyle(_rowStyles[0]);
								}
							} 
							else {
								if (_rowStyles[1] != null) {
									_bodyCell.setCellStyle(_rowStyles[1]);
								}
							}
						} 
						else {
							if (j == 0) {
								if (_rowStyles[0] != null) {
									_bodyCell.setCellStyle(_rowStyles[0]);
								}
							} 
							else if (j == titles.size()) {
								if (_rowStyles[2] != null) {
									_bodyCell.setCellStyle(_rowStyles[2]);
								}
							} 
							else {
								if (_rowStyles[1] != null) {
									_bodyCell.setCellStyle(_rowStyles[1]);
								}
							}
						}
					}					
				}
			}
		}
		
		// 宽度调整
		for (int i = 0; i < titles.size(); i++){
			sheet.autoSizeColumn(i); // 自适应宽度, 含汉字的列略有不准
			sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 2);
		}
		
		// 删除模板Sheet
		workbook.removeSheetAt(0);
		
		workbook.write(stream);
		stream.close();
		stream = null;
	}
	
	@Override
	public String transformToCss() {
		StringBuffer cssText = new StringBuffer();
		String lineSeparator = System.getProperty("line.separator");
		String moduleName = FilenameUtils.getBaseName(modulePath);
		String tableCssName = "table." + moduleName;
		cssText.append(tableCssName).append(" { border-collapse: collapse; border: 0px; }").append(lineSeparator);
		cssText.append(tableCssName).append(" td { padding-left: 3px; }").append(lineSeparator);
		if (headerStyle != null) {
			cssText.append(tableCssName).append(" > thead td {");
			cssText.append("height:").append(headerHeight / 20).append("px;");
			cssText.append(PoiUtils.transformCellStyleToCss(workbook, headerStyle));
			cssText.append("}").append(lineSeparator);
		}
		if (bodyStyle != null) {
			int x = bodyStyle.length;
			int y = bodyStyle[0] != null ? bodyStyle[0].length : 0;			
			String tdStyleName = null;
			for (int i = 0; i < x; i++) {
				for (int j = 0; j < y; j++) {
					if (bodyStyle[i][j] == null) {
						continue;
					}
					tdStyleName = j == 0 ? "td:first-child" : (j == 1 ? "td" : "td:last-child");
					cssText.append(tableCssName).append(" > tbody tr:nth-child(").append(x).append("n+").append(i == x - 1 ? 0 : i + 1).append(") ").append(tdStyleName).append(" {");
					cssText.append(PoiUtils.transformCellStyleToCss(workbook, bodyStyle[i][j]));
					cssText.append("}").append(lineSeparator);
				}
			}
		}
		return cssText.toString();
	}
	
	/**
	 * 加载模板文件
	 */
	private void loadModule() throws FileNotFoundException, FileErrorException, IOException {
		File file = new File(modulePath);
		if (!file.exists() || !file.isFile()) {
			throw new FileNotFoundException();
		}
		String ext = FilenameUtils.getExtension(modulePath);
		try {
			format = ExcelFileFormat.valueOf(ext.toUpperCase());
		}
		catch (Exception e) {
			throw new FileErrorException(ext, FileErrorException.FILE_EXTENSION_ERROR);
		}
		boolean isXls = format == ExcelFileFormat.XLS;
		workbook = isXls ? new HSSFWorkbook(new FileInputStream(modulePath)) : new XSSFWorkbook(new FileInputStream(modulePath));
	}

}
