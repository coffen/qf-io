package com.qf.io.excel.writer.module;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Pattern;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qf.io.FileErrorException;
import com.qf.io.excel.ExcelFileFormat;
import com.qf.io.excel.PoiUtils;
import com.qf.io.excel.writer.ELModule;
import com.qf.io.util.SpelExprParsor;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Poi表达式语言导出模板实现类
 * <br>
 * File Name: PoiELModule.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:15:22 
 * @version: v1.0
 *
 */
public class PoiELModule implements ELModule {
	
	private static final Pattern valueRegePattern = Pattern.compile("\\$\\{\\w+(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*(\\.\\w+\\[\\?\\])?(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*\\}");	
	private static final Pattern formulaRegePattern = Pattern.compile("#\\{\\w+(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*(\\.\\w+\\[\\?\\])?(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*\\}");

	private Workbook workbook;
	private ExcelFileFormat format;
	
	private String modulePath;
	
	private boolean hasInitialized = false;
	
	/**
	 * 单元格int数组代表每个单元格的坐标: int[rowIndex, columnIndex]
	 */
	private List<int[]> staticExprCells;   // 表达式单元格
	private List<List<int[]>> dynamicExprCells;  // 动态表达式单元格, 格式：List<表达式行<表达式单元格>>
	private List<Cell> formulaExprCells;  // 公式单元格（用于动态表达式的合计运算）
	
	/**
	 * <p>合并单元格区域映射表</p>
	 * 
	 * 格式：Map<int[rowIndex, columnIndex], int[startRowIndex, startColumnIndex, endRowIndex, endColumnIndex]>
	 * 当表达式单元格是一个合并单元格式, 记录其区域范围
	 */
	private Map<int[], int[]> regionMap;
	
	public PoiELModule(String path) throws FileNotFoundException, FileErrorException, IOException {
		this.modulePath = path;
		
		init();
	}

	public String getModulePath() {
		return modulePath;
	}
	
	public ExcelFileFormat getFormat() {
		return format;
	}
	
	private void init() throws FileErrorException, IOException {
		staticExprCells = new ArrayList<int[]>();
		dynamicExprCells = new ArrayList<List<int[]>>();
		formulaExprCells = new ArrayList<Cell>();
		regionMap = new HashMap<int[], int[]>();
		
		parse();
	}

	@Override
	public void parse() throws FileErrorException, IOException {
		if (hasInitialized) {
			return;
		}
		
		loadModule();	// 加载模板
		
		Sheet sheet = workbook.getSheetAt(0);  // 目前仅支持单个Sheet的运算
		if (sheet.getPhysicalNumberOfRows() == 0) {
			return;
		}
		
		// 计算合并单元格的区间
		List<int[]> mergedCells = new ArrayList<int[]>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
            mergedCells.add(new int[] { cellRangeAddress.getFirstRow(), cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastRow(),  cellRangeAddress.getLastColumn() });
        }
		
        // 遍历解析单元格的表达式
		int lastRowIndex = sheet.getLastRowNum();
		Row _row = null;
		Cell _cell = null;
		String _expr = null;
		int[] _point = null;
		List<int[]> _dynamicRow = null;
		for (int i = 0; i <= lastRowIndex; i++) {
			_row = sheet.getRow(i);
			_dynamicRow = null;
			if (_row.getPhysicalNumberOfCells() == 0) {
				continue;
			}
			for (int j = 0; j < _row.getLastCellNum(); j++) {
				_cell = _row.getCell(j);
				_expr = _cell.getStringCellValue();
				if (StringUtils.isBlank(_expr)) {
					continue;
				}
				_point = new int[] { _cell.getRowIndex(), _cell.getColumnIndex() };
				if (valueRegePattern.matcher(_expr).find()) {
					if (_expr.indexOf("[?]") > 0) {
						if (_dynamicRow == null) {
							_dynamicRow = new ArrayList<int[]>();
						}
						_dynamicRow.add(_point);
					} 
					else {
						staticExprCells.add(_point);
					}
				} 
				else if (formulaRegePattern.matcher(_expr).find()) {
					formulaExprCells.add(_cell);
				} 
				else {
					continue;
				}
				// 建立合并单元格的映射关系
				for (int[] _mCell : mergedCells) {
					if (_point[0] >= _mCell[0] && _point[0] <= _mCell[2] && _point[1] >= _mCell[1] && _point[1] <= _mCell[3]) {
						regionMap.put(_point, _mCell);
					}
				}
			}
			if (_dynamicRow != null) {
				dynamicExprCells.add(_dynamicRow);
			}
		}		
		hasInitialized = true;
	}

	@Override
	public void export(Serializable bean, OutputStream stream) throws IOException {
		Sheet sheet = workbook.getSheetAt(0);
		if (bean != null) {
			// 设置解析器
			SpelExprParsor parsor = new SpelExprParsor();
			parsor.setRootVariable(bean);
			
			Cell _cell = null;
			Object _cellVal = null;
			int[] _region = null;
			// 表达式单元格处理
			if (CollectionUtils.isNotEmpty(staticExprCells)) {
				for (int[] point : staticExprCells) {
					if (point == null || point.length != 2) {
						continue;
					}
					_cell = sheet.getRow(point[0]).getCell(point[1]);
					if (_cell == null) {
						continue;
					}
					_cellVal = parsor.getValue(_cell.getStringCellValue());
					if (_cellVal instanceof byte[]) {
						// 图片处理
						if (regionMap.containsKey(point)) {
							_region = regionMap.get(point);
						} 
						else {
							_region = new int[] { _cell.getRowIndex(), _cell.getColumnIndex(), _cell.getRowIndex() + 1, _cell.getColumnIndex() + 1 };
						}
						PoiUtils.renderImage(workbook, 0, _region, (byte[])_cellVal);
					}
					else if (_cellVal != null) {
						PoiUtils.assignValue(_cell, _cellVal);
					}
				}
			}
			// 动态表达式单元格处理
			Map<String, int[]> dynamicCellsRegion = new HashMap<String, int[]>();
			if (CollectionUtils.isNotEmpty(dynamicExprCells)) {
				int loopCount = 0;
				int[] startPoint = null;
				String listExpr, cellExpr = null;
				Row srcRow, tmpRow = null;
				int gap = 0; // 当前一个动态行完成遍历后，后面的行的行坐标实际已经发生了变化, 该变量记录差值
				for (List<int[]> dynalist : dynamicExprCells) {
					if (dynalist == null || dynalist.size() == 0) {
						continue;
					}
					startPoint = dynalist.get(0);
					if (startPoint == null || startPoint.length != 2) {
						continue;
					}
					srcRow = sheet.getRow(startPoint[0] + gap);
					// 首个单元格的表达式（用于计算动态行长度）
					listExpr = srcRow.getCell(startPoint[1]).getStringCellValue();
					// 获取动态行的长度
					loopCount = parsor.getListObjectSize(listExpr);
					if (loopCount > 0) {
						int insertCount = PoiUtils.copyRows(sheet, startPoint[0], startPoint[0] + 1, startPoint[0] + loopCount, false);
						// 逐行赋值
						for (int i = startPoint[0] + 1; i <= startPoint[0] + insertCount; i++) {
							tmpRow = sheet.getRow(i);
							String[] _elArr = null;
							for (int[] _point : dynalist) {
								cellExpr = srcRow.getCell(_point[1]).getStringCellValue();
								// 存储动态表达式单元格的起始坐标（用于公式表达式的解析）
								if (i == startPoint[0] + 1) {
									_elArr = parsor.getElExpressions(cellExpr, true);
									if (_elArr != null && _elArr.length > 0) {
										dynamicCellsRegion.put(_elArr[0], new int[] { startPoint[0] + 1, _point[1], startPoint[0] + insertCount, _point[1] });
									}
								}
								cellExpr = cellExpr.replace("[?]", "[" + String.valueOf(i - startPoint[0] - 1) + "]");
								_cellVal = parsor.getValue(cellExpr);
								_cell = tmpRow.getCell(_point[1]);
								if (_cellVal instanceof byte[]) {
									// 图片处理
									if (regionMap.containsKey(_point)) {
										_region = regionMap.get(_point);
									} 
									else {
										_region = new int[] { _cell.getRowIndex(), _cell.getColumnIndex(), _cell.getRowIndex() + 1, _cell.getColumnIndex() + 1 };
									}
									PoiUtils.renderImage(workbook, 0, _region, (byte[])_cellVal);
								}
								else if (_cellVal != null) {
									PoiUtils.assignValue(_cell, _cellVal);
								}
							}						
						}
					}
					// 动态行遍历完毕必须删除源数据行
					sheet.shiftRows(startPoint[0] + 1, sheet.getLastRowNum(), -1);
					
					gap = loopCount - 1;
				}
			}
			// 公式表达式单元格处理
			String _formulaExpr;
			if (CollectionUtils.isNotEmpty(formulaExprCells)) {
				String[] _elArr = null;
				int[] _point = null;
				String _columnName = null;
				for (Cell _formulaCell : formulaExprCells) {
					if (_formulaCell == null) {
						continue;
					}
					_formulaExpr = _formulaCell.getStringCellValue();
					_elArr = parsor.getElExpressions(_formulaExpr, false);
					if (_elArr != null && _elArr.length > 0) {
						for (String _el : _elArr) {
							if (dynamicCellsRegion.containsKey(_el)) {
								_point =  dynamicCellsRegion.get(_el);
								_columnName = PoiUtils.translateColumnIndex2Name(_point[1]);
								_formulaExpr = _formulaExpr.replace("#{" + _el + "}", _columnName + _point[0] + ":" + _columnName + _point[2]);
							}
						}
						_formulaCell.setCellValue("");
						_formulaCell.setCellFormula(_formulaExpr);
					}
				}
			}
		}
		
		workbook.write(stream);
		stream.close();
		stream = null;
	}
	
	@Override
	public String transformToCss() {
		return null;
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
