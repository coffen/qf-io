package com.qf.io.excel.writer.module;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.util.List;
import java.util.concurrent.locks.ReentrantLock;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.qf.io.FileErrorException;
import com.qf.io.ModuleParseException;
import com.qf.io.excel.ExcelFileFormat;
import com.qf.io.excel.PoiUtils;
import com.qf.io.excel.writer.ELModule;
import com.qf.io.excel.writer.module.ElModuleConfig.ElCell;
import com.qf.io.excel.writer.module.ElModuleConfig.ElCellType;
import com.qf.io.excel.writer.module.ElModuleConfig.ElRow;
import com.qf.io.excel.writer.module.ElModuleConfig.ElSheet;
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
public class PoiELModule2 implements ELModule {
	
	private static final Pattern REGE_PATTERN = Pattern.compile("(\\$|#)\\{\\w+(\\[\\?\\])?(\\.\\w+(\\[\\?\\])?)*}");	

	private Workbook workbook;
	private ExcelFileFormat format;	
	private String modulePath;
	
	private ElModuleConfig config;
	
	private boolean hasInitialized = false;
	
	private ReentrantLock lock = new ReentrantLock();
	
	public PoiELModule2(String path) throws Exception {
		this.modulePath = path;
		
		init();
	}
	
	private void init() throws Exception {
		if (hasInitialized) {
			return;
		}
		
		if (lock.tryLock()) {
			if (!hasInitialized) {
				try {
					parse();
					hasInitialized = true;
				}
				catch (Exception e) {
					throw e;
				}
				finally {
					lock.unlock();
				}
			}
			else {
				lock.unlock();
			}
		}
	}
	
	public String getModulePath() {
		return modulePath;
	}

	@Override
	public void parse() throws FileErrorException, IOException, ModuleParseException {
		loadModule();	// 加载模板
		
		config = new ElModuleConfig();
		
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			ElSheet elSheet = parseSheet(i);
			if (elSheet != null) {
				config.addElSheet(elSheet);
			}
		}
	}
	
	private ElSheet parseSheet(int idx) throws ModuleParseException {
		ElSheet elSheet = null;
		
		Sheet sheet = workbook.getSheetAt(idx);
		if (sheet != null) {
			elSheet = new ElSheet(idx);
			
			if (sheet.getPhysicalNumberOfRows() > 0) {
				// 计算合并单元格的区间
		        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
		            CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
		            elSheet.addMergedCell(new int[] { cellRangeAddress.getFirstRow(), cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastRow(),  cellRangeAddress.getLastColumn() });
		        }
		        // 遍历解析单元格的表达式
				int lastRowIndex = sheet.getLastRowNum();
				for (int i = 0; i <= lastRowIndex; i++) {
					Row row = sheet.getRow(i);
					ElRow elRow = parseRow(row, elSheet);
					if (elRow != null) {
						elSheet.addElRow(elRow);
					}					
				}
			}
		}
		return elSheet;
	}
	
	private ElRow parseRow(Row row, ElSheet sheet) throws ModuleParseException {
		ElRow elRow = null;
		if (row != null) {
			Cell cell = null;
			int[] point = null;
			for (int i = 0; i < row.getLastCellNum(); i++) {
				cell = row.getCell(i);
				if (cell == null) {
					continue;
				}
				ElCell elCell = parseCell(cell);
				if (elCell == null) {
					continue;
				}
				if (elRow == null) {
					elRow = new ElRow(row.getRowNum());
				}
				elRow.addElCell(elCell);
				point = new int[] { cell.getRowIndex(), cell.getColumnIndex() };
				// 建立合并单元格的映射关系
				for (int[] mCell : sheet.getMergedCells()) {
					if (point[0] >= mCell[0] && point[0] <= mCell[2] && point[1] >= mCell[1] && point[1] <= mCell[3]) {
						sheet.putMergedRegion(point, mCell);
						break;
					}
				}
			}
		}
		return elRow;
	}
	
	private ElCell parseCell(Cell cell) throws ModuleParseException {
		ElCell elCell = null;
		if (cell != null) {
			int rowIndex = cell.getRowIndex();
			int columnIndex = cell.getColumnIndex();
			String srcExpr = cell.getStringCellValue();
			ElCellType type = null;
			Matcher matcher = REGE_PATTERN.matcher(srcExpr);
			if (matcher.matches() && srcExpr.startsWith("#")) {
				type = ElCellType.FORMULA;
			}
			else {
				while (matcher.find()) {
					String group = matcher.group(0);
					if (srcExpr.startsWith("#")) {
						throw new ModuleParseException("Formula expr error.");
					}
					if (group.indexOf("[?]") > 0) {
						type = ElCellType.DYNAMIC;
						break;
					}
					else {
						type = ElCellType.STATIC;
					}
				}
			}
			if (type != null) {
				elCell = new ElCell(rowIndex, columnIndex, srcExpr, type);				
			}
		}
		return elCell;
	}

	@Override
	public ExcelFileFormat getFormat() {
		return format;
	}

	@Override
	public String transformToCss() {
		return null;
	}

	@Override
	public void export(OutputStream stream, Serializable... beans) throws IOException {
		int sheetCount = workbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			Serializable bean = (beans != null && beans.length > i) ? beans[i] : null;
			ElSheet elSheet = config.getElSheet(i);
			Sheet sheet = workbook.getSheetAt(i);
			exportSheet(elSheet, sheet, bean);
		}
		// 清除Sheet模板
	}
	
	private void exportSheet(ElSheet sheetConfig, Sheet moduleSheet, Serializable bean) {
		Sheet targetSheet = null;
		
		if (bean != null) {
			targetSheet = workbook.createSheet();
			// 设置解析器
			SpelExprParsor parsor = new SpelExprParsor();
			parsor.setRootVariable(bean);
			
			for (int i = 0; i <= moduleSheet.getLastRowNum(); i++) {
				exportRow(sheetConfig, i, targetSheet, parsor);
			}
		}
		else {
			targetSheet = workbook.cloneSheet(sheetConfig.getSheetIndex());
		}
	}
	
	private int exportRow(ElSheet sheetConfig, int rowIndex, Sheet targetSheet, SpelExprParsor parsor) {
		int newRowCount = 0;
		Sheet moduleSheet = workbook.getSheetAt(sheetConfig.getSheetIndex());
		Row srcRow = moduleSheet.getRow(rowIndex);
		if (srcRow == null) {
			targetSheet.createRow(rowIndex);
		}
		ElRow elRow = sheetConfig.getElRow(rowIndex);
		if (elRow != null && elRow.getRowIndex() == rowIndex) {
			List<ElCell> dynamicElCells = elRow.getDynamicElCells();
			List<ElCell> staticElCells = elRow.getStaticElCells();
			List<ElCell> formulaElCells = elRow.getFormulaElCells();
			
			if (CollectionUtils.isNotEmpty(dynamicElCells)) {
				ElCell firstDynamicElCell = dynamicElCells.get(0);
				// 首个单元格的表达式（计算动态行长度）
				String listExpr = firstDynamicElCell.getSrcExpr();
				// 获取动态行的长度
				int loopCount = parsor.getListObjectSize(listExpr);
				elRow.setDynamicRowCount(loopCount);
				if (loopCount > 0) {
					int insertCount = PoiUtils.copyRows(moduleSheet, rowIndex, targetSheet, rowIndex, rowIndex + loopCount - 1, false);
					newRowCount = insertCount - 1;
					// 逐行赋值
					for (int i = rowIndex; i <= rowIndex + insertCount; i++) {
						Row tmpRow = targetSheet.getRow(i);
						for (ElCell elCell : dynamicElCells) {
							String cellExpr = elCell.getSrcExpr();
							// 存储动态表达式单元格的起始坐标（用于公式表达式的解析）
							if (i == rowIndex) {
								String[] _elArr = parsor.getElExpressions(cellExpr, true);
								if (_elArr != null && _elArr.length > 0) {
									sheetConfig.putDynamicRegion(_elArr[0], new int[] { rowIndex, rowIndex + insertCount, elCell.getColumnIndex() });
								}
							}
							cellExpr = cellExpr.replace("[?]", "[" + String.valueOf(i - rowIndex) + "]");
							Object _cellVal = parsor.getValue(cellExpr);
							Cell _cell = tmpRow.getCell(elCell.getColumnIndex());
							if (_cellVal instanceof byte[]) {
								// 图片处理
								int[] _region = sheetConfig.getMergedRegion(new int[] { elCell.getRowIndex(), elCell.getColumnIndex() });
								if (_region == null) {
									_region =  new int[] { _cell.getRowIndex(), _cell.getColumnIndex(), _cell.getRowIndex() + 1, _cell.getColumnIndex() + 1 };
								}
								PoiUtils.renderImage(workbook, 0, _region, (byte[])_cellVal);
							}
							else if (_cellVal != null) {
								PoiUtils.assignValue(_cell, _cellVal);
							}
						}						
					}					
				}
				else {
					
				}
			}
			else {
				
			}
		}
		else {
			PoiUtils.copyRows(moduleSheet, rowIndex, targetSheet, rowIndex, rowIndex, false);
		}
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
