package com.qf.io.excel.writer.module;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 模板El表达式配置
 * <br>
 * File Name: ElModuleConfig.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年8月20日 下午4:15:50 
 * @version: v1.0
 *
 */
public class ElModuleConfig {
	
	private List<ElSheet> elSheets = new ArrayList<ElSheet>();
	
	public List<ElSheet> getElSheets() {
		return Collections.unmodifiableList(elSheets);
	}
	
	public void addElSheet(ElSheet sheet) {
		if (sheet == null) {
			return;
		}
		elSheets.add(sheet);
	}
	
	static class ElSheet {
		
		int sheetIndex;

		List<ElRow> elRows = new ArrayList<ElRow>();
		
		List<int[]> mergedCells = new ArrayList<int[]>();
		
		/**
		 * <p>合并单元格区域映射表</p>
		 * 
		 * 格式：Map<int[rowIndex, columnIndex], int[startRowIndex, startColumnIndex, endRowIndex, endColumnIndex]>
		 * 当表达式单元格是一个合并单元格式, 记录其区域范围
		 */
		Map<int[], int[]> regionMap = new HashMap<int[], int[]>();
		
		public ElSheet(int sheetIndex) {
			this.sheetIndex = sheetIndex;
		}
		
		public int getSheetIndex() {
			return sheetIndex;
		}
		
		public List<ElRow> getElRows() {
			return Collections.unmodifiableList(elRows);
		}
		
		public List<int[]> getMergedCells() {
			return Collections.unmodifiableList(mergedCells);
		}
		
		public void addElRow(ElRow row) {
			if (row == null) {
				return;
			}
			elRows.add(row);
		}
		
		public void addMergedCell(int[] region) {
			if (region == null || region.length != 4) {
				return;
			}
			mergedCells.add(region);
		}
		
		public void putRegion(int[] point, int[] region) {
			if (point == null || region == null) {
				return;
			}
			regionMap.put(point, region);
		}
		
		public int[] getRegion(int[] point) {
			return regionMap.get(point);
		}
		
	}
	
	static class ElRow {
		
		int rowIndex;	// 行索引
		
		List<ElCell> staticElCells = new ArrayList<ElCell>();	// 静态表达式列表
		List<ElCell> dynamicElCells = new ArrayList<ElCell>();	// 动态表达式列表
		List<ElCell> formulaElCells = new ArrayList<ElCell>();	// 公式表达式列表
		
		public ElRow(int rowIndex) {
			this.rowIndex = rowIndex;
		}
		
		public int getRowIndex() {
			return rowIndex;
		}
		
		List<ElCell> getStaticElCells() {
			return Collections.unmodifiableList(staticElCells);
		}
		
		List<ElCell> getDynamicElCells() {
			return Collections.unmodifiableList(dynamicElCells);
		}
		
		List<ElCell> getFormulaElCells() {
			return Collections.unmodifiableList(formulaElCells);
		}
		
		void addElCell(ElCell cell) {
			if (cell == null) {
				return;
			}
			if (cell.getType() == ElCellType.DYNAMIC) {
				dynamicElCells.add(cell);
			}
			else if (cell.getType() == ElCellType.FORMULA) {
				formulaElCells.add(cell);
			}
			else {
				staticElCells.add(cell);
			}
		}
		
	}
	
	static class ElCell {
		
		int rowIndex;			// 行索引
		int columnIndex;		// 列索引		
		String srcExpr;			// 原生字符串		
		ElCellType type;		// El表达式类型
		
		public ElCell(int rowIndex, int columnIndex, String srcExpr, ElCellType type) {
			this.rowIndex = rowIndex;
			this.columnIndex = columnIndex;
			
			this.srcExpr = srcExpr;
			
			this.type = type;
		}
		
		public int getRowIndex() {
			return rowIndex;
		}
		
		public int getColumnIndex() {
			return columnIndex;
		}
		
		public ElCellType getType() {
			return type;
		}
		
		public String getSrcExpr() {
			return srcExpr;
		}
		
	}
	
	public static enum ElCellType {
		
		STATIC, DYNAMIC, FORMULA;
		
	}

}
