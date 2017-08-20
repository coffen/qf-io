package com.qf.io.excel.writer.module;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 模板El表达式配置
 * <br>
 * File Name: WorkbookConfig.java
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
public class WorkbookConfig {
	
	
	private class SheetConfig {
		
	}
	
	private class RowConfig {
		
		List<ElCell> staticElCells = new ArrayList<ElCell>();	// 静态表达式列表
		List<ElCell> dynamicElCells = new ArrayList<ElCell>();	// 动态表达式列表
		List<ElCell> formulaElCells = new ArrayList<ElCell>();	// 公式表达式列表
		
		public List<ElCell> getStaticElCells() {
			return Collections.unmodifiableList(staticElCells);
		}
		
		public List<ElCell> getDynamicElCells() {
			return Collections.unmodifiableList(dynamicElCells);
		}
		
		public List<ElCell> getFormulaElCells() {
			return Collections.unmodifiableList(formulaElCells);
		}
		
		public void addElCell(ElCell cell) {
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
	
	private class ElCell {
		
		Integer rowIndex;		// 行索引
		Integer columnIndex;	// 列索引
		
		String srcExpr;			// 原生字符串
		List<String> elList;	// 表达式列表
		
		ElCellType type;		// El表达式类型
		
		public Integer getRowIndex() {
			return rowIndex;
		}
		
		public void setRowIndex(Integer rowIndex) {
			this.rowIndex = rowIndex;
		}
		
		public Integer getColumnIndex() {
			return columnIndex;
		}
		
		public void setColumnIndex(Integer columnIndex) {
			this.columnIndex = columnIndex;
		}
		
		public String getSrcExpr() {
			return srcExpr;
		}
		
		public void setSrcExpr(String srcExpr) {
			this.srcExpr = srcExpr;
		}
		
		public List<String> getElList() {
			return elList;
		}
		
		public void setElList(List<String> elList) {
			this.elList = elList;
		}
		
		public ElCellType getType() {
			return type;
		}
		
		public void setType(ElCellType type) {
			this.type = type;
		}
		
	}
	
	private enum ElCellType {
		
		STATIC, DYNAMIC, FORMULA;
		
	}
}
