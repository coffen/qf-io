package com.qf.io.excel.writer.module;

import java.util.List;

import com.qf.io.excel.writer.ELModule;

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
	
	private WorkbookConfig config;
	
	private class WorkbookConfig {
		
	}
	
	private class SheetConfig {
		
	}
	
	private class RowConfig {
		
		List<ElCell> staticElCells;
		List<ElCell> dynamicElCells;
		List<ElCell> formulaElCells;
		
		public List<ElCell> getStaticElCells() {
			return staticElCells;
		}
		
		public void setStaticElCells(List<ElCell> staticElCells) {
			this.staticElCells = staticElCells;
		}
		
		public List<ElCell> getDynamicElCells() {
			return dynamicElCells;
		}
		
		public void setDynamicElCells(List<ElCell> dynamicElCells) {
			this.dynamicElCells = dynamicElCells;
		}
		
		public List<ElCell> getFormulaElCells() {
			return formulaElCells;
		}
		
		public void setFormulaElCells(List<ElCell> formulaElCells) {
			this.formulaElCells = formulaElCells;
		}
		
	}
	
	private class ElCell {
		
		Integer rowIndex;		// 行索引
		Integer columnIndex;	// 列索引
		
		String srcExpr;			// 原生字符串
		List<String> elList;	// 表达式列表
		
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
		
	}

}
