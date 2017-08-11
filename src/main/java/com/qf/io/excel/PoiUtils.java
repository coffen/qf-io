package com.qf.io.excel;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder.BorderSide;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: POI读写工具类
 * <br>
 * File Name: PoiUtils.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:17:37 
 * @version: v1.0
 *
 */
public class PoiUtils {
	
	private final static Logger log = LoggerFactory.getLogger(PoiUtils.class);
	
	/**
	 * 获取当前Sheet的最大列数（即最后一列有数据的单元格的列数）
	 * 
	 * @param Sheet 指定Sheet
	 * 
	 * @return int 最大列数, -1代表错误参数, 0代表Sheet中没有数据
	 */
	public static int getMaxColumns(Sheet sheet) {
		if (sheet == null) {
			return -1;
		}
		return getMaxColumns(sheet, 0, sheet.getLastRowNum());
	}
	
	/**
	 * 获取当前Sheet指定区域内（开始行->结束行）的最大列数（即最后一列有数据的单元格的列数）
	 * <br>
	 * 开始行, 结束行都是闭区间
	 * 
	 * @param Sheet     指定Sheet
	 * @param startRow  开始行（zero-based）
	 * @param endRow    结束行（zero-based）
	 * 
	 * @return int 最大列数, -1代表错误参数, 0代表Sheet中没有数据
	 */
	public static int getMaxColumns(Sheet sheet, int startRow, int endRow) {
		int maxCols = -1;
		if (sheet == null || endRow < startRow) {
			return maxCols;
		}
		int lastRowIndex = sheet.getLastRowNum();
		if (startRow > lastRowIndex) {
			return maxCols;
		}
		int physicalRownum = sheet.getPhysicalNumberOfRows();
		if (physicalRownum == 0) {
			return 0;
		}
		endRow = endRow > lastRowIndex ? lastRowIndex : endRow;
		// 遍历行计算最大列数
		Row row = null;
		int lastColumnIndex = sheet.getLastRowNum() - 1;
		for (int i = startRow; i <= endRow; i++) {
			row = sheet.getRow(i);
			if (row.getPhysicalNumberOfCells() == 0) {
				continue;
			}
			for (int j = 0; j <= lastColumnIndex; j++) {
				if (j > maxCols && StringUtils.isNotBlank(row.getCell(j).getStringCellValue())) {
					maxCols = j;
				}
			}
		}
		return maxCols;
	}
	
	/**
	 * 获取当前Sheet指定区域内（开始行->结束行）的最大非空行数和最大非空列数（每行只计算有数据的单元格）
	 * <br>
	 * 开始行, 结束行都是闭区间
	 * 
	 * @param Sheet     指定Sheet
	 * @param startRow  开始行（zero-based）
	 * @param endRow    结束行（zero-based）
	 * 
	 * @return int[] 格式：int[行数, 列数], null代表没有数据
	 */
	public static int[] getMaxPhysicalRegion(Sheet sheet, int startRow, int endRow) {
		if (sheet == null || endRow < startRow) {
			return null;
		}
		int lastRowIndex = sheet.getLastRowNum();
		if (startRow > lastRowIndex) {
			return null;
		}
		int physicalRownum = sheet.getPhysicalNumberOfRows();
		if (physicalRownum == 0) {
			return null;
		}
		endRow = endRow > lastRowIndex ? lastRowIndex : endRow;
		// 遍历行计算最大有效列数
		int _cellNum = 0;
		int maxRow = 0, maxCol = 0;
		for (int i = startRow; i <= endRow; i++) {
			_cellNum = sheet.getRow(i).getPhysicalNumberOfCells();
			if (_cellNum > 0) {
				maxRow++;
			}
			if (_cellNum > maxCol) {
				maxCol = _cellNum;
			}
		}
		return new int[] {maxRow, maxCol};
	}
	
	/**
	 * 获取当前行第一个样式
	 * 
	 * @param row 指定行
	 * 
	 * @return Cell
	 */
	public static CellStyle getFirstCellStyle(Row row) {
		if (row == null || row.getPhysicalNumberOfCells() == 0) {
			return null;
		}
		int lastCellIndex = row.getLastCellNum() - 1;
		for (int i = 0; i <= lastCellIndex; i++) {
			if (row.getCell(i).getCellStyle() != null) {
				return row.getCell(i).getCellStyle();
			}
		}
		return null;
	}
	
	/**
	 * 获取当前行最后一个非空单元格
	 * 
	 * @param row 指定行
	 * 
	 * @return Cell
	 */
	public static Cell getLastNotNullCell(Row row) {
		if (row == null || row.getPhysicalNumberOfCells() == 0) {
			return null;
		}
		int lastCellIndex = row.getLastCellNum() - 1;
		for (int i = lastCellIndex; i >= 0; i--) {
			if (StringUtils.isNotBlank(row.getCell(i).getStringCellValue())) {
				return row.getCell(i);
			}
		}
		return null;
	}
	
	/**
	 * 判断是否合并单元格. 如果是, 返回起始行、起始列组成的数组; 否则返回当前单元格的起始行
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * 
	 * @return int[]  : [ firstRow, firstColumn, lastRow, lastColumn ]
	 */
	public int[] getCellRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return new int[] { firstRow, firstColumn, lastRow + 1, lastColumn + 1 };
				}
			}
		}
		return new int[] { row, column, row + 1, column + 1 };
	}
	
	/**
	 * 复制行（目前只支持同一Sheet的复制）
	 * 
	 * @param sheet            源工作单
	 * @param srcRow           源行号
	 * @param firstRow         拷贝目标首行行号（zero-based）
	 * @param lastRow          拷贝目标末行行号（zero-based）
	 * @param needCopyValue    是否需要复制目标单元格的值
	 * 
	 * @return int 插入的行数
	 * 
	 */
	public static int copyRows(Sheet sheet, int srcRow, int firstRow, int lastRow, boolean needCopyValue) {
		if (sheet == null || srcRow < 0 || firstRow < 0 || lastRow < 0 || srcRow >= sheet.getPhysicalNumberOfRows()) {
			return -1;
		}
		// 参数传反的处理
		int x = 0;
		if (lastRow < firstRow) {
			x = firstRow;
			firstRow = lastRow;
			lastRow = x;
		}
		// 计算复制行合并单元格的起始列位置
		List<int[]> mergedCells = new ArrayList<int[]>();
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = sheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() <= srcRow && cellRangeAddress.getLastRow() >= srcRow) {
            	mergedCells.add(new int[] { cellRangeAddress.getFirstColumn(), cellRangeAddress.getLastColumn() });
            }
        }		
		// 逐条插入行到指定位置, 如目标行存在，则需连同之后的行一起向下移动一行
		Row fromRow = sheet.getRow(srcRow);
		Row toRow = null;
		for (int i = firstRow; i <= lastRow; i++) {
			toRow = sheet.getRow(i);
			if (toRow == null) {
	        	toRow = sheet.createRow(i);
	        }
			sheet.shiftRows(i + 1, sheet.getLastRowNum(), 1);
			// 复制单元格并设置同样的样式、注解、数据等
			Cell srcCell, tmpCell = null;
			for (int j = fromRow.getFirstCellNum(); j < fromRow.getPhysicalNumberOfCells(); j++) {
				srcCell = fromRow.getCell(j);     // 源单元格
				tmpCell = toRow.createCell(j);    // 目标单元格
				if (srcCell == null) {
					tmpCell = null;
					continue;
				}
				// 样式设置
				tmpCell.setCellStyle(srcCell.getCellStyle());
				// 格式设置
				tmpCell.setCellType(srcCell.getCellType());
				// 是否允许赋值单元格数据
				if (needCopyValue) {
					// 数据类型设置
					switch (srcCell.getCellType()) {
	                case Cell.CELL_TYPE_BLANK:
	                	tmpCell.setCellValue(srcCell.getStringCellValue());
	                    break;
	                case Cell.CELL_TYPE_BOOLEAN:
	                	tmpCell.setCellValue(srcCell.getBooleanCellValue());
	                    break;
	                case Cell.CELL_TYPE_ERROR:
	                	tmpCell.setCellErrorValue(srcCell.getErrorCellValue());
	                    break;
	                case Cell.CELL_TYPE_FORMULA:
	                	tmpCell.setCellFormula(srcCell.getCellFormula());
	                    break;
	                case Cell.CELL_TYPE_NUMERIC:
	                	tmpCell.setCellValue(srcCell.getNumericCellValue());
	                    break;
	                case Cell.CELL_TYPE_STRING:
	                	tmpCell.setCellValue(srcCell.getRichStringCellValue());
	                    break;
	                }				
					// 链接设置
					if (srcCell.getHyperlink() != null) {
						tmpCell.setHyperlink(srcCell.getHyperlink());
		            }
					// 注解设置
					if (srcCell.getCellComment() != null) {
						tmpCell.setCellComment(srcCell.getCellComment());
		            }
				}
			}
			// 合并单元格
			if (mergedCells.size() > 0) {
				for (int[] columnRegion : mergedCells) {
					CellRangeAddress newCellRangeAddress = new CellRangeAddress(i, i, columnRegion[0], columnRegion[1]);
					sheet.addMergedRegion(newCellRangeAddress);
				}
			}
		}
		return lastRow - firstRow + 1;
	}
	
	/**
	 * 获取单元格的值, 只返回String, Double, Date, Boolean三种类型
	 * 
	 * @param cell  单元格
	 * 
	 * @return Object 返回对象, 如单元格为空则返回Null
	 */
	public static Object getValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		int cellType = cell.getCellType();
		switch (cellType) {
			case Cell.CELL_TYPE_BLANK:
				return null;
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			case Cell.CELL_TYPE_NUMERIC:
				return cell.getNumericCellValue();
			case Cell.CELL_TYPE_BOOLEAN:
				return cell.getBooleanCellValue();
			case Cell.CELL_TYPE_FORMULA:
				return cell.getStringCellValue();
			case Cell.CELL_TYPE_ERROR:
				return null;
			default:
				return null;
		}
	}
	
	/**
	 * 对单元格赋值, 目前单元格只支持String, Integer, Double, Date, Boolean五种类型
	 * 
	 * @param cell  待赋值单元格
	 * @param value 赋值对象
	 */
	public static void assignValue(Cell cell, Object value) {
		if (cell == null) {
			return;
		}
		if (value == null) {
			cell.setCellValue("");
			return;
		}
		if (value instanceof String) {
			cell.setCellValue((String)value);
		} 
		else if (value instanceof Integer) {
			cell.setCellValue((Integer)value);
		} 
		else if (value instanceof Double) {
			cell.setCellValue((Double)value);
		} 
		else if (value instanceof Date) {
			cell.setCellValue(new DateTime((Date)value).toString("yyyy-MM-dd"));
		} 
		else if (value instanceof Boolean) {
			cell.setCellValue((Boolean)value);
		}
		else {
			cell.setCellValue(value.toString());
		}
	}
	
	/**
	 * 在制定的区域内放置图片（图片统一设置为PNG格式）
	 * 
	 * @param work  文档
	 * @param sheetIndex  sheet页索引
	 * @param region  图片放置区域（int[ firstRowIndex, firstColumnIndex, lastRowIndex + 1, lastColumnIndex + 1]）
	 * @param value 赋值对象
	 */
	public static void renderImage(Workbook work, int sheetIndex, int[] region, byte[] imgData) {
		if (work == null || sheetIndex < 0 || region == null || region.length != 4 || imgData == null) {
			return;
		}
		Sheet sheet = null;
		try {
			sheet = work.getSheetAt(sheetIndex);
		} catch (Exception e) {
			log.error("图片渲染失败, 非法的SheetIndex: ", sheetIndex);
			return;
		}
		Drawing patriarch = sheet.createDrawingPatriarch();
		CreationHelper helper = work.getCreationHelper();
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setDx1(0);
		anchor.setDy1(0);
		anchor.setDx2(0);
		anchor.setDy2(sheet.getRow(region[0]).getHeight() * 10000);
		anchor.setRow1(region[0]);
		anchor.setCol1((short)region[1]);
		anchor.setRow2(region[2]);
		anchor.setCol2((short)region[3]);  
		patriarch.createPicture(anchor, work.addPicture(imgData, Workbook.PICTURE_TYPE_PNG));
	}
	
	/**
	 * 转换CellStyle为CSS样式字符串
	 * 
	 * @param cell  待赋值单元格
	 * @param value 赋值对象
	 * 
	 * @return String Css字符串
	 */
	public static String transformCellStyleToCss(Workbook work, CellStyle style) {
		if (style == null) {
			return null;
		}
		StringBuffer css = new StringBuffer();
		int format = work instanceof HSSFWorkbook ? 1 : 2;
		// 水平对齐
		short alignment = style.getAlignment();
		switch (alignment) {
		case CellStyle.ALIGN_LEFT:
			css.append("text-align:left;");
			break;
		case CellStyle.ALIGN_CENTER:
			css.append("text-align:center;");
			break;
		case CellStyle.ALIGN_RIGHT:
			css.append("text-align:right;");
			break;
		default:
			break;
		}
		// 垂直对齐
		short vAlignment = style.getVerticalAlignment();
		switch (vAlignment) {
		case CellStyle.VERTICAL_TOP:
			css.append("vertical-align:top;");
			break;
		case CellStyle.VERTICAL_CENTER:
			css.append("vertical-align:middle;");
			break;
		case CellStyle.VERTICAL_BOTTOM:
			css.append("vertical-align:bottom;");
			break;
		default:
			break;
		}
		// 字体
		Font font = work.getFontAt(style.getFontIndex());
		if (font != null) {
			css.append("font-weight:").append(font.getBoldweight()).append(";");
			css.append("font-family:").append(font.getFontName()).append(";");
			css.append("font-size:").append(font.getFontHeightInPoints()).append("px;");
			Color fontColor = null;
			if (format == 1) {
				fontColor = ((HSSFFont)font).getHSSFColor((HSSFWorkbook)work);
			} else {
				fontColor = ((XSSFFont)font).getXSSFColor();
			}
			css.append("color:").append(convertToStardColor(fontColor)).append(";");
		}
		// 背景颜色
		if (style.getFillPattern() == CellStyle.SOLID_FOREGROUND) {
			Color bgColor = style.getFillForegroundColorColor();
			css.append("background-color:").append(convertToStardColor(bgColor)).append(";");
		}
		// 边框样式, 颜色
		String[] borderArr = new String[] { "top", "bottom", "left", "right" };
		short borderStyle = 0;
		Color borderColor = null;
		for (String border : borderArr) {
			if ("top".equals(border)) {
				borderStyle = style.getBorderTop();
				if (format == 1) {
					borderColor = ((HSSFWorkbook)work).getCustomPalette().getColor(style.getTopBorderColor());
				} else {
					borderColor = ((XSSFCellStyle)style).getBorderColor(BorderSide.TOP);
				}
			} else if ("bottom".equals(border)) {
				borderStyle = style.getBorderBottom();
				if (format == 1) {
					borderColor = ((HSSFWorkbook)work).getCustomPalette().getColor(style.getBottomBorderColor());
				} else {
					borderColor = ((XSSFCellStyle)style).getBorderColor(BorderSide.BOTTOM);
				}
			} else if ("left".equals(border)) {
				borderStyle = style.getBorderLeft();
				if (format == 1) {
					borderColor = ((HSSFWorkbook)work).getCustomPalette().getColor(style.getLeftBorderColor());
				} else {
					borderColor = ((XSSFCellStyle)style).getBorderColor(BorderSide.LEFT);
				}
			} else {
				borderStyle = style.getBorderRight();
				if (format == 1) {
					borderColor = ((HSSFWorkbook)work).getCustomPalette().getColor(style.getRightBorderColor());
				} else {
					borderColor = ((XSSFCellStyle)style).getBorderColor(BorderSide.RIGHT);
				}
			}
			// 边框样式
			switch (borderStyle) {
			case CellStyle.BORDER_DASH_DOT:
				css.append("border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_DASH_DOT_DOT:
				css.append("border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_DASHED:
				css.append("border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_DOTTED:
				css.append("border-").append(border).append("-style:dotted;");
				break;
			case CellStyle.BORDER_HAIR:
				css.append("border-").append(border).append("-style:dotted;");
				break;
			case CellStyle.BORDER_DOUBLE:
				css.append("border-").append(border).append("-style:double;");
				break;
			case CellStyle.BORDER_MEDIUM:
				css.append("border-").append(border).append("-width:medium;");
				break;
			case CellStyle.BORDER_MEDIUM_DASH_DOT:
				css.append("border-").append(border).append("-width:medium;border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
				css.append("border-").append(border).append("-width:medium;border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_MEDIUM_DASHED:
				css.append("border-").append(border).append("-width:medium;border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_THIN:
				css.append("border-").append(border).append("-width:thin;");
				break;
			case CellStyle.BORDER_THICK:
				css.append("border-").append(border).append("-width:thick;");
				break;
			case CellStyle.BORDER_SLANTED_DASH_DOT:
				css.append("border-").append(border).append("-width:thick;border-").append(border).append("-style:dashed;");
				break;
			case CellStyle.BORDER_NONE:
				css.append("border-").append(border).append(":none;");
				break;
			default:
				break;
			}
			// 边框颜色
			if (borderColor != null) {			
				css.append("border-").append(border).append("-color:").append(convertToStardColor(borderColor)).append(";");
			}
		}		
		// 文字换行
		boolean borderBoolean = style.getWrapText();
		if (borderBoolean) {
			css.append("word-wrap:break-word;word-break:normal;");
		} else {
			css.append("white-space:nowrap;");
		}
		return css.toString();
	}
	
	// POI颜色转为标准色彩
	private static String convertToStardColor(Color color) {
		if (color == null) {
			return "";
		}
		StringBuffer sb = new StringBuffer();
		if (color instanceof HSSFColor) {
			HSSFColor hc = (HSSFColor)color;
			if (HSSFColor.AUTOMATIC.index == hc.getIndex()) {
				return "";
			}
			sb.append("#");
			for (int i = 0; i < hc.getTriplet().length; i++) {
				sb.append(String.format("%02x", hc.getTriplet()[i]));
			}
		} else if (color instanceof XSSFColor) {
			XSSFColor xc = (XSSFColor)color;
			sb.append("#");
			for (int i = 0; i < xc.getRGB().length; i++) {
				sb.append(String.format("%02x", xc.getRGB()[i]));
			}
	    }
		return sb.toString();
	}
	
	/**
	 * 将xls的列字符串转换为列号（如AT --> 45）
	 * 
	 * @param columnName 列名字符串
	 * 
	 * @return int 列号（zero-based）
	 */
	public static int translateName2ColumnIndex(String columnName) {
		if (StringUtils.isBlank(columnName)) {
			return -1;
		}
		columnName = columnName.toUpperCase();
		int value = 0;
		for (int i = 0; i < columnName.length(); i++) {
			int delta = columnName.charAt(i) - 64;
			value = value * 26 + delta;
		}
		return value-1;
	}
	
	/**
	 * 将xls的列号转换为列字符串（如26 --> AA）
	 * 
	 * @param index 列号（zero-based）
	 * 
	 * @return String 列名字符串
	 */
	public static String translateColumnIndex2Name(int index) {
		if (index < 0) {
			return null;
		}
		int quotient = index / 26;
        if (quotient > 0) {
            return translateColumnIndex2Name(quotient-1) + (char)((index % 26) + 65);
        } 
        else {
            return "" + (char)((index % 26) + 65);
        }
	}
	
}
