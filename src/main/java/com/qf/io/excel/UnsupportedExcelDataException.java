package com.qf.io.excel;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 未支持的Excel导出数据规格
 * <br>
 * File Name: UnsupportedExcelDataException.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:18:02 
 * @version: v1.0
 *
 */
public class UnsupportedExcelDataException extends Exception {

	private static final long serialVersionUID = 7333553524549707383L;
	
	public static final int ILLEGAL_PARAMS = 1;
	
	public static final int EXCEED_ROW_COUNT_LIMIT_ERROR = 1 << 1;
	public static final int EXCEED_COLUMN_COUNT_LIMIT_ERROR = 1 << 2;
	
	public static final int UNRECOGNIZED_CELL_DATA_ERROR = 1 << 3;
	
	private int errType;
	
	public UnsupportedExcelDataException(int err) {
		this.errType = err;
	}
	
	public int getErrType() {
		return errType;
	}
	
	public String getErrMsg() {
		switch (errType) {
			case EXCEED_ROW_COUNT_LIMIT_ERROR:
				return "超出最大行限制";
			case EXCEED_COLUMN_COUNT_LIMIT_ERROR:
				return "超出最大列限制";
			case UNRECOGNIZED_CELL_DATA_ERROR:
				return "超出最大行限制";
			default:
				return "未知错误";
		}
	}

}
