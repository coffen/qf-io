package com.qf.io;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 文件读取错误
 * <br>
 * File Name: FileErrorException.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月27日 上午9:41:37 
 * @version: v1.0
 *
 */
public class FileErrorException extends Exception {
	
	private static final long serialVersionUID = -7729683727415612039L;
	
	public static final int FILE_NOT_FOUND = 1 << 1;
	public static final int FILE_EXTENSION_ERROR = 1 << 2;
	
	private String suffix;  // 文件后缀或文件类型
	private int errType;  	// 错误类型
	
	public FileErrorException(String suffix, int error) {
		this.suffix = suffix;
		this.errType = error;
	}
	
	public FileErrorException(int error) {
		this.errType = error;
	}
	
	public String getSuffix() {
		return suffix;
	}
	
	public int getErrType() {
		return errType;
	}
	
	public String getErrMsg() {
		switch (errType) {
			case FILE_NOT_FOUND:
				return "文件不存在";
			case FILE_EXTENSION_ERROR:
				return "文件扩展名不符合规范";
			default:
				return "未知错误";
		}
	}

}
