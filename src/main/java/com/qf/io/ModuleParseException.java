package com.qf.io;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 模板解析错误
 * <br>
 * File Name: ModuleParseException.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017-08-28 20:19:19 
 * @version: v1.0
 *
 */
public class ModuleParseException extends Exception {

	private static final long serialVersionUID = -5619356565318189680L;
	
	private Exception exception;
	private String msg;
	
	public ModuleParseException(String msg) {
		this.msg = msg;
	}
	
	public ModuleParseException(Exception exception, String msg) {
		this.exception = exception;
		this.msg = msg;
	}
	
	public String getMsg() {
		return msg;
	}
	
	@Override
	public String getMessage() {
		String innerError = exception == null ? "" : exception.getMessage();
		return super.getMessage() + innerError;
	}

}
