package com.qf.io.excel.writer.module;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: 单元格坐标
 * <br>
 * File Name: Point.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017-10-09 20:45:15 
 * @version: v1.0
 *
 */
public class Point {
	
	private int x;
	private int y;
	
	public Point(int x, int y) {
		this.x = x;
		this.y = y;
	}

	int getX() {
		return x;
	}
	
	public int getY() {
		return y;
	}
	
	@Override
	public boolean equals(Object obj) {
		if (obj == null || !(obj instanceof Point)) {
			return false;
		}
		Point p = (Point)obj;
		return p.getX() == x && p.getY() == y;
	}
	
	@Override
	public int hashCode() {
		return (x + "," + y).hashCode();
	}

}
