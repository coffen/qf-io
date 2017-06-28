package com.qf.io.util;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.joda.time.DateTime;
import org.springframework.expression.spel.standard.SpelExpressionParser;
import org.springframework.expression.spel.support.StandardEvaluationContext;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: Spel表达式解析工具
 * <br>
 * File Name: SpelExprParsor.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年6月28日 上午9:59:31 
 * @version: v1.0
 *
 */
public class SpelExprParsor {
	
	// Spel解析器
	SpelExpressionParser parser = new SpelExpressionParser();
	// SPEL上下文
	StandardEvaluationContext context = new StandardEvaluationContext();
	
	/**
	 * 设置根对象到Spel上下文中
	 * <br>
	 * 根对象设为固定值root，意味着模板中的根对象名都应该是root
	 * 
	 * @param obj 根对象
	 */
	public void setRootVariable(Object obj) {
		context.setVariable("rootObj", obj);
	}
	
	/**
	 * 获取源字符串的EL表达式, 如不符合条件则返回Null
	 * <p>
	 * 1) isRestrict = true; 则源字符串不包含其他字符且只由一个EL表达式组成, 否则返回Null
	 * 2) isRestrict = false; 如果存在多个EL表达式, 返回多个EL表达式组成的数组
	 * </p>
	 * 
	 * @param expression 待处理字符串
	 * @param isRestrict 是否严格匹配
	 * 
	 * @return String 数组
	 */
	public String[] getElExpressions(String expression, boolean isRestrict) {
		if (StringUtils.isBlank(expression)) {
			return null;
		}
		String exprRege = "[#\\$]\\{\\w+(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*(\\.\\w+\\[\\?\\])?(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*\\}";
		if (isRestrict) {
			return Pattern.matches(exprRege, expression) ? new String[] { expression.replaceAll("[#\\$]\\{|}", "") } : null;
		}
		else {
			Pattern pattern = Pattern.compile(exprRege);
			Matcher matcher = pattern.matcher(expression);
			List<String> elExprList = new ArrayList<String>();
			String group = null;
			String queryName = null;
			while (matcher.find()) {
				group = matcher.group(0);
				queryName = group.replaceAll("[#\\$]\\{|}", "");
				elExprList.add(queryName);
			}
			if (elExprList.size() > 0) {
				return elExprList.toArray(new String[0]);
			}
		}
		return null;
	}
	
	/**
	 * 获取表达式替换数据后的对象（目前只返回两类对象：byte[]和String, 分别对应图片和文本内容）
	 * 
	 * @param expression 包含EL表达式的字符串, 可能包含多个EL表达式, 例: 计划名称:${plan.planName}
	 * 
	 * @return String 替换后的字符串
	 */	
	public Object getValue(String expression) {
		if (StringUtils.isBlank(expression)) {
			return "";
		}
		String exprRege = "\\$\\{\\w+(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*\\}";
		Pattern pattern = Pattern.compile(exprRege);
		Matcher matcher = pattern.matcher(expression);
		while (matcher.find()) {
			String group = matcher.group(0);
			String queryName = group.replaceAll("\\$\\{|}", "");
			String fragment = null;
			try {
				Object obj = parser.parseExpression("#" + queryName).getValue(context);
				// 如果传入参数完全匹配EL表达式的规则, 直接返回对象
				if (expression.equalsIgnoreCase(group)) {
					return obj;
				}
				if (obj instanceof Number) {
					fragment = obj.toString();
				} else if (obj instanceof String) {
					fragment = (String)obj;
				} else if (obj instanceof Date) {
					fragment = new DateTime((Date)obj).toString("yyyy-MM-dd");
				} else if (obj instanceof Boolean) {
					fragment = obj.toString();
				} else if (obj instanceof byte[]) { // byte[]不予处理
					return obj;
				} else {
					fragment = obj.toString();
				}
			} catch (Exception e) {
				return null;
			}
			expression = expression.replace(group, fragment != null ? fragment : "");
		}
		return expression;
	}
	
	/**
	 * 获取目标表达式指定的列表对象长度
	 * 
	 * @param expression 表达式, 例: ${plan.planDetailList[?].smsContent}
	 * 
	 * @return Object 返回的list长度
	 */	
	public int getListObjectSize(String expression) {
		if (StringUtils.isBlank(expression)) {
			return 0;
		}
		String expr = "\\$\\{\\w+(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*\\.\\w+\\[\\?\\](\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*\\}";
		Pattern pattern = Pattern.compile(expr);
		Matcher matcher = pattern.matcher(expression);
		Object target = null;
		while (matcher.find()) {
			String group = matcher.group(0);
			String queryName = group.replaceAll("\\$\\{|}", "");
			int idx = queryName.indexOf("[?]");
			if (idx <= 0) {
				return 0;
			}
			queryName = queryName.substring(0, idx);
			try {
				target = parser.parseExpression("#" + queryName).getValue(context);
			} catch (Exception e) {
				return 0;
			}
			break;
		}
		if (target == null || !(target instanceof List)) {
			return 0;
		}
		return ((List<?>)target).size();
	}
	
}
