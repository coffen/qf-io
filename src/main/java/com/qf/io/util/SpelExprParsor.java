package com.qf.io.util;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.joda.time.DateTime;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
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
	
	private final static Logger log = LoggerFactory.getLogger(SpelExprParsor.class);
	
	private final SpelExpressionParser parser = new SpelExpressionParser();				// Spel解析器
	private final StandardEvaluationContext context = new StandardEvaluationContext();	// Spel上下文
	
	private final static String ROOT_PREFIX = "_root";
	
	/**
	 * <p>设置根对象到Spel上下文中</p>
	 * 
	 * @param obj 根对象
	 */
	public void setRootVariable(Object obj) {
		context.setVariable(ROOT_PREFIX, obj);
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
		String exprRege = "[#\\$]\\{\\w+(\\[\\?\\])?(\\.\\w+(\\[\\?\\])?)*}";
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
	 * 获取静态表达式替换数据后的对象（目前只返回两类对象：byte[]和String, 分别对应图片和文本内容）
	 * 
	 * @param expression 包含EL表达式的字符串, 可能包含多个EL表达式, 例: 计划名称:${plan.planName}
	 * 
	 * @return String 替换后的字符串
	 */	
	public Object getValue(String expression) {
		Object result = null;
		if (StringUtils.isBlank(expression)) {
			return result;
		}
		String exprRege = "\\$\\{\\w+(\\[(0|[1-9][0-9]*)\\])?(\\.\\w+(\\[(0|[1-9][0-9]*)\\])?)*}";
		Pattern pattern = Pattern.compile(exprRege);
		Matcher matcher = pattern.matcher(expression);
		String prefix = buildPrefix();
		
		StringBuffer buffer = new StringBuffer();
		while (matcher.find()) {
			String group = matcher.group(0);
			String queryName = group.replaceAll("\\$\\{|}", "");
			String fragment = null;
			try {
				Object obj = parser.parseExpression(prefix + queryName).getValue(context);
				if (obj instanceof byte[]) { // byte[]不予处理, 直接返回对象
					return obj;
				}
				if (obj instanceof Number) {
					if (group.equals(expression)) {
						return obj;
					}
					fragment = obj.toString();
				} 
				if (obj instanceof BigDecimal) {
					if (group.equals(expression)) {
						return obj;
					}
					fragment = ((BigDecimal) obj).toEngineeringString();
				} 
				else if (obj instanceof String) {
					fragment = (String)obj;
				} 
				else if (obj instanceof Date) {
					fragment = new DateTime((Date)obj).toString("yyyy-MM-dd HH:mm:ss");
				} 
				else if (obj instanceof Boolean) {
					fragment = obj.toString();
				}
				else {
					fragment = obj.toString();
				}
			} 
			catch (Exception e) {
				log.error("表达式解析错误: expr=" + group);
			}
			matcher.appendReplacement(buffer, fragment != null ? fragment : "");
		}
		matcher.appendTail(buffer);		
		result = buffer.toString();
		return result;
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
		String expr = "\\$\\{\\w+(\\[\\?\\])?(\\.\\w+(\\[\\?\\])?)*}";
		Pattern pattern = Pattern.compile(expr);
		Matcher matcher = pattern.matcher(expression);
		String prefix = buildPrefix();
		
		Object target = null;
		while (matcher.find()) {
			String group = matcher.group(0);
			String queryName = group.replaceAll("\\$\\{|}", "");
			int idx = queryName.indexOf("[?]");
			if (idx <= 0) {
				return 0;
			}
			queryName = queryName.substring(0, idx);	// 截取数组变量名称
			try {
				target = parser.parseExpression(prefix + queryName).getValue(context);
			} 
			catch (Exception e) {
				log.error("表达式解析错误: expr=" + group, e);
				return 0;
			}
			break;
		}
		if (target == null || !(target instanceof List)) {
			return 0;
		}
		return ((List<?>)target).size();
	}
	
	private String buildPrefix() {
		return "#" + ROOT_PREFIX + ".";
	}
	
}
