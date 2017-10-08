package com.qf.io.test.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.joda.time.DateTime;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.qf.io.ModuleParseException;
import com.qf.io.excel.writer.module.PoiELModule;

/**
 * 
 * <p>
 * Project Name: C2C商城
 * <br>
 * Description: PoiEl模板导出测试用例
 * <br>
 * File Name: PoiELExporterTest.java
 * <br>
 * Copyright: Copyright (C) 2015 All Rights Reserved.
 * <br>
 * Company: 杭州偶尔科技有限公司
 * <br>
 * @author 穷奇
 * @create time：2017年8月3日 下午3:20:45 
 * @version: v1.0
 *
 */
public class PoiELExporterTest {
	
	private final static Logger log = LoggerFactory.getLogger(PoiELExporterTest.class);
	
	private final String path = Thread.currentThread().getContextClassLoader().getResource("").getPath();
	
	@Test
	public void export() throws IOException, ModuleParseException {
		String filename = "elModule.xlsx";		
		PoiELModule module = new PoiELModule(path + filename);
		
		String exportFile = "elModuleExportFile.xlsx";
		OutputStream os = new FileOutputStream(new File(path + exportFile));
		
		ExporterBean bean = new ExporterBean();
		SmsPlan smsPlanBean = new SmsPlan();
		smsPlanBean.setSubject("测试标题");
		smsPlanBean.setMatchMemberCount(1256);
		smsPlanBean.setSendCount(1198);
		smsPlanBean.setSendMemberCount(1160);
		smsPlanBean.setFailMemberCount(23);
		
		bean.setSmsPlanBean(smsPlanBean);
		
		bean.setStatBeginDate(new DateTime(2017, 8, 1, 0, 0).toDate());
		bean.setStatEndDate(new DateTime(2017, 8, 2, 0, 0).toDate());
		
		bean.setRoi(1.1);
		
		bean.setTotalRepeatMemberCount(103);
		bean.setTotalRepeatOrderCount(115);
		bean.setTotalRepeatSalesAmount(new BigDecimal("33598.0"));
		bean.setPayedRepeatMemberCount(85);
		bean.setPayedRepeatOrderCount(91);
		bean.setPayedRepeatSalesAmount(new BigDecimal("29877.0"));
		bean.setRepeatRatio(32.3);
		
		bean.setTotalShopMemberCount(25);
		bean.setTotalShopSalesAmount(new BigDecimal("1250.0"));
		bean.setTotalShopOrderCount(28);
		
		bean.setTotalFirstGradeRepeatMemberCount(250);
		bean.setTotalSecondGradeRepeatMemberCount(45);
		bean.setTotalThirdGradeRepeatMemberCount(12);
		bean.setTotalFourthGradeRepeatMemberCount(1);
		bean.setTotalFifthGradeRepeatMemberCount(0);
		
		bean.setTotalSmsCost(new BigDecimal("50.24"));
		
		for (int i = 1; i <= 3; i++) {
			File f = new File(path + "img" + i + ".png");
			FileInputStream fs = new FileInputStream(f);
			byte[] b = new byte[(int)f.length()];
			fs.read(b);
			fs.close();
			if (i == 1) {
				bean.setImg1(b);
			}
			else if (i == 2) {
				bean.setImg2(b);
			}
			else {
				bean.setImg3(b);
			}
		}
		
		List<PlanStat> statList = new ArrayList<PlanStat>();
		bean.setStatListForPlan(statList);
		
		for (int i = 0; i < Math.random() * 20; i++) {
			PlanStat stat = new PlanStat();
			
			stat.setStatDate(new DateTime().toDate());
			
			stat.setRepeatMemberCount((int)(Math.random() * 120));
			stat.setPayedMemberCount((int)(Math.random() * 90));
			stat.setRepeatOrderCount((int)(Math.random() * 140));
			stat.setPayedOrderCount((int)(Math.random() * 80));
			
			stat.setRepeatSalesAmount(new BigDecimal((int)(Math.random() * 5200)));
			stat.setPayedSalesAmount(new BigDecimal((int)(Math.random() * 4500)));
			
			statList.add(stat);
		}
		
		List<PlanStat> repeatList = new ArrayList<PlanStat>();
		bean.setRepeatListForPlan(repeatList);
		
		for (int i = 0; i < Math.random() * 20; i++) {
			PlanStat stat = new PlanStat();
			
			stat.setStatDate(new DateTime().toDate());
			
			stat.setRepeatMemberCount((int)(Math.random() * 120));
			stat.setPayedMemberCount((int)(Math.random() * 90));
			stat.setRepeatOrderCount((int)(Math.random() * 140));
			stat.setPayedOrderCount((int)(Math.random() * 80));
			
			stat.setRepeatSalesAmount(new BigDecimal((int)(Math.random() * 5200)));
			stat.setPayedSalesAmount(new BigDecimal((int)(Math.random() * 4500)));
			
			repeatList.add(stat);
		}
		
		log.error("导出开始...");
		
		module.export(os, bean);
		
		log.error("导出完成.");
	}
	
	class ExporterBean implements Serializable {

		private static final long serialVersionUID = -6438069775426804791L;
		
		private SmsPlan smsPlanBean;
		
		private Date statBeginDate;
		private Date StatEndDate;
		
		private double roi;
		
		private int totalRepeatMemberCount;
		private int payedRepeatMemberCount;
		private int totalRepeatOrderCount;
		private int payedRepeatOrderCount;
		private BigDecimal totalRepeatSalesAmount;
		private BigDecimal payedRepeatSalesAmount;
		
		private double repeatRatio;
		
		private int totalShopMemberCount;
		private BigDecimal totalShopSalesAmount;
		private int totalShopOrderCount;
		
		private int totalFirstGradeRepeatMemberCount;
		private int totalSecondGradeRepeatMemberCount;
		private int totalThirdGradeRepeatMemberCount;
		private int totalFourthGradeRepeatMemberCount;
		private int totalFifthGradeRepeatMemberCount;
		
		private BigDecimal totalSmsCost;
		
		private byte[] img1;
		private byte[] img2;
		private byte[] img3;
		
		private List<PlanStat> statListForPlan;
		private List<PlanStat> repeatListForPlan;
		
		public SmsPlan getSmsPlanBean() {
			return smsPlanBean;
		}
		
		public void setSmsPlanBean(SmsPlan smsPlanBean) {
			this.smsPlanBean = smsPlanBean;
		}
		
		public Date getStatBeginDate() {
			return statBeginDate;
		}
		
		public void setStatBeginDate(Date statBeginDate) {
			this.statBeginDate = statBeginDate;
		}
		
		public Date getStatEndDate() {
			return StatEndDate;
		}
		
		public void setStatEndDate(Date statEndDate) {
			StatEndDate = statEndDate;
		}
		
		public double getRoi() {
			return roi;
		}
		
		public void setRoi(double roi) {
			this.roi = roi;
		}

		public int getTotalRepeatMemberCount() {
			return totalRepeatMemberCount;
		}

		public void setTotalRepeatMemberCount(int totalRepeatMemberCount) {
			this.totalRepeatMemberCount = totalRepeatMemberCount;
		}

		public int getPayedRepeatMemberCount() {
			return payedRepeatMemberCount;
		}

		public void setPayedRepeatMemberCount(int payedRepeatMemberCount) {
			this.payedRepeatMemberCount = payedRepeatMemberCount;
		}

		public int getTotalRepeatOrderCount() {
			return totalRepeatOrderCount;
		}

		public void setTotalRepeatOrderCount(int totalRepeatOrderCount) {
			this.totalRepeatOrderCount = totalRepeatOrderCount;
		}

		public int getPayedRepeatOrderCount() {
			return payedRepeatOrderCount;
		}

		public void setPayedRepeatOrderCount(int payedRepeatOrderCount) {
			this.payedRepeatOrderCount = payedRepeatOrderCount;
		}

		public BigDecimal getTotalRepeatSalesAmount() {
			return totalRepeatSalesAmount;
		}

		public void setTotalRepeatSalesAmount(BigDecimal totalRepeatSalesAmount) {
			this.totalRepeatSalesAmount = totalRepeatSalesAmount;
		}

		public BigDecimal getPayedRepeatSalesAmount() {
			return payedRepeatSalesAmount;
		}

		public void setPayedRepeatSalesAmount(BigDecimal payedRepeatSalesAmount) {
			this.payedRepeatSalesAmount = payedRepeatSalesAmount;
		}
		
		public double getRepeatRatio() {
			return repeatRatio;
		}
		
		public void setRepeatRatio(double repeatRatio) {
			this.repeatRatio = repeatRatio;
		}

		public int getTotalShopMemberCount() {
			return totalShopMemberCount;
		}

		public void setTotalShopMemberCount(int totalShopMemberCount) {
			this.totalShopMemberCount = totalShopMemberCount;
		}

		public BigDecimal getTotalShopSalesAmount() {
			return totalShopSalesAmount;
		}

		public void setTotalShopSalesAmount(BigDecimal totalShopSalesAmount) {
			this.totalShopSalesAmount = totalShopSalesAmount;
		}

		public int getTotalShopOrderCount() {
			return totalShopOrderCount;
		}

		public void setTotalShopOrderCount(int totalShopOrderCount) {
			this.totalShopOrderCount = totalShopOrderCount;
		}

		public int getTotalFirstGradeRepeatMemberCount() {
			return totalFirstGradeRepeatMemberCount;
		}

		public void setTotalFirstGradeRepeatMemberCount(int totalFirstGradeRepeatMemberCount) {
			this.totalFirstGradeRepeatMemberCount = totalFirstGradeRepeatMemberCount;
		}

		public int getTotalSecondGradeRepeatMemberCount() {
			return totalSecondGradeRepeatMemberCount;
		}

		public void setTotalSecondGradeRepeatMemberCount(int totalSecondGradeRepeatMemberCount) {
			this.totalSecondGradeRepeatMemberCount = totalSecondGradeRepeatMemberCount;
		}

		public int getTotalThirdGradeRepeatMemberCount() {
			return totalThirdGradeRepeatMemberCount;
		}

		public void setTotalThirdGradeRepeatMemberCount(int totalThirdGradeRepeatMemberCount) {
			this.totalThirdGradeRepeatMemberCount = totalThirdGradeRepeatMemberCount;
		}

		public int getTotalFourthGradeRepeatMemberCount() {
			return totalFourthGradeRepeatMemberCount;
		}

		public void setTotalFourthGradeRepeatMemberCount(int totalFourthGradeRepeatMemberCount) {
			this.totalFourthGradeRepeatMemberCount = totalFourthGradeRepeatMemberCount;
		}

		public int getTotalFifthGradeRepeatMemberCount() {
			return totalFifthGradeRepeatMemberCount;
		}

		public void setTotalFifthGradeRepeatMemberCount(int totalFifthGradeRepeatMemberCount) {
			this.totalFifthGradeRepeatMemberCount = totalFifthGradeRepeatMemberCount;
		}
		
		public BigDecimal getTotalSmsCost() {
			return totalSmsCost;
		}
		
		public void setTotalSmsCost(BigDecimal totalSmsCost) {
			this.totalSmsCost = totalSmsCost;
		}
		
		public byte[] getImg1() {
			return img1;
		}
		
		public void setImg1(byte[] img1) {
			this.img1 = img1;
		}
		
		public byte[] getImg2() {
			return img2;
		}
		
		public void setImg2(byte[] img2) {
			this.img2 = img2;
		}
		
		public byte[] getImg3() {
			return img3;
		}
		
		public void setImg3(byte[] img3) {
			this.img3 = img3;
		}
		
		public List<PlanStat> getStatListForPlan() {
			return statListForPlan;
		}
		
		public void setStatListForPlan(List<PlanStat> statListForPlan) {
			this.statListForPlan = statListForPlan;
		}
		
		public List<PlanStat> getRepeatListForPlan() {
			return repeatListForPlan;
		}
		
		public void setRepeatListForPlan(List<PlanStat> repeatListForPlan) {
			this.repeatListForPlan = repeatListForPlan;
		}
		
	}
	
	class SmsPlan {
		
		private String subject;
		
		private int matchMemberCount;
		private int sendMemberCount;
		private int failMemberCount;
		private int sendCount;
		
		public String getSubject() {
			return subject;
		}
		
		public void setSubject(String subject) {
			this.subject = subject;
		}

		public int getMatchMemberCount() {
			return matchMemberCount;
		}

		public void setMatchMemberCount(int matchMemberCount) {
			this.matchMemberCount = matchMemberCount;
		}

		public int getSendMemberCount() {
			return sendMemberCount;
		}

		public void setSendMemberCount(int sendMemberCount) {
			this.sendMemberCount = sendMemberCount;
		}

		public int getFailMemberCount() {
			return failMemberCount;
		}

		public void setFailMemberCount(int failMemberCount) {
			this.failMemberCount = failMemberCount;
		}

		public int getSendCount() {
			return sendCount;
		}

		public void setSendCount(int sendCount) {
			this.sendCount = sendCount;
		}
		
	}
	
	class PlanStat {
		
		private Date statDate;
		
		private int repeatMemberCount;
		private int payedMemberCount;
		private int repeatOrderCount;
		private int payedOrderCount;
		private BigDecimal repeatSalesAmount;
		private BigDecimal payedSalesAmount;
		
		public Date getStatDate() {
			return statDate;
		}
		
		public void setStatDate(Date statDate) {
			this.statDate = statDate;
		}
		
		public int getRepeatMemberCount() {
			return repeatMemberCount;
		}
		
		public void setRepeatMemberCount(int repeatMemberCount) {
			this.repeatMemberCount = repeatMemberCount;
		}
		
		public int getPayedMemberCount() {
			return payedMemberCount;
		}
		
		public void setPayedMemberCount(int payedMemberCount) {
			this.payedMemberCount = payedMemberCount;
		}
		
		public int getRepeatOrderCount() {
			return repeatOrderCount;
		}
		
		public void setRepeatOrderCount(int repeatOrderCount) {
			this.repeatOrderCount = repeatOrderCount;
		}
		
		public int getPayedOrderCount() {
			return payedOrderCount;
		}
		
		public void setPayedOrderCount(int payedOrderCount) {
			this.payedOrderCount = payedOrderCount;
		}
		
		public BigDecimal getRepeatSalesAmount() {
			return repeatSalesAmount;
		}
		
		public void setRepeatSalesAmount(BigDecimal repeatSalesAmount) {
			this.repeatSalesAmount = repeatSalesAmount;
		}
		
		public BigDecimal getPayedSalesAmount() {
			return payedSalesAmount;
		}
		
		public void setPayedSalesAmount(BigDecimal payedSalesAmount) {
			this.payedSalesAmount = payedSalesAmount;
		}		
	}

}
