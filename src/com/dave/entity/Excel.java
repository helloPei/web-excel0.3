package com.dave.entity;

import java.util.Date;

import com.fasterxml.jackson.annotation.JsonFormat;
/**
 * Excel基本信息
 * @author davewpw
 *
 */
public class Excel {
	/**自增主键ID*/
	private Integer excelId;
	/**Excel自定义名称（Excel内容第1行,第1单元格）*/
	private String excelName;
	/**Excel自定义日期（Excel内容第4行,第2单元格）*/
	private String excelDate;
	/**Excel自定义周期（Excel内容第4行,第3单元格）*/
	private String week;
	/**Excel导入日期（创建日期）*/
	@JsonFormat(pattern = "yyyy-MM-dd HH:mm:ss", timezone = "GMT+8")
	private Date createDate;
	/**Occupancy Rate Max*/
	private String occupancyRate;
	/**Excel内容样式*/
	private String type;
	
	public Integer getExcelId() {
		return excelId;
	}
	public void setExcelId(Integer excelId) {
		this.excelId = excelId;
	}
	public String getType() {
		return type;
	}
	public void setType(String type) {
		this.type = type;
	}
	public String getOccupancyRate() {
		return occupancyRate;
	}
	public void setOccupancyRate(String occupancyRate) {
		this.occupancyRate = occupancyRate;
	}
	public String getExcelName() {
		return excelName;
	}
	public void setExcelName(String excelName) {
		this.excelName = excelName;
	}
	public Date getCreateDate() {
		return createDate;
	}
	public void setCreateDate(Date createDate) {
		this.createDate = createDate;
	}
	public String getExcelDate() {
		return excelDate;
	}
	public void setExcelDate(String excelDate) {
		this.excelDate = excelDate;
	}
	public String getWeek() {
		return week;
	}
	public void setWeek(String week) {
		this.week = week;
	}
}