package com.dave.entity;
/**
 * Excel内容
 * @author davewpw
 *
 */
public class ExcelAll {
	/**自增主键ID*/
	private Integer excelAllId;
	/**ExcelID*/
	private Integer excelId;
	/**--Time*/
	private String time;
	/**--Incoming Call Answer*/
	private String inCallAnswer;
	/**--Incoming Average Holding Time Per Call (sec)*/
	private String inAverageHoldingTPC;
	/**--Outgoing Call Answer*/
	private String outCallAnswer;
	/**--Outgoing Average Holding Time Per Call (sec)*/
	private String outAverageHoldingTPC;
	/**--Service Capacity*/
	private String serviceCapacity;
	/**--Capacity Needed*/
	private String capacityNeeded;
	/**Incoming Call Answer * Incoming Average Holding Time Per Call (sec)
	 * --Incoming total seconds in the hour*/
	private String inTotalHour;
	/**Outgoing Call Answer * Outgoing Average Holding Time Per Call (sec)
	 * --Outgoing total seconds in the hour*/
	private String outTotalHour;
	/**--Occupancy Hour(hour)*/
	private String occupancyHour;
	/**--Occupancy Rate(%)*/
	private String occupancyRate;
	
	public Integer getExcelAllId() {
		return excelAllId;
	}
	public void setExcelAllId(Integer excelAllId) {
		this.excelAllId = excelAllId;
	}
	public Integer getExcelId() {
		return excelId;
	}
	public void setExcelId(Integer excelId) {
		this.excelId = excelId;
	}
	public String getOccupancyHour() {
		return occupancyHour;
	}
	public void setOccupancyHour(String occupancyHour) {
		this.occupancyHour = occupancyHour;
	}
	public String getOccupancyRate() {
		return occupancyRate;
	}
	public void setOccupancyRate(String occupancyRate) {
		this.occupancyRate = occupancyRate;
	}
	public String getTime() {
		return time;
	}
	public void setTime(String time) {
		this.time = time;
	}
	public String getInCallAnswer() {
		return inCallAnswer;
	}
	public void setInCallAnswer(String inCallAnswer) {
		this.inCallAnswer = inCallAnswer;
	}
	public String getInAverageHoldingTPC() {
		return inAverageHoldingTPC;
	}
	public void setInAverageHoldingTPC(String inAverageHoldingTPC) {
		this.inAverageHoldingTPC = inAverageHoldingTPC;
	}
	public String getOutCallAnswer() {
		return outCallAnswer;
	}
	public void setOutCallAnswer(String outCallAnswer) {
		this.outCallAnswer = outCallAnswer;
	}
	public String getOutAverageHoldingTPC() {
		return outAverageHoldingTPC;
	}
	public void setOutAverageHoldingTPC(String outAverageHoldingTPC) {
		this.outAverageHoldingTPC = outAverageHoldingTPC;
	}
	public String getServiceCapacity() {
		return serviceCapacity;
	}
	public void setServiceCapacity(String serviceCapacity) {
		this.serviceCapacity = serviceCapacity;
	}
	public String getCapacityNeeded() {
		return capacityNeeded;
	}
	public void setCapacityNeeded(String capacityNeeded) {
		this.capacityNeeded = capacityNeeded;
	}
	public String getInTotalHour() {
		return inTotalHour;
	}
	public void setInTotalHour(String inTotalHour) {
		this.inTotalHour = inTotalHour;
	}
	public String getOutTotalHour() {
		return outTotalHour;
	}
	public void setOutTotalHour(String outTotalHour) {
		this.outTotalHour = outTotalHour;
	}
}