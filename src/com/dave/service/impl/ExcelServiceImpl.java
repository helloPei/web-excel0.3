package com.dave.service.impl;

import com.dave.common.util.ExcelUtil;
import com.dave.common.vo.JsonResult;
import com.dave.common.vo.PageObject;
import com.dave.dao.ExcelDao;
import com.dave.entity.Excel;
import com.dave.entity.ExcelAll;
import com.dave.entity.ExcelTotal;
import com.dave.service.ExcelService;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Excel业务层接口实现类
 * @author davewpw
 *
 */
@Service
public class ExcelServiceImpl implements ExcelService {
	/**Excel数据持久层接口*/
	@Autowired
	private ExcelDao excelDao;
	/**根据Excel ID 查询Excel内容*/
	@Override
	public List<ExcelAll> selectExcelAll(int excelId) {
		return excelDao.selectExcelAll(excelId);
	}
	/**根据Excel ID 查询Excel*/
	@Override
	public Excel selectExcelById(int excelId){
		return excelDao.selectExcelById(excelId);
	}
	/**查询框功能查询所有 Excel,以及页面初始数据查询*/
	@Override
	public PageObject<Excel> searchExcel(String excelDate, String excelName, int isSearchMax, int pageCurrent){
	    //计算startIndex的值
        int pageSize = 20;
        //依据条件获取当前页数据
        int startIndex = (pageCurrent-1) * pageSize;
        int rowCount = excelDao.selectCountExcel(excelDate, excelName);
        //重置当前页面参数（总数小于开始查询下标或isSearchMax = 1）
        if(isSearchMax == 1 || (rowCount < startIndex)){
            pageCurrent = 1;
            startIndex = (pageCurrent-1) * pageSize;
        }
        List<Excel> records = null;
        if(isSearchMax == 1){
            records = excelDao.searchExcelMax(excelName);
            //isSearchMax = 1重置总数量
            rowCount = 1;
        }else{
            records = excelDao.searchExcel(excelDate, excelName, startIndex, pageSize*pageCurrent);
        }
        if(StringUtils.isEmpty(records)) {
            return null;
        }
        //设置分页对象参数
        PageObject<Excel> pageObject = new PageObject<>();
        pageObject.setPageCurrent(pageCurrent);
        pageObject.setRowCount(rowCount);
        pageObject.setPageSize(pageSize);
        pageObject.setRecords(records);
        return pageObject;
    }
	/**导入Excel*/
	@Override
	public JsonResult batchImport(String fileName, MultipartFile file)throws Exception {
		if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$") 
				&& !fileName.matches("^.+\\.(?i)(csv)$"))return new JsonResult("上传文件格式不正确");
		InputStream is = file.getInputStream();
		List<String[]> list = null;
		if (fileName.matches("^.+\\.(?i)(csv)$"))list = ExcelUtil.readCsvByIs(is);
		if (fileName.matches("^.+\\.(?i)(xlsx)$") || fileName.matches("^.+\\.(?i)(xls)$"))
			list = ExcelUtil.readXlsAndXlsxByIs(is);
		if(list == null)return new JsonResult("上传文件为空！");
		Excel excel = new Excel();
		String[] row1 = list.get(0);
		excel.setExcelName(row1[0]);
		String[] row2 = list.get(1);
		String excelDate = row2[1];
		if(excelDate.indexOf("/") != 2)return new JsonResult("上传文件格式错误！");
		String[] splitAddress = excelDate.split("/");
		excelDate = splitAddress[2] + "-" + splitAddress[1] + "-" + splitAddress[0];
		excel.setExcelDate(excelDate);
		excel.setWeek(row2[2]);
		excel.setCreateDate(new Date());
		String[] row3 = list.get(2);
		if(row3.length == 13 || row3.length >= 22){
			excelDao.addExcel(excel);
			int excelId = excelDao.selectExcelByName(excel.getExcelName());
			excel.setExcelId(excelId);
			String excelType = null;
			if(row3.length == 13){
				excelType = "Outgoing Only";
				excel.setType("Outgoing only");
			}else{
				excelType = "Incoming And Outgoing";
				excel.setType("IN & OUT");
			}
			if(excelType.equals("Outgoing Only")){
				//Outgoing Only Excel
				for(int i = 3; i < 27; i++){
					String[] row = list.get(i);
					if (row == null)continue;
					ExcelAll excelAll = new ExcelAll();
					excelAll.setExcelId(excelId);				//ExcelId
					excelAll.setTime(row[0]);					//Time
					excelAll.setOutCallAnswer(row[5]);			//OutCallAnswer
					excelAll.setOutAverageHoldingTPC(row[9]);	//OutAverageHoldingTPC
					excelAll.setServiceCapacity(row[10]);		//ServiceCapacity
					excelAll.setCapacityNeeded(row[11]);		//CapacityNeeded
					//计算 Outgoing total seconds in the hour
					BigDecimal outCallAnswer = new BigDecimal(row[5]);
					BigDecimal outAverageHoldingTPC = new BigDecimal(row[9]);
					//Outgoing Call Answer * Outgoing Average Holding Time Per Call (sec)
					BigDecimal outTotalHour = outCallAnswer.multiply(outAverageHoldingTPC).setScale(2, BigDecimal.ROUND_HALF_UP);
					excelAll.setOutTotalHour(outTotalHour.toString());//OutTotalHour
					//计算 Occupancy in hours
					//Outgoing Occupancy hour = Outgoing total hour / 3600
					BigDecimal outOccHour = outTotalHour.divide(new BigDecimal("3600"), 4, RoundingMode.HALF_UP);
					BigDecimal outOccHour2 = outOccHour.setScale(2, RoundingMode.HALF_UP);
					excelAll.setOccupancyHour(outOccHour2.toString());//OccupancyHour
					//计算 Occupancy Rate
					String serviceCapacity = row[10];
					//Outgoing Occupancy Rate(%) = Outgoing Occupancy hour / T1 Capacity(Service Capacity) * 100
					BigDecimal outOccRate = (outOccHour.divide(new BigDecimal(serviceCapacity), 4, RoundingMode.HALF_UP))
							.multiply(new BigDecimal("100")).setScale(2, BigDecimal.ROUND_HALF_UP);
					excelAll.setOccupancyRate(outOccRate.toString());//OccupancyRate
					excelDao.addExcelAll(excelAll);
				}	
			}else{
				//Incoming And Outgoing Excel
				for(int i = 3; i < 27; i++){
					String[] row = list.get(i);
					if (row == null)continue;
					ExcelAll excelAll = new ExcelAll();
					excelAll.setExcelId(excelId);				//ExcelId
					excelAll.setTime(row[0]);					//Time	
					excelAll.setInCallAnswer(row[5]);			//InCallAnswer
					excelAll.setInAverageHoldingTPC(row[9]);	//InAverageHoldingTPC
					excelAll.setOutCallAnswer(row[14]);			//OutCallAnswer
					excelAll.setOutAverageHoldingTPC(row[18]);	//OutAverageHoldingTPC
					excelAll.setServiceCapacity(row[19]);		//ServiceCapacity
					excelAll.setCapacityNeeded(row[20]);		//CapacityNeeded
					//计算 Incoming total seconds in the hour
					BigDecimal inCallAnswer = new BigDecimal(row[5]);
					BigDecimal inAverageHoldingTPC = new BigDecimal(row[9]);
					//Incoming Call Answer * Incoming Average Holding Time Per Call (sec)
					BigDecimal inTotalHour = inCallAnswer.multiply(inAverageHoldingTPC).setScale(2, BigDecimal.ROUND_HALF_UP);
					excelAll.setInTotalHour(inTotalHour.toString());
					//计算 Outgoing total seconds in the hour
					BigDecimal outCallAnswer = new BigDecimal(row[14]);
					BigDecimal outAverageHoldingTPC = new BigDecimal(row[18]);
					//Outgoing Call Answer * Outgoing Average Holding Time Per Call (sec)
					BigDecimal outTotalHour = outCallAnswer.multiply(outAverageHoldingTPC).setScale(2, BigDecimal.ROUND_HALF_UP);
					excelAll.setOutTotalHour(outTotalHour.toString());
					//计算 Occupancy in hours
					//Total hour = Incoming total hour + Outgoing total hour
					//Occupancy hour = Total hour / 3600
					BigDecimal occHour = (inTotalHour.add(outTotalHour))
							.divide(new BigDecimal("3600"), 4, RoundingMode.HALF_UP);
					BigDecimal occHour2 = occHour.setScale(2, RoundingMode.HALF_UP);
					excelAll.setOccupancyHour(occHour2.toString());
					//计算 Occupancy Rate
					String serviceCapacity = row[19];
					//Occupancy Rate(%) = Occupancy Hour / T1 Capacity(Service Capacity) * 100
					BigDecimal occRate = (occHour.divide(new BigDecimal(serviceCapacity), 4, RoundingMode.HALF_UP))
							.multiply(new BigDecimal("100")).setScale(2, BigDecimal.ROUND_HALF_UP);
					excelAll.setOccupancyRate(occRate.toString());
					excelDao.addExcelAll(excelAll);
				}
			}
			//单独计算Total
			ExcelAll excelTotal = new ExcelAll();
			ExcelTotal totals = excelDao.selectExcelAllTotal(excelId);
			String[] row28 = list.get(27);
			if (row28 == null)return new JsonResult("Total行不为空！");
			excelTotal.setExcelId(excelId);
			excelTotal.setTime(row28[0]);
			excelTotal.setOutCallAnswer(totals.getOutCallAnswerTotal());
			excelTotal.setOutAverageHoldingTPC(totals.getOutAverageHoldingTPCTotal());
			excelTotal.setServiceCapacity(totals.getServiceCapacityTotal());
			excelTotal.setCapacityNeeded(totals.getCapacityNeededTotal());
			//计算 Outgoing total seconds in the hour
			BigDecimal outCallAnswer = new BigDecimal(totals.getOutCallAnswerTotal());
			BigDecimal outAverageHoldingTPC = new BigDecimal(totals.getOutAverageHoldingTPCTotal());
			//Outgoing Call Answer * Outgoing Average Holding Time Per Call (sec)
			BigDecimal outTotalHour = outCallAnswer.multiply(outAverageHoldingTPC).setScale(2, BigDecimal.ROUND_HALF_UP);
			excelTotal.setOutTotalHour(outTotalHour.toString());
			BigDecimal occHour = null;
			if(excelType.equals("Incoming And Outgoing")){//Incoming And Outgoing
				excelTotal.setInCallAnswer(totals.getInCallAnswerTotal());
				excelTotal.setInAverageHoldingTPC(totals.getInAverageHoldingTPCTotal());
				//计算 Incoming total seconds in the hour
				BigDecimal inCallAnswer = new BigDecimal(totals.getInCallAnswerTotal());
				BigDecimal inAverageHoldingTPC = new BigDecimal(totals.getInAverageHoldingTPCTotal());
				//Incoming Call Answer * Incoming Average Holding Time Per Call (sec)
				BigDecimal inTotalHour = inCallAnswer.multiply(inAverageHoldingTPC).setScale(2, BigDecimal.ROUND_HALF_UP);
				excelTotal.setInTotalHour(inTotalHour.toString());			
				//计算 Occupancy in hours
				//Total hour = Incoming total hour + Outgoing total hour
				//Occupancy hour = Total hour / 3600
				occHour = (inTotalHour.add(outTotalHour))
						.divide(new BigDecimal("3600"), 4, RoundingMode.HALF_UP);
				BigDecimal occHour2 = occHour.setScale(2, RoundingMode.HALF_UP);
				excelTotal.setOccupancyHour(occHour2.toString());
			}else{
				occHour = outTotalHour.divide(new BigDecimal("3600"), 4, RoundingMode.HALF_UP);
				BigDecimal occHour2 = occHour.setScale(2, RoundingMode.HALF_UP);
				excelTotal.setOccupancyHour(occHour2.toString());
			}
			//计算 Occupancy Rate
			//Occupancy Rate(%) = Occupancy Hour / T1 Capacity(Service Capacity) * 100
			BigDecimal occRate = (occHour
					.divide(new BigDecimal(totals.getServiceCapacityTotal()), 4, RoundingMode.HALF_UP))
					.multiply(new BigDecimal("100")).setScale(2, BigDecimal.ROUND_HALF_UP);
			excelTotal.setOccupancyRate(occRate.toString());
			
			//单独计算Max
			ExcelAll excelAllMax = excelDao.selectExcelAllMax(excelId);
			excel.setOccupancyRate(excelAllMax.getOccupancyRate());
			excelDao.updateExcel(excel);
			excelDao.addExcelAll(excelTotal);
			
			//v0.1
//				//单独计算Total保存到数据库
//				ExcelTotal totals = excelDao.selectExcelAllTotal(excelId);
//				String[] row28 = csvList.get(27);//获取第28行数据
//				if (row28 == null)return new JsonResult("Total行不为空！");
//				excelAll.setExcelId(excelId);
//				excelAll.setTime(row28[0]);
//				if(excelType.equals("Incoming And Outgoing")){//Incoming And Outgoing
//					excelAll.setInCallAnswer(totals.getInCallAnswerTotal());
//					excelAll.setInAverageHoldingTPC(totals.getInAverageHoldingTPCTotal());
//					excelAll.setInTotalHour(totals.getInTotalHourTotal());
//					excelAll.setServiceCapacity(row28[19]);
//				}else{
//					excelAll.setServiceCapacity(row28[10]);
//				}
//				excelAll.setOutCallAnswer(totals.getOutCallAnswerTotal());
//				excelAll.setOutAverageHoldingTPC(totals.getOutAverageHoldingTPCTotal());
//				excelAll.setOutTotalHour(totals.getOutTotalHourTotal());		
////				excelAll.setServiceCapacity(totals.getServiceCapacityTotal());
//				excelAll.setCapacityNeeded(totals.getCapacityNeededTotal());		
//				excelAll.setOccupancyHour(totals.getOccupancyHourTotal());
//				excelAll.setOccupancyRate(totals.getOccupancyRateTotal());
		}else{
			return new JsonResult("上传文件格式错误！");
		}
		return new JsonResult("Upload successful", 1);
	}
	/**
	 * 导出Excel
	 */
	@Override
	public Workbook batchExport(int excelId){
		Excel excel = excelDao.selectExcelById(excelId);
		HSSFWorkbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet(excel.getExcelDate());
		Row row = null;
		row = sheet.createRow(0);//创建第1单元行
		row.createCell(0).setCellValue(excel.getExcelName());
		row = sheet.createRow(3);//创建第4单元行
		row.createCell(0).setCellValue("Date:");
		row.createCell(1).setCellValue(excel.getExcelDate());
		row.createCell(2).setCellValue(excel.getWeek());
		List<ExcelAll> excelAllList = excelDao.selectExcelAllById(excelId);
		if("".equals(excelAllList.get(0).getInCallAnswer()) 
				|| excelAllList.get(0).getInCallAnswer() == null){
			//Outgoing Only Excel
			row = sheet.createRow(5);//创建第5单元行
			row.createCell(0).setCellValue("Time");
			row.createCell(1).setCellValue("Outgoing Call Answer");
			row.createCell(2).setCellValue("Outgoing Average Holding Time Per Call (sec)");
			row.createCell(3).setCellValue("Outgoing Total Seconds In The Hour (sec)");
			row.createCell(4).setCellValue("Service Capacity");
			row.createCell(5).setCellValue("Capacity Needed");
			row.createCell(6).setCellValue("Occupancy Hour(hour)");
			row.createCell(7).setCellValue("Occupancy Rate(%)");
			for (int i = 0; i < excelAllList.size(); i++) {
				row = sheet.createRow(i + 6);
				ExcelAll excelAll = excelAllList.get(i);
				row.createCell(0).setCellValue(excelAll.getTime());
				row.createCell(1).setCellValue(excelAll.getOutCallAnswer());
				row.createCell(2).setCellValue(excelAll.getOutAverageHoldingTPC());
				row.createCell(3).setCellValue(excelAll.getOutTotalHour());
				row.createCell(4).setCellValue(excelAll.getServiceCapacity());
				row.createCell(5).setCellValue(excelAll.getCapacityNeeded());
				row.createCell(6).setCellValue(excelAll.getOccupancyHour());
				row.createCell(7).setCellValue(excelAll.getOccupancyRate()+"%");
			}
			for (int i = 0; i <= 21; i++) {
				if(i == 0) {
					sheet.setColumnWidth(0, 12 * 256);
				} else{
					sheet.autoSizeColumn(i);			
				}
			}
		} else {
			//Incoming And Outgoing Excel
			row = sheet.createRow(5);
			row.createCell(0).setCellValue("Time");
			row.createCell(1).setCellValue("Incoming Call Answer");
			row.createCell(2).setCellValue("Incoming Average Holding Time Per Call (sec)");
			row.createCell(3).setCellValue("Incoming Total Seconds In The Hour (sec)");
			row.createCell(4).setCellValue("Outgoing Call Answer");
			row.createCell(5).setCellValue("Outgoing Average Holding Time Per Call (sec)");
			row.createCell(6).setCellValue("Outgoing Total Seconds In The Hour (sec)");
			row.createCell(7).setCellValue("Service Capacity");
			row.createCell(8).setCellValue("Capacity Needed");
			row.createCell(9).setCellValue("Occupancy Hour(hour)");
			row.createCell(10).setCellValue("Occupancy Rate(%)");
			for (int i = 0; i < excelAllList.size(); i++) {
				row = sheet.createRow(i + 6);
				ExcelAll excelAll = excelAllList.get(i);
				row.createCell(0).setCellValue(excelAll.getTime());
				row.createCell(1).setCellValue(excelAll.getInCallAnswer());
				row.createCell(2).setCellValue(excelAll.getInAverageHoldingTPC());
				row.createCell(3).setCellValue(excelAll.getInTotalHour());
				row.createCell(4).setCellValue(excelAll.getOutCallAnswer());
				row.createCell(5).setCellValue(excelAll.getOutAverageHoldingTPC());
				row.createCell(6).setCellValue(excelAll.getOutTotalHour());
				row.createCell(7).setCellValue(excelAll.getServiceCapacity());
				row.createCell(8).setCellValue(excelAll.getCapacityNeeded());
				row.createCell(9).setCellValue(excelAll.getOccupancyHour());
				row.createCell(10).setCellValue(excelAll.getOccupancyRate()+"%");
			}
			for (int i = 0; i <= 21; i++) {
				if(i == 0) {
					sheet.setColumnWidth(0, 12 * 256);
				} else{
					sheet.autoSizeColumn(i);
				}
			}
		}
		return wb;
	}
	/**删除Excel*/
	@Override
	public int deleteExcel(Integer... excelIds){
		int deleteExcel = 0;
		for(int excelId : excelIds){
			int deleteAll = excelDao.deleteExcelAll(excelId);
			if(deleteAll != 0)deleteExcel = excelDao.deleteExcel(excelId);
		}
		return deleteExcel;
	}
	
}