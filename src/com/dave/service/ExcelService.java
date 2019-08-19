package com.dave.service;

import com.dave.common.vo.JsonResult;
import com.dave.common.vo.PageObject;
import com.dave.entity.Excel;
import com.dave.entity.ExcelAll;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

/**
 * Excel业务层接口
 * @author davewpw
 *
 */
public interface ExcelService {
	/**
	 * 根据Excel ID 查询Excel内容
	 * @param excelId
	 * @return
	 */
	List<ExcelAll> selectExcelAll(int excelId);
	/**
	 * 根据Excel ID 查询Excel
	 * @param excelId
	 * @return
	 */
	Excel selectExcelById(int excelId);
	/**
	 * 导入Excel
	 * @param fileName
	 * @param file
	 * @return
	 * @throws Exception
	 */
	JsonResult batchImport(String fileName, MultipartFile file) throws Exception;
	/**
	 * 导出Excel
	 * @param excelId
	 * @return
	 */
	Workbook batchExport(int excelId);
	/**
	 * 删除Excel
	 * @param excelIds
	 * @return
	 */
	int deleteExcel(Integer... excelIds);
	/**
	 * 查询框功能查询所有 Excel,以及页面初始数据查询
	 * @param excelDate
	 * @param excelName
	 * @param isSearchMax
	 * @return
	 */
    PageObject<Excel> searchExcel(String excelDate, String excelName, int isSearchMax, int pageCurrent);
}