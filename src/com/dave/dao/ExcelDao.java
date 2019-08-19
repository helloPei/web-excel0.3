package com.dave.dao;

import com.dave.entity.Excel;
import com.dave.entity.ExcelAll;
import com.dave.entity.ExcelTotal;
import org.apache.ibatis.annotations.Delete;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.annotations.Select;

import java.util.List;

/**
 * Excel数据持久层接口
 * @author davewpw
 *
 */
public interface ExcelDao {
	/**根据 ExcelID 查询所有 Excel内容*/
	List<ExcelAll> selectExcelAll(int excelId);
	/**根据 ExcelID 查询 Excel*/
	@Select("SELECT * FROM EXCEL_REPORT WHERE EXCEL_ID = #{excelId}")
	Excel selectExcelById(@Param("excelId") int excelId);
	/**添加 Excel*/
	int addExcel(Excel excel);
	/**更改 Excel, 添加 Occupancy Rate 的内容*/
	int updateExcel(Excel excel);
	/**根据 Excel名称查询, 最新的自增 ExcelID*/
	int selectExcelByName(String excelName);
	/**添加 Excel内容*/
	int addExcelAll(ExcelAll excelAll);
	/**根据 ExcelID 查询 Excel内容*/
	@Select("SELECT * FROM EXCEL_REPORT_ALL WHERE EXCEL_ID = #{excelId} ORDER BY TIME ASC")
	List<ExcelAll> selectExcelAllById(@Param("excelId") int excelId);
	/**根据 ExcelID 删除 Excel内容*/
	@Delete("DELETE FROM EXCEL_REPORT_ALL WHERE EXCEL_ID = #{excelId}")
	int deleteExcelAll(@Param("excelId") int excelId);
	/**根据 ExcelID 删除 Excel*/
	@Delete("DELETE FROM EXCEL_REPORT WHERE EXCEL_ID = #{excelId}")
	int deleteExcel(@Param("excelId") int excelId);
	/**根据 ExcelID 查询 Excel 内容, 获取需要的Total*/
	ExcelTotal selectExcelAllTotal(int excelId);
	/**查询框功能查询所有 Excel,以及页面初始数据查询*/
	List<Excel> searchExcel(
			@Param("excelDate")String excelDate, @Param("excelName")String excelName,
			@Param("startIndex")int startIndex, @Param("pageSize")int pageSize
	);
	int selectCountExcel(@Param("excelDate")String excelDate, @Param("excelName")String excelName);
	/**根据 Excel名称查询 Excel Occupancy Rate 的最大数*/
	List<Excel> searchExcelMax(String excelName);
	/**根据 ExcelID 查询 Excel Occupancy Rate 的最大数*/
	ExcelAll selectExcelAllMax(int excelId);
}