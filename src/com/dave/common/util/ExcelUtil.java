package com.dave.common.util;

import java.io.InputStream;
import java.nio.charset.Charset;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.csvreader.CsvReader;

/**
 * Excel工具类
 * @author davewpw
 *
 */
public class ExcelUtil {
	public static List<String[]> readCsvByIs(InputStream csvIs) {
		List<String[]> csvList = new ArrayList<String[]>();
		try {
			CsvReader reader = new CsvReader(csvIs, Charset.forName("GBK"));
			while (reader.readRecord()) {
				csvList.add(reader.getValues());
			}
			reader.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return csvList;
	}
	public static List<String[]> readXlsAndXlsxByIs(InputStream is) {
		List<String[]> xlsList = new ArrayList<String[]>();
		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(is);
			Sheet sheet = wb.getSheetAt(0);
			for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {
				Row row = sheet.getRow(i);
				if (row == null)continue;
				String[] val = new String[row.getLastCellNum()];
				for (int r = 0; r < row.getLastCellNum(); r++) {
					Cell cell = row.getCell(r);
					if (cell != null) {
						String cellValue = null;
						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							if (cell.getCellStyle().getDataFormatString().indexOf("%") != -1) {
								NumberFormat numFormat = NumberFormat.getPercentInstance();
								numFormat.setMaximumFractionDigits(2);
								cellValue = numFormat.format(cell.getNumericCellValue());
							} else {
								cellValue = NumberFormat.getInstance().format(cell.getNumericCellValue());
							}
							break;
						case Cell.CELL_TYPE_STRING:
							cellValue = String.valueOf(cell.getStringCellValue());
							break;
						case Cell.CELL_TYPE_BOOLEAN:
							cellValue = String.valueOf(cell.getBooleanCellValue());
							break;
						case Cell.CELL_TYPE_FORMULA:
							cellValue = String.valueOf(cell.getCellFormula());
							break;
						case Cell.CELL_TYPE_BLANK:
							cellValue = "";
							break;
						case Cell.CELL_TYPE_ERROR:
							cellValue = "非法字符";
							break;
						default:
							cellValue = "未知类型";
							break;
						}
						val[r] = cellValue.trim();
					}
				}
				xlsList.add(val);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return xlsList;
	}

}