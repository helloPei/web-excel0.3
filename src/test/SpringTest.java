package test;

import com.dave.dao.ExcelDao;
import com.dave.entity.Excel;
import org.junit.Test;
import org.springframework.beans.factory.annotation.Autowired;

import java.util.Date;
import java.util.List;

/**
 * Spring测试类
 * @author davewpw
 *
 */
public class SpringTest extends SpringTestBase {
	
	@Autowired
	private ExcelDao excelDao;
	
	@Test
	public void searchExcel(){
		String excelDate = null;
		String excelName = "1033";

		//excelDate = "2019-03-21";
//		excelName = "HTHK-CALL CENTRE - (1033";
		//List<Excel> excels = excelDao.searchExcel(excelDate, excelName);
		//ystem.out.println(excels.get(0).getExcelName());
	}

	@Test
	public void addExcel(){
		Excel excel = new Excel();
		excel.setExcelName("aaa");
		excel.setExcelDate("2019-07-31");
		excel.setWeek("星期三");
		excel.setCreateDate(new Date());
		excelDao.addExcel(excel);
		System.out.println(excel.getExcelId());
	}

	public void helloGit() {
		System.out.println("left");
		System.out.println("right");
		System.out.println("develop");
		System.out.println("test IDEA");
		System.out.println("test Add");
		System.out.println("eclipse update develop");
		System.out.println("IDEA update master");
        System.out.println("eclipse pull");
        System.out.println("eclipse update master");
        System.out.println("IDEA update master2");
	}
	
}