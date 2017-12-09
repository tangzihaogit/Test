package readExcel;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 
 * @author liucong
 *
 * @date 2017年7月23日
 */
public class ReadExcel {
	public static void main(String[] args) {
		try {
		    System.out.println("吊打唐子壕");
			File filePath = new File("C:/Users/Administrator/Desktop/sysm_user.xlsx");  
			showExcel(filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 读取所有的excel表内容
	 */
	@SuppressWarnings("resource")
	public static void showExcel(File filePath) throws Exception {
		// 构造 XSSFWorkbook对象,filePath 传入文件路径
		// 读取Excel 2003
		// HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(filePath));
		// 读取Excel 2007
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath));//得到excel对象
		
		XSSFSheet sheet = null;
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// 获取每个Sheet表
			//workbook.getSheet(arg0)
			sheet = workbook.getSheetAt(i);//循环得到Sheet
			System.out.println(sheet.getSheetName());
			for (int j = 0; j <= sheet.getLastRowNum(); j++) {// 获取每行
				XSSFRow row = sheet.getRow(j);//得到每一行
				if (null != row) {// 当空不等于row循环获取每个单元格
					for (int k = 0; k < row.getLastCellNum(); k++) {// 获取每个单元格
						if (null != row.getCell(k) || "".equals(row.getCell(k))) {
							System.out.print(row.getCell(k) + "\t");
						}
					}
				}
				System.out.println();
			}
		}
	}
}
