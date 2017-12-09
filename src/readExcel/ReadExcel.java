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
 * @date 2017��7��23��
 */
public class ReadExcel {
	public static void main(String[] args) {
		try {
		    System.out.println("----------");
			File filePath = new File("C:/Users/Administrator/Desktop/sysm_user.xlsx");  
			showExcel(filePath);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * ��ȡ���е�excel������
	 */
	@SuppressWarnings("resource")
	public static void showExcel(File filePath) throws Exception {
		// ���� XSSFWorkbook����,filePath �����ļ�·��
		// ��ȡExcel 2003
		// HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(filePath));
		// ��ȡExcel 2007
		XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(filePath));//�õ�excel����
		
		XSSFSheet sheet = null;
		for (int i = 0; i < workbook.getNumberOfSheets(); i++) {// ��ȡÿ��Sheet��
			//workbook.getSheet(arg0)
			sheet = workbook.getSheetAt(i);//ѭ���õ�Sheet
			System.out.println(sheet.getSheetName());
			for (int j = 0; j <= sheet.getLastRowNum(); j++) {// ��ȡÿ��
				XSSFRow row = sheet.getRow(j);//�õ�ÿһ��
				if (null != row) {// ���ղ�����rowѭ����ȡÿ����Ԫ��
					for (int k = 0; k < row.getLastCellNum(); k++) {// ��ȡÿ����Ԫ��
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