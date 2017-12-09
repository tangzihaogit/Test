package readExcel;

import java.io.File;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * ��ȡexcl
 * @����   LiuCong
 *
 * @ʱ�� 2017��5��2������10:38:57
 */
public class TestReadMergeRegionExcel {

	public static void main(String[] args) throws Exception {     
		new TestReadMergeRegionExcel().readExcelToObj("C:/Users/Administrator/Desktop/sysm_user.xlsx");
	}

	/**
	 * ��ȡexcel����
	 * 
	 * @param path
	 */
	private void readExcelToObj(String path) {

		Workbook wb = null;
		try {
			wb = WorkbookFactory.create(new File(path));
			readExcel(wb, 0, 1, 0);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * ��ȡexcel�ļ�
	 * 
	 * @param wb
	 * @param sheetIndex
	 *            sheetҳ�±꣺��0��ʼ
	 * @param startReadLine
	 *            ��ʼ��ȡ����:��0��ʼ
	 * @param tailLine
	 *            ȥ������ȡ����
	 */
	@SuppressWarnings("deprecation")
	private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) {
		Sheet sheet = wb.getSheetAt(sheetIndex);
		Row row = null;

		for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
			row = sheet.getRow(i);
			if (null != row) {
				//for (Cell c : row) {//row.getCell(k)
				for (int k = 0; k < row.getLastCellNum(); k++) {// ��ȡÿ����Ԫ��
					Cell c = row.getCell(k);
					
					boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
					// �ж��Ƿ���кϲ���Ԫ��
					if (null != c) {
						if (isMerge) {
							String rs = getMergedRegionValue(sheet, row.getRowNum(), c.getColumnIndex());
							System.out.print(rs + " ");
						} else {
							switch (c.getCellType()) {
							case 0:
								System.out.print(c.getNumericCellValue() + " ");
								break;
							case 1:
								System.out.print(c.getStringCellValue() + " ");
								break;
							}
						}
					}
				}
				System.out.println();
			}
		}
	}

	/**
	 * ��ȡ�ϲ���Ԫ���ֵ
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public String getMergedRegionValue(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			int firstRow = ca.getFirstRow();
			int lastRow = ca.getLastRow();

			if (row >= firstRow && row <= lastRow) {

				if (column >= firstColumn && column <= lastColumn) {
					Row fRow = sheet.getRow(firstRow);
					Cell fCell = fRow.getCell(firstColumn);
					return getCellValue(fCell);
				}
			}
		}

		return null;
	}

	/**
	 * �жϺϲ�����
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	@SuppressWarnings("unused")
	private boolean isMergedRow(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row == firstRow && row == lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}

	/**
	 * �ж�ָ���ĵ�Ԫ���Ƿ��Ǻϲ���Ԫ��
	 * 
	 * @param sheet
	 * @param row
	 *            ���±�
	 * @param column
	 *            ���±�
	 * @return
	 */
	private boolean isMergedRegion(Sheet sheet, int row, int column) {
		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					return true;
				}
			}
		}
		return false;
	}

	/**
	 * �ж�sheetҳ���Ƿ��кϲ���Ԫ��
	 * 
	 * @param sheet
	 * @return
	 */
	@SuppressWarnings("unused")
	private boolean hasMerged(Sheet sheet) {
		return sheet.getNumMergedRegions() > 0 ? true : false;
	}

	/**
	 * �ϲ���Ԫ��
	 * 
	 * @param sheet
	 * @param firstRow
	 *            ��ʼ��
	 * @param lastRow
	 *            ������
	 * @param firstCol
	 *            ��ʼ��
	 * @param lastCol
	 *            ������
	 */
	@SuppressWarnings("unused")
	private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * ��ȡ��Ԫ���ֵ
	 * 
	 * @param cell
	 * @return
	 */
	@SuppressWarnings("deprecation")
	public String getCellValue(Cell cell) {

		if (cell == null)
			return "";

		if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

			return cell.getStringCellValue();

		} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

			return String.valueOf(cell.getBooleanCellValue());

		} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

			return cell.getCellFormula();

		} else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

			return String.valueOf(cell.getNumericCellValue());

		}
		return "";
	}
}
