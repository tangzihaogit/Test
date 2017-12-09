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
 * 读取excl
 * @作者   LiuCong
 *
 * @时间 2017年5月2日上午10:38:57
 */
public class TestReadMergeRegionExcel {

	public static void main(String[] args) throws Exception {     
		new TestReadMergeRegionExcel().readExcelToObj("C:/Users/Administrator/Desktop/sysm_user.xlsx");
	}

	/**
	 * 读取excel数据
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
	 * 读取excel文件
	 * 
	 * @param wb
	 * @param sheetIndex
	 *            sheet页下标：从0开始
	 * @param startReadLine
	 *            开始读取的行:从0开始
	 * @param tailLine
	 *            去除最后读取的行
	 */
	@SuppressWarnings("deprecation")
	private void readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) {
		Sheet sheet = wb.getSheetAt(sheetIndex);
		Row row = null;

		for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
			row = sheet.getRow(i);
			if (null != row) {
				//for (Cell c : row) {//row.getCell(k)
				for (int k = 0; k < row.getLastCellNum(); k++) {// 获取每个单元格
					Cell c = row.getCell(k);
					
					boolean isMerge = isMergedRegion(sheet, i, c.getColumnIndex());
					// 判断是否具有合并单元格
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
	 * 获取合并单元格的值
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
	 * 判断合并了行
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
	 * 判断指定的单元格是否是合并单元格
	 * 
	 * @param sheet
	 * @param row
	 *            行下标
	 * @param column
	 *            列下标
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
	 * 判断sheet页中是否含有合并单元格
	 * 
	 * @param sheet
	 * @return
	 */
	@SuppressWarnings("unused")
	private boolean hasMerged(Sheet sheet) {
		return sheet.getNumMergedRegions() > 0 ? true : false;
	}

	/**
	 * 合并单元格
	 * 
	 * @param sheet
	 * @param firstRow
	 *            开始行
	 * @param lastRow
	 *            结束行
	 * @param firstCol
	 *            开始列
	 * @param lastCol
	 *            结束列
	 */
	@SuppressWarnings("unused")
	private void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
	}

	/**
	 * 获取单元格的值
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
