package readExcel;

import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 写入excl
 * @作者   LiuCong
 *
 * @时间 2017年5月2日上午10:39:06
 */
public class TestWrite2 {

        /**
         * @param args
         */
        @SuppressWarnings("deprecation")
		public static void main(String[] args) throws Exception {
                // 创建Excel的工作书册 Workbook,对应到一个excel文档
                @SuppressWarnings("resource")
				HSSFWorkbook wb = new HSSFWorkbook();

                // 创建Excel的工作sheet,对应到一个excel文档的tab
                
                wb.createSheet("唐骞");
                HSSFSheet sheet = wb.createSheet("sheet1");

                // 设置excel每列宽度
                //sheet.setColumnWidth(0, 4000);
                //sheet.setColumnWidth(1, 3500);

                // 创建字体样式
                HSSFFont font = wb.createFont();
                font.setFontName("Verdana");
                font.setBoldweight((short) 100);
                font.setFontHeight((short) 300);
                font.setColor(HSSFColor.BLUE.index);

                //ROW  CELL
                // 创建单元格样式
//                HSSFCellStyle style = wb.createCellStyle();
//                style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//                style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
//                style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
//                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

                // 设置边框
                //style.setBottomBorderColor(HSSFColor.RED.index);
               // style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                //style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               // style.setBorderRight(HSSFCellStyle.BORDER_THIN);
               // style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                
                HSSFCellStyle style = wb.createCellStyle();
                style.setFont(font);// 设置字体

                // 创建Excel的sheet的一行
                HSSFRow row = sheet.createRow(0);
                //row.setHeight((short) 500);// 设定行的高度
                // 创建一个Excel的单元格
                HSSFCell cell0 = row.createCell(0);
                HSSFCell cell1 = row.createCell(1);
                HSSFCell cell2 = row.createCell(2);
                cell0.setCellValue("唐子壕");
                cell1.setCellValue("吃");
                cell2.setCellValue("饭");
                
                cell0.setCellStyle(style);
                cell2.setCellStyle(style);

                // 合并单元格(startRow，endRow，startColumn，endColumn)
               // sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 3));

                // 给Excel的单元格设置样式和赋值
                //cell.setCellStyle(style);
                //cell.setCellValue("hello world");

                // 设置单元格内容格式
                //HSSFCellStyle style1 = wb.createCellStyle();
                //style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));

               // style1.setWrapText(true);// 自动换行

                //row = sheet.createRow(1);

                // 设置单元格的样式格式

                //cell = row.createCell(0);
                //cell.setCellStyle(style1);
                //cell.setCellValue(new Date());

                // 创建超链接
            /*    HSSFHyperlink link = new HSSFHyperlink(HSSFHyperlink.LINK_URL);
                link.setAddress("http://www.baidu.com");
                cell = row.createCell(1);
                cell.setCellValue("百度");
                cell.setHyperlink(link);// 设定单元格的链接
             */
                FileOutputStream os = new FileOutputStream("C:/Users/Administrator/Desktop/abc.xls");
                wb.write(os);
                os.close();

        }

}