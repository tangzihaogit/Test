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
 * д��excl
 * @����   LiuCong
 *
 * @ʱ�� 2017��5��2������10:39:06
 */
public class TestWrite2 {

        /**
         * @param args
         */
        @SuppressWarnings("deprecation")
		public static void main(String[] args) throws Exception {
                // ����Excel�Ĺ������ Workbook,��Ӧ��һ��excel�ĵ�
                @SuppressWarnings("resource")
				HSSFWorkbook wb = new HSSFWorkbook();

                // ����Excel�Ĺ���sheet,��Ӧ��һ��excel�ĵ���tab
                
                wb.createSheet("���");
                HSSFSheet sheet = wb.createSheet("sheet1");

                // ����excelÿ�п��
                //sheet.setColumnWidth(0, 4000);
                //sheet.setColumnWidth(1, 3500);

                // ����������ʽ
                HSSFFont font = wb.createFont();
                font.setFontName("Verdana");
                font.setBoldweight((short) 100);
                font.setFontHeight((short) 300);
                font.setColor(HSSFColor.BLUE.index);

                //ROW  CELL
                // ������Ԫ����ʽ
//                HSSFCellStyle style = wb.createCellStyle();
//                style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
//                style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
//                style.setFillForegroundColor(HSSFColor.LIGHT_TURQUOISE.index);
//                style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

                // ���ñ߿�
                //style.setBottomBorderColor(HSSFColor.RED.index);
               // style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                //style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
               // style.setBorderRight(HSSFCellStyle.BORDER_THIN);
               // style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                
                HSSFCellStyle style = wb.createCellStyle();
                style.setFont(font);// ��������

                // ����Excel��sheet��һ��
                HSSFRow row = sheet.createRow(0);
                //row.setHeight((short) 500);// �趨�еĸ߶�
                // ����һ��Excel�ĵ�Ԫ��
                HSSFCell cell0 = row.createCell(0);
                HSSFCell cell1 = row.createCell(1);
                HSSFCell cell2 = row.createCell(2);
                cell0.setCellValue("���Ӻ�");
                cell1.setCellValue("��");
                cell2.setCellValue("��");
                
                cell0.setCellStyle(style);
                cell2.setCellStyle(style);

                // �ϲ���Ԫ��(startRow��endRow��startColumn��endColumn)
               // sheet.addMergedRegion(new CellRangeAddress(5, 6, 1, 3));

                // ��Excel�ĵ�Ԫ��������ʽ�͸�ֵ
                //cell.setCellStyle(style);
                //cell.setCellValue("hello world");

                // ���õ�Ԫ�����ݸ�ʽ
                //HSSFCellStyle style1 = wb.createCellStyle();
                //style1.setDataFormat(HSSFDataFormat.getBuiltinFormat("h:mm:ss"));

               // style1.setWrapText(true);// �Զ�����

                //row = sheet.createRow(1);

                // ���õ�Ԫ�����ʽ��ʽ

                //cell = row.createCell(0);
                //cell.setCellStyle(style1);
                //cell.setCellValue(new Date());

                // ����������
            /*    HSSFHyperlink link = new HSSFHyperlink(HSSFHyperlink.LINK_URL);
                link.setAddress("http://www.baidu.com");
                cell = row.createCell(1);
                cell.setCellValue("�ٶ�");
                cell.setHyperlink(link);// �趨��Ԫ�������
             */
                FileOutputStream os = new FileOutputStream("C:/Users/Administrator/Desktop/abc.xls");
                wb.write(os);
                os.close();

        }

}