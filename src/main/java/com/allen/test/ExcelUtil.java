package com.allen.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class ExcelUtil {
	
	
	//��ȡexcel��sheetҳ������
	@Test
	public  void readFromExcel() throws IOException {
		/**
		 * ��ȡExcel���е���������
		 * Workbook: excel���ĵ����� sheet: excel�ı� row: excel�е��� cell: excel�еĵ�Ԫ����
		 */
		Workbook workbook = getWeebWork("E:/test.xlsx");
		System.out.println("�ܱ�ҳ��Ϊ��" + workbook.getNumberOfSheets());// ��ȡ��ҳ��
		Sheet sheet =workbook.getSheetAt(1);
		//Sheet sheet = workbook.getSheetAt(2); ��ȡ�ڶ�����
		int rownum = sheet.getLastRowNum();// ��ȡ������
		System.out.println("������ͷ���м�¼����Ϊ��"  + rownum);
//		for (int i = 0; i <= rownum; i++) {
//			Row row = sheet.getRow(i);// ��ȡ���ĵ�i��
//			//row.getFirstCellNum(): ��ȡ�еĵ�һ����Ԫ���λ�� row.getLastCellNum():��ȡ�е����һ����Ԫ���λ��
//			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {// ����һ���е�������
//				Cell celldata = row.getCell(j);// ��ȡһ���еĵ�j�з���Cell���͵�����
//				System.out.print( celldata + "\t");
//			}
//			System.out.println();
//		}

		/**
		 * ��ȡָ��λ�õĵ�Ԫ��
		 */
		// Row row1 = sheet.getRow(1);
		// Cell cell1 = row1.getCell(2);
		// System.out.print("(1,2)λ�õ�Ԫ���ֵΪ��"+cell1);
		// BigDecimal big = new
		// BigDecimal(cell1.getNumericCellValue());//����ѧ��������ʾ������ת��ΪString����
		// System.out.print("\t"+String.valueOf(big));

	}
	
	
	public static Workbook getWeebWork(String filename) throws IOException {
		Workbook workbook = null;
		if (null != filename) {
			String fileType = filename.substring(filename.lastIndexOf("."),
					filename.length());
			FileInputStream fileStream = new FileInputStream(new File(filename));
			if (".xls".equals(fileType.trim().toLowerCase())) {
				workbook = new HSSFWorkbook(fileStream);// ���� Excel 2003 ����������
			} else if (".xlsx".equals(fileType.trim().toLowerCase())) {
				workbook = new XSSFWorkbook(fileStream);// ���� Excel 2007 ����������
			}
		}
		return workbook;
	}
}
