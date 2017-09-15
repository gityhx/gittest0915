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
	
	
	//读取excel中sheet页的内容
	@Test
	public  void readFromExcel() throws IOException {
		/**
		 * 读取Excel表中的所有数据
		 * Workbook: excel的文档对象 sheet: excel的表单 row: excel中的行 cell: excel中的单元格子
		 */
		Workbook workbook = getWeebWork("E:/test.xlsx");
		System.out.println("总表页数为：" + workbook.getNumberOfSheets());// 获取表页数
		Sheet sheet =workbook.getSheetAt(1);
		//Sheet sheet = workbook.getSheetAt(2); 获取第二个表单
		int rownum = sheet.getLastRowNum();// 获取总行数
		System.out.println("不包括头，有记录条数为："  + rownum);
//		for (int i = 0; i <= rownum; i++) {
//			Row row = sheet.getRow(i);// 获取表达的第i行
//			//row.getFirstCellNum(): 获取行的第一个单元格的位置 row.getLastCellNum():获取行的最后一个单元格的位置
//			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {// 遍历一行中的所有列
//				Cell celldata = row.getCell(j);// 获取一行中的第j列返回Cell类型的数据
//				System.out.print( celldata + "\t");
//			}
//			System.out.println();
//		}

		/**
		 * 读取指定位置的单元格
		 */
		// Row row1 = sheet.getRow(1);
		// Cell cell1 = row1.getCell(2);
		// System.out.print("(1,2)位置单元格的值为："+cell1);
		// BigDecimal big = new
		// BigDecimal(cell1.getNumericCellValue());//将科学计数法表示的数据转化为String类型
		// System.out.print("\t"+String.valueOf(big));

	}
	
	
	public static Workbook getWeebWork(String filename) throws IOException {
		Workbook workbook = null;
		if (null != filename) {
			String fileType = filename.substring(filename.lastIndexOf("."),
					filename.length());
			FileInputStream fileStream = new FileInputStream(new File(filename));
			if (".xls".equals(fileType.trim().toLowerCase())) {
				workbook = new HSSFWorkbook(fileStream);// 创建 Excel 2003 工作簿对象
			} else if (".xlsx".equals(fileType.trim().toLowerCase())) {
				workbook = new XSSFWorkbook(fileStream);// 创建 Excel 2007 工作簿对象
			}
		}
		return workbook;
	}
}
