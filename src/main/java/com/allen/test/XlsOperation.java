package com.allen.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XlsOperation {

	public static final String HEADERINFO = "headInfo";
	public static final String DATAINFON = "dataInfo";

	/**
	 * 
	 * @Title: getWeebWork
	 * @Description: TODO(���ݴ�����ļ�����ȡ����������(Workbook))
	 * @param filename
	 * @return
	 * @throws IOException
	 */
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

	/**
	 * 
	 * @Title: writeExcel
	 * @Description: TODO(����Excel��)
	 * @param pathname
	 *            :����Excel����ļ�·��
	 * @param map
	 *            ����װ��Ҫ����������(HEADERINFO��װ��ͷ��Ϣ��DATAINFON����װҪ������������Ϣ,�˴���Ҫʹ��TreeMap
	 *            ) ���磺 map.put(ExcelUtil.HEADERINFO,List<String> headList);
	 *            map.put(ExcelUtil.DATAINFON,List<TreeMap<String,Object>>
	 *            dataList);
	 * @param wb
	 * @throws IOException
	 */
	public static void writeExcel(String pathname, Map<String, Object> map,
			Workbook wb) throws IOException {
		if (null != map && null != pathname) {
			List<Object> headList = (List<Object>) map
					.get(XlsOperation.HEADERINFO);
			List<TreeMap<String, Object>> dataList = (List<TreeMap<String, Object>>) map
					.get(XlsOperation.DATAINFON);
			CellStyle style = getCellStyle(wb);
			Sheet sheet = wb.createSheet();// ���ĵ������д���һ����..Ĭ���Ǳ�������Sheet0��Sheet1....
			// Sheet sheet = wb.createSheet("hell poi");//�ڴ����������ʱ��ָ����������
			
			/**
			 * ����Excel��ĵ�һ�м���ͷ
			 */
			Row row = sheet.createRow(0);
			for (int i = 0; i < headList.size(); i++) {
				Cell headCell = row.createCell(i);
				headCell.setCellType(Cell.CELL_TYPE_STRING);// ���������Ԫ������ݵ�����,���ı����ͻ�����������
				headCell.setCellStyle(style);// ���ñ�ͷ��ʽ
				headCell.setCellValue(String.valueOf(headList.get(i)));// �������Ԫ������ֵ
			}

			for (int i = 0; i < dataList.size(); i++) {
				Row rowdata = sheet.createRow(i + 1);// ����������
				TreeMap<String, Object> mapdata = dataList.get(i);
				Iterator it = mapdata.keySet().iterator();
				int j = 0;
				while (it.hasNext()) {
					String strdata = String.valueOf(mapdata.get(it.next()));
					Cell celldata = rowdata.createCell(j);// ��һ���д���ĳ��..
					celldata.setCellType(Cell.CELL_TYPE_STRING);
					celldata.setCellValue(strdata);
					j++;
				}
			}

			// �ļ���
			File file = new File(pathname);
			OutputStream os = new FileOutputStream(file);
			os.flush();
			wb.write(os);
			os.close();
		}
	}

	/**
	 * 
	 * @Title: getCellStyle
	 * @Description: TODO�����ñ�ͷ��ʽ��
	 * @param wb
	 * @return
	 */
	public static CellStyle getCellStyle(Workbook wb) {
		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setFontName("����");
		font.setFontHeightInPoints((short) 12);// ���������С
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);// �Ӵ�
		style.setFillForegroundColor(HSSFColor.LIME.index);// ���ñ���ɫ
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setAlignment(HSSFCellStyle.SOLID_FOREGROUND);// �õ�Ԫ�����
		// style.setWrapText(true);//�����Զ�����
		style.setFont(font);
		return style;
	}

	/**
	 * 
	 * @Title: readerExcelDemo
	 * @Description: TODO(��ȡExcel���е�����)
	 * @throws IOException
	 */
	public static void readFromExcelDemo() throws IOException {
		/**
		 * ��ȡExcel���е���������
		 */
		Workbook workbook = getWeebWork("E:/test.xlsx");
		System.out.println("�ܱ�ҳ��Ϊ��" + workbook.getNumberOfSheets());// ��ȡ��ҳ��
		Sheet sheet = workbook.getSheetAt(0);
		// Sheet sheet = workbook.getSheetAt(1);
		int rownum = sheet.getLastRowNum();// ��ȡ������
		for (int i = 0; i <= rownum; i++) {
			Row row = sheet.getRow(i);
			Cell orderno = row.getCell(2);// ��ȡָ����Ԫ���е�����
			// System.out.println(orderno.getCellType());//�����ӡ����cell��type
			short cellnum = row.getLastCellNum(); // ��ȡ��Ԫ���������
			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				Cell celldata = row.getCell(j);
				System.out.print(celldata + "\t");
			}
			System.out.println();
		}

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

	public static void readFromExcelDemo1() throws IOException {
		/**
		 * ��ȡExcel���е���������
		 * 
		 * Workbook: excel���ĵ����� sheet: excel�ı� row: excel�е��� cell: excel�еĵ�Ԫ����
		 * 
		 */
		Workbook workbook = getWeebWork("E:/test.xlsx");
		System.out.println("�ܱ�ҳ��Ϊ��" + workbook.getNumberOfSheets());// ��ȡ��ҳ��
		// Sheet sheet =workbook.getSheetAt(0);
		Sheet sheet = workbook.getSheetAt(2);// ��ȡ�ڶ�����
		int rownum = sheet.getLastRowNum();// ��ȡ������
		for (int i = 0; i <= rownum; i++) {
			Row row = sheet.getRow(i);// ��ȡ���ĵ�i��
			// Cell orderno =
			// row.getCell(2);//��ȡָ����Ԫ���е�����(��ȡһ���еĵ�2��(�����2ָ����0,1,2.���ڵ���λ))
			// System.out.println(orderno.getCellType());//�����ӡ����cell��type
			// short cellnum=row.getLastCellNum();
			// //��ȡ��Ԫ���������(��ȡһ�����ж��ٸ���Ԫ��(Ҳ���Ƕ�����))

			/**
			 * row.getFirstCellNum(): ��ȡ�еĵ�һ����Ԫ���λ�� row.getLastCellNum():
			 * ��ȡ�е����һ����Ԫ���λ��
			 */
			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {// ����һ���е�������
				Cell celldata = row.getCell(j);// ��ȡһ���еĵ�j�з���Cell���͵�����
				System.out.print(celldata + "\t");//
			}

			// ��ӡָ����
			// Cell celldata = row.getCell(4);//��ȡ��һ���еĵ�4��(�ڵ�5��λ����)
			// System.out.print( "\"" + celldata+"\",");

			System.out.println();
		}

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

	
	public static void readFromExcelDemo(String fileAbsolutePath) throws IOException {
		/**
		 * ��ȡExcel���е���������
		 */
		Workbook workbook = getWeebWork(fileAbsolutePath);
		System.out.println("�ܱ�ҳ��Ϊ��" + workbook.getNumberOfSheets());// ��ȡ��ҳ��
		Sheet sheet = workbook.getSheetAt(0);
		// Sheet sheet = workbook.getSheetAt(1);
		int rownum = sheet.getLastRowNum();// ��ȡ������
		for (int i = 0; i <= rownum; i++) {
			Row row = sheet.getRow(i);
			Cell orderno = row.getCell(2);// ��ȡָ����Ԫ���е�����
			// System.out.println(orderno.getCellType());//�����ӡ����cell��type
			short cellnum = row.getLastCellNum(); // ��ȡ��Ԫ���������
			for (int j = row.getFirstCellNum(); j < row.getLastCellNum(); j++) {
				Cell celldata = row.getCell(j);
				System.out.print(celldata + "\t");
			}
			System.out.println();
		}

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
	
	
	public static void main(String[] args) throws IOException {
//		 readFromExcelDemo1();

//		String filePath = "E:/test.xlsx";
//		readFromExcelDemo(filePath);//��һ��ָ����excel�ļ��ж�ȡ����
		
		
//		writeToExcelDemo();
	}
	
	

	public static void writeToExcelDemo() throws IOException {
		/**
		 * HSSF: .xls XSSF: .xlsx ���Դ�һ��ڹ����п���Excel = HSSF+XSSF
		 * 
		 * HSSF��POI���̶�Excel 97(-2007)�ļ������Ĵ�Javaʵ�� XSSF��POI���̶�Excel 2007 OOXML
		 * (.xlsx)�ļ������Ĵ�Javaʵ��
		 * 
		 * ��POI 3.8�汾��ʼ���ṩ��һ�ֻ���XSSF�ĵ��ڴ�ռ�õ�API----SXSSF
		 * 
		 */

		Workbook wb = new XSSFWorkbook();// ����һ���µ�excel���ĵ�����
		Map<String, Object> map = new HashMap<String, Object>();
		List headList = new ArrayList();// ��ͷ����
		headList.add("�µ�ʱ��");
		headList.add("����ʱ��");
		headList.add("�������");
		headList.add("�������");
		headList.add("�û���");// excel�Ķ�

		/**
		 * TreeMap���ں����ʵ��
		 */
		List dataList = new ArrayList();// ����ڵ�����
		for (int i = 0; i < 15; i++) {
			TreeMap<String, Object> treeMap = new TreeMap<String, Object>();// �˴������ݱ���Ϊ�������ݣ�����ʹ��TreeMap���з�װ
			treeMap.put("m1", "2013-10-" + i + 1);
			treeMap.put("m2", "2013-11-" + i + 1);
			treeMap.put("m3", "20124" + i + 1);
			treeMap.put("m4", 23.5 + i + 1);
			treeMap.put("m5", "����_" + i);
			dataList.add(treeMap);
		}

		/*
		 * �Ȳ�Ҫ����������һ��,�������´���: Cannot get a numeric value from a text
		 * cell(���ܴ�һ��text cell�л�ȡ�������͵�����)
		 * 
		 * ����취: http://blog.csdn.net/ysughw/article/details/9288307
		 */
		// TreeMap<String,Object> treeMap1 = new TreeMap<String, Object>();
		// treeMap1.put("asd", null);
		// treeMap1.put("��ͷ", "zhutou");
		// dataList.add(treeMap1);
		map.put(XlsOperation.HEADERINFO, headList);
		map.put(XlsOperation.DATAINFON, dataList);
		writeExcel("E:/test1.xlsx", map, wb);//��wb����дmap�����ݣ�����E:/test1.xlsx����ļ�....
	}
}