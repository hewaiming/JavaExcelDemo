package com.hewaiming.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) {
		//生成Workbook
		HSSFWorkbook workbook=new HSSFWorkbook();
		HSSFSheet sheet=workbook.createSheet("sheet1");
		HSSFRow row=sheet.createRow(1);
		row.setHeight((short) 300);
		HSSFCell cell_id=row.createCell(0);
		cell_id.setCellValue("ID");
		HSSFCell cell_name=row.createCell(1);
		cell_name.setCellValue("NAME");
		HSSFCell cell_address=row.createCell(2);
		cell_address.setCellValue("Address");
		
		FileOutputStream oStream=null;
		try {
			oStream=new FileOutputStream("aostar.xls");
			workbook.write(oStream);
			oStream.flush();
			oStream.close();
		} catch (Exception e) {			
			e.printStackTrace();
		}
		System.out.println("生产XLS成功！");

	}

}
