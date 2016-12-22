package com.hewaiming.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ReadExcel {

	public static void main(String[] args) {
	try {
		FileInputStream fileInputStream=new FileInputStream("aostar.xls");
		POIFSFileSystem ts=new POIFSFileSystem(fileInputStream);
		HSSFWorkbook workbook=new HSSFWorkbook(ts);
		HSSFSheet sheet=workbook.getSheetAt(0);
		HSSFRow row=null;
		for(int i=0;i<workbook.getNumberOfSheets();i++){
			sheet=workbook.getSheetAt(i);
			if(sheet==null){
				continue;
			}
		//	row=sheet.getRow(i);
			for(int j=0;j<=sheet.getLastRowNum();j++){
				row=sheet.getRow(j);	
				if(row!=null){
					System.out.println(row.getCell(0)+" "+row.getCell(1)+" "+row.getCell(2));
				}
				
			}
			System.out.println();
		}
		fileInputStream.close();
	} catch (Exception e) {		
		e.printStackTrace();
	}
	
    System.out.println("read excel ok!");	

	}

}
