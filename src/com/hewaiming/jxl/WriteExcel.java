package com.hewaiming.jxl;

import java.io.File;
import java.io.IOException;
import java.util.Calendar;
import java.util.Date;

import jxl.Workbook;
import jxl.format.Colour;
import jxl.write.DateFormats;
import jxl.write.DateTime;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class WriteExcel {

	public static void main(String[] args) {		
		try {
			//创建工作薄
			WritableWorkbook workbook=Workbook.createWorkbook(new File("myexcel.xls"));
			 //创建新的一页
			WritableSheet sheet=workbook.createSheet("表单1",0);
			//构造表头
	        sheet.mergeCells(0, 0, 4, 0);//添加合并单元格，第一个参数是起始列，第二个参数是起始行，第三个参数是终止列，第四个参数是终止行
	        WritableFont bold = new WritableFont(WritableFont.ARIAL,10,WritableFont.BOLD);//设置字体种类和黑体显示,字体为Arial,字号大小为10,采用黑体显示
	        WritableCellFormat titleFormate = new WritableCellFormat(bold);//生成一个单元格样式控制对象
	        titleFormate.setAlignment(jxl.format.Alignment.CENTRE);//单元格中的内容水平方向居中
	        titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);//单元格的内容垂直方向居中
	        Label title = new Label(0,0,"JExcelApi支持数据类型详细说明",titleFormate);
	        sheet.setRowView(0, 600, false);//设置第一行的高度
	        sheet.addCell(title);
			
	        Integer BeginRow=1;
	        WritableFont color = new WritableFont(WritableFont.ARIAL);//选择字体
	        color.setColour(Colour.BLUE);//设置字体颜色为金黄色
	        WritableCellFormat colorFormat = new WritableCellFormat(color);
			//增加数据项名
			Label school=new Label(0,0+BeginRow,"学校",colorFormat);
			sheet.addCell(school);
			Label professional=new Label(1, 0+BeginRow, "专业",colorFormat);
			sheet.addCell(professional);
			Label Compet=new Label(2, 0+BeginRow, "竞争力",colorFormat);
			sheet.addCell(Compet);
			Label grade=new Label(3, 0+BeginRow, "得分",colorFormat);
			sheet.addCell(grade);
			Label ddate=new Label(4, 0+BeginRow, "日期",colorFormat);
			sheet.addCell(ddate);
			//增加数据
			Label qinghua = new Label(0,1+BeginRow,"清华大学");
	        sheet.addCell(qinghua);
	        Label jisuanji = new Label(1,1+BeginRow,"计算机专业");
	        sheet.addCell(jisuanji);
	        Label gao = new Label(2,1+BeginRow,"高");
	        sheet.addCell(gao);	        
	        jxl.write.Number myGrade=new jxl.write.Number(3, 1+BeginRow, 89.5);
			sheet.addCell(myGrade);
			Calendar c=Calendar.getInstance();			
			Date myDate=c.getTime();
			WritableCellFormat cf1 = new WritableCellFormat(DateFormats.FORMAT1);
	        DateTime dt = new DateTime(4,1+BeginRow,myDate,cf1);
	        sheet.addCell(dt);
			        
	        Label beida = new Label(0,2+BeginRow,"北京大学");
	        sheet.addCell(beida);
	        Label falv = new Label(1,2+BeginRow,"法律专业");
	        sheet.addCell(falv);
	        Label zhong = new Label(2,2+BeginRow,"中");
	        sheet.addCell(zhong);
	        jxl.write.Number myGrade2=new jxl.write.Number(3, 2+BeginRow, 82.9);
			sheet.addCell(myGrade2);
			//Calendar c=Calendar.getInstance();			
			Date myDate2=c.getTime();
			//WritableCellFormat cf1 = new WritableCellFormat(DateFormats.FORMAT1);
	        DateTime dt2 = new DateTime(4,2+BeginRow,myDate2,cf1);
	        sheet.addCell(dt2);
	        
	        
	        Label ligong = new Label(0,3+BeginRow,"北京理工大学");
	        sheet.addCell(ligong);
	        Label hangkong = new Label(1,3+BeginRow,"航空专业");
	        sheet.addCell(hangkong);
	        Label di = new Label(2,3+BeginRow,"低");
	        sheet.addCell(di);
	        //把创建的内容写入到输出文件中，并关闭文件
	        workbook.write();
	        workbook.close();	
	        System.out.println("电子表格已生成！");
		} catch (Exception e) {			
			e.printStackTrace();
		}

	}

}
