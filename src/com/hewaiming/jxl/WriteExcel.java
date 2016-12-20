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
			//����������
			WritableWorkbook workbook=Workbook.createWorkbook(new File("myexcel.xls"));
			 //�����µ�һҳ
			WritableSheet sheet=workbook.createSheet("��1",0);
			//�����ͷ
	        sheet.mergeCells(0, 0, 4, 0);//��Ӻϲ���Ԫ�񣬵�һ����������ʼ�У��ڶ�����������ʼ�У���������������ֹ�У����ĸ���������ֹ��
	        WritableFont bold = new WritableFont(WritableFont.ARIAL,10,WritableFont.BOLD);//������������ͺ�����ʾ,����ΪArial,�ֺŴ�СΪ10,���ú�����ʾ
	        WritableCellFormat titleFormate = new WritableCellFormat(bold);//����һ����Ԫ����ʽ���ƶ���
	        titleFormate.setAlignment(jxl.format.Alignment.CENTRE);//��Ԫ���е�����ˮƽ�������
	        titleFormate.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);//��Ԫ������ݴ�ֱ�������
	        Label title = new Label(0,0,"JExcelApi֧������������ϸ˵��",titleFormate);
	        sheet.setRowView(0, 600, false);//���õ�һ�еĸ߶�
	        sheet.addCell(title);
			
	        Integer BeginRow=1;
	        WritableFont color = new WritableFont(WritableFont.ARIAL);//ѡ������
	        color.setColour(Colour.BLUE);//����������ɫΪ���ɫ
	        WritableCellFormat colorFormat = new WritableCellFormat(color);
			//������������
			Label school=new Label(0,0+BeginRow,"ѧУ",colorFormat);
			sheet.addCell(school);
			Label professional=new Label(1, 0+BeginRow, "רҵ",colorFormat);
			sheet.addCell(professional);
			Label Compet=new Label(2, 0+BeginRow, "������",colorFormat);
			sheet.addCell(Compet);
			Label grade=new Label(3, 0+BeginRow, "�÷�",colorFormat);
			sheet.addCell(grade);
			Label ddate=new Label(4, 0+BeginRow, "����",colorFormat);
			sheet.addCell(ddate);
			//��������
			Label qinghua = new Label(0,1+BeginRow,"�廪��ѧ");
	        sheet.addCell(qinghua);
	        Label jisuanji = new Label(1,1+BeginRow,"�����רҵ");
	        sheet.addCell(jisuanji);
	        Label gao = new Label(2,1+BeginRow,"��");
	        sheet.addCell(gao);	        
	        jxl.write.Number myGrade=new jxl.write.Number(3, 1+BeginRow, 89.5);
			sheet.addCell(myGrade);
			Calendar c=Calendar.getInstance();			
			Date myDate=c.getTime();
			WritableCellFormat cf1 = new WritableCellFormat(DateFormats.FORMAT1);
	        DateTime dt = new DateTime(4,1+BeginRow,myDate,cf1);
	        sheet.addCell(dt);
			        
	        Label beida = new Label(0,2+BeginRow,"������ѧ");
	        sheet.addCell(beida);
	        Label falv = new Label(1,2+BeginRow,"����רҵ");
	        sheet.addCell(falv);
	        Label zhong = new Label(2,2+BeginRow,"��");
	        sheet.addCell(zhong);
	        jxl.write.Number myGrade2=new jxl.write.Number(3, 2+BeginRow, 82.9);
			sheet.addCell(myGrade2);
			//Calendar c=Calendar.getInstance();			
			Date myDate2=c.getTime();
			//WritableCellFormat cf1 = new WritableCellFormat(DateFormats.FORMAT1);
	        DateTime dt2 = new DateTime(4,2+BeginRow,myDate2,cf1);
	        sheet.addCell(dt2);
	        
	        
	        Label ligong = new Label(0,3+BeginRow,"��������ѧ");
	        sheet.addCell(ligong);
	        Label hangkong = new Label(1,3+BeginRow,"����רҵ");
	        sheet.addCell(hangkong);
	        Label di = new Label(2,3+BeginRow,"��");
	        sheet.addCell(di);
	        //�Ѵ���������д�뵽����ļ��У����ر��ļ�
	        workbook.write();
	        workbook.close();	
	        System.out.println("���ӱ�������ɣ�");
		} catch (Exception e) {			
			e.printStackTrace();
		}

	}

}
