package com.sitech;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class FeeHandle {
	
	public static String printResult(String phoneNos, String remark) throws IOException{
		String result = 
        		" --处理步骤: " +
        		"--1.  先把号码记录插入到赠费删除记录表里面\n" + 
        		"insert into tf_b_task_delete  (task_id, serial_number, next_exe_time, task_code, start_date,"+
        		"end_date, exe_number, service_message, login_no, delete_reason)"+ 
				"select"+ 
				"k.task_id,"+ 
				"k.serial_number, "+ 
				"k.next_exe_time, "+ 
				"k.task_code,"+ 
				"k.start_date, "+ 
				"k.end_date, "+ 
				"k.exe_number,"+ 
				"k.service_message,"+ 
				"k.login_no, "+ 
				" \""+remark+"来自张甜甜\""+
				" from tf_b_task k where k.serial_number in("+
				""+phoneNos+");\n"+

			 	"-- 2.  删除原赠费表的记录\n"+
			 	"delete   from tf_b_task k where k.serial_number in("+
			 	""+phoneNos+");";
		File f = new File("D:\\777.txt");
		FileWriter fw = new FileWriter(f);
		fw.write(result);
		fw.close();
	 	return result;
	}
	public static void main(String[] args) throws FileNotFoundException, IOException {
		 //1.读取Excel文档对象
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream("D:\\222.xls"));
        //2.获取要解析的表格（第一个表格）
        HSSFSheet sheet = hssfWorkbook.getSheetAt(0);
        //获得最后一行的行号
        int lastRowNum = sheet.getLastRowNum();
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i <= lastRowNum; i++) {//遍历每一行
            //3.获得要解析的行
            HSSFRow row = sheet.getRow(i);
            //4. 设置单元格格式为字符串
            for (int j = 0; j < row.getLastCellNum(); j++) {
            	HSSFCell cell = row.getCell(j);
                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			}
            
            String stringCellValue0 = row.getCell(0).getStringCellValue();
            sb.append("'" +stringCellValue0+"',\n");
        }
        String result = printResult(sb.substring(0, sb.lastIndexOf(",")), "合肥");     
        System.out.println(result);
	}
}
