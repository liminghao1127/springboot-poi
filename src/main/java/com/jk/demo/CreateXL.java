/**
 * Copyright (C), 2019, 金科教育
 * FileName: CreateXL
 * Author:   zyl
 * Date:     2019/8/8 14:10
 * History:
 * zyl          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.jk.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * 〈一句话功能简述〉<br> 
 * 〈〉
 *
 * @author zyl
 * @create 2019/8/8
 * @since 1.0.0
 */
public class CreateXL {
/** Excel 文件要存放的位置，假定在D盘下*/
public static String outputFile="D:\\test1.xlsx";
public static void main(String argv[]){
try{
// 创建新的Excel 工作簿
 XSSFWorkbook workbook = new XSSFWorkbook();
// 在Excel工作簿中建一工作表，其名为缺省值
 // 如要新建一名为"效益指标"的工作表，其语句为：
 // HSSFSheet sheet = workbook.createSheet("sheet1");
 XSSFSheet sheet = workbook.createSheet();
 // 在索引0的位置创建行（第一行）
XSSFRow row = sheet.createRow((short)0);
//在索引0的位置创建单元格（第一列）
 XSSFCell cell = row.createCell( 0);
 // 定义单元格为字符串类型（Excel-设置单元格格式-数字-文本；不设置默认为“常规”，也可以设置成其他的，具体设置参考相关文档）
cell.setCellType(HSSFCell.CELL_TYPE_STRING);
// 在单元格中输入一些内容
cell.setCellValue("1902B最厉害！！！！其他的渣渣");
// 新建一输出文件流
FileOutputStream fOut = new FileOutputStream(outputFile);
// 把相应的Excel 工作簿存盘
workbook.write(fOut);
fOut.flush();
// 操作结束，关闭文件
fOut.close();
System.out.println("文件生成");
}catch(Exception e) {
System.out.println("已运行 xlCreate() : " + e );
}
}
}
