/**
 * Copyright (C), 2019, 金科教育
 * FileName: ReadXL
 * Author:   zyl
 * Date:     2019/8/8 14:18
 * History:
 * zyl          <time>          <version>          <desc>
 * 作者姓名           修改时间           版本号              描述
 */
package com.jk.demo;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

/**
 * 〈一句话功能简述〉<br> 
 * 〈〉
 *
 * @author zyl
 * @create 2019/8/8
 * @since 1.0.0
 */
public class ReadXL {
    /** Excel文件的存放位置。注意是正斜线*/
    public static String fileToBeRead="D:\\test1.xlsx";

    public static void main(String argv[]){
        try{
            // 创建对Excel工作簿文件的引用
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(fileToBeRead));
            // 创建对工作表的引用。
            // 本例是按名引用（让我们假定那张表有着缺省名"Sheet1"）
            XSSFSheet sheet = workbook.getSheet("Sheet0");
            // 也可用getSheetAt(int index)按索引引用，
            // 在Excel文档中，第一张工作表的缺省索引是0，
            // 其语句为：HSSFSheet sheet = workbook.getSheetAt(0);
            // 读取左上端单元
            XSSFRow row = sheet.getRow(0);
            XSSFCell cell = row.getCell(0);
            // 输出单元内容，cell.getStringCellValue()就是取所在单元的值
            System.out.println("左上端单元是： " + cell.getStringCellValue());
        }catch(Exception e) {
            System.out.println("已运行xlRead() : " + e );
        }
    }
}
