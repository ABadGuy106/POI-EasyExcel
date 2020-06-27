package net.abadguy.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {

    String PATH="C:\\Users\\yy\\Desktop\\POI-EasyExcel";

    @Test
    public void testWrite03() throws IOException {
        //创建工作簿 - 03版本
        Workbook workbook=new HSSFWorkbook();
        //创建一个工作表
        Sheet sheet=workbook.createSheet("创建测试统计表");
        //创建一个行
        Row row1=sheet.createRow(0);
        //创建一个单元格,第一行第一个单元格，坐标(1,1)
        Cell cell = row1.createCell(0);
        //设置单元格内容
        cell.setCellValue("测试第一个单元格(0,0)");
        //创建二个单元格,第一行第二个单元格，坐标(1,2)
        Cell cell2 = row1.createCell(2);
        //设置单元格内容
        cell2.setCellValue(100);

        //创建第二行
        Row row2=sheet.createRow(1);
        //创建单元格,第一行第一个单元格，坐标(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue(true);
        //创建单元格,第一行第一个单元格，坐标(2,2)
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(23.33);

        Cell cell23 = row2.createCell(3);
        cell23.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //生成一张表(I/O流)  03版就是使用xls版本
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/test.xls");

        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("文件生成完毕");
    }



    @Test
    public void testWrite07() throws IOException {
        //创建工作簿 - 07版本
        Workbook workbook=new XSSFWorkbook();
        //创建一个工作表
        Sheet sheet=workbook.createSheet("创建测试统计表");
        //创建一个行
        Row row1=sheet.createRow(0);
        //创建一个单元格,第一行第一个单元格，坐标(1,1)
        Cell cell = row1.createCell(0);
        //设置单元格内容
        cell.setCellValue("测试第一个单元格(0,0)");
        //创建二个单元格,第一行第二个单元格，坐标(1,2)
        Cell cell2 = row1.createCell(2);
        //设置单元格内容
        cell2.setCellValue(100);

        //创建第二行
        Row row2=sheet.createRow(1);
        //创建单元格,第一行第一个单元格，坐标(2,1)
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue(true);
        //创建单元格,第一行第一个单元格，坐标(2,2)
        Cell cell22 = row2.createCell(1);
        cell22.setCellValue(23.33);

        Cell cell23 = row2.createCell(3);
        cell23.setCellValue(new DateTime().toString("yyyy-MM-dd HH:mm:ss"));

        //生成一张表(I/O流)  07版就是使用xlsx版本
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "/test.xlsx");

        workbook.write(fileOutputStream);
        //关闭流
        fileOutputStream.close();
        System.out.println("文件生成完毕");
    }
}
