package net.abadguy.test;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExeclWriteTestBigData {
    String PATH="C:\\Users\\yy\\Desktop\\POI-EasyExcel";

    @Test
    public void testWrite03BigData() throws IOException {
        long startTime = System.currentTimeMillis();
        //创建一个工作簿
        Workbook workbook=new HSSFWorkbook();
        //创建工作表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for(int rowNum=0;rowNum<65536;rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum=0; cellNum<10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "\\ExeclWriteTestBigData.xls");
        workbook.write(fileOutputStream);
        fileOutputStream.close();


        long useTime=System.currentTimeMillis()-startTime;
        System.out.println("一共使用时间 ： "+useTime);

    }


    @Test
    public void testWrite07BigData() throws IOException {
        long startTime = System.currentTimeMillis();
        //创建一个工作簿
        Workbook workbook=new XSSFWorkbook();
        //创建工作表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for(int rowNum=0;rowNum<65537;rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum=0; cellNum<10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "\\testWrite07BigData.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();


        long useTime=System.currentTimeMillis()-startTime;
        System.out.println("一共使用时间 ： "+useTime);

    }


    @Test
    public void testWrite07BigDataS() throws IOException {
        long startTime = System.currentTimeMillis();
        //创建一个工作簿
        Workbook workbook=new SXSSFWorkbook();
        //创建工作表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for(int rowNum=0;rowNum<65537;rowNum++){
            Row row = sheet.createRow(rowNum);
            for (int cellNum=0; cellNum<10;cellNum++){
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "\\testWrite07BigDataS.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        //清除临时文件
        ((SXSSFWorkbook)workbook).dispose();

        long useTime=System.currentTimeMillis()-startTime;
        System.out.println("一共使用时间 ： "+useTime);

    }






}
