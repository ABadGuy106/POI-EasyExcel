package net.abadguy.test;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Date;

public class ExecReadTest {
    String PATH="C:\\Users\\yy\\Desktop\\POI-EasyExcel\\";


    @Test
    public void testRead03() throws IOException {
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "Ttest03.xls");

        //创建一个工作簿
        Workbook workbook = new HSSFWorkbook(inputStream);
        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到表格
        Cell cell = row.getCell(0);
        Cell cell2 = row.getCell(1);
        //得到表格内容
        System.out.println(cell.getNumericCellValue());
        System.out.println(cell2.getStringCellValue());
    }


    @Test
    public void testRead07() throws IOException {
        //获取文件流
        FileInputStream inputStream = new FileInputStream(PATH + "TestRead.xlsx");

        //创建一个工作簿
        Workbook workbook = new XSSFWorkbook(inputStream);
        //得到表
        Sheet sheet = workbook.getSheetAt(0);
        //得到行
        Row row = sheet.getRow(0);
        //得到表格
        Cell cell = row.getCell(0);
        Cell cell2 = row.getCell(1);
        //得到表格内容
        System.out.println(cell.getNumericCellValue());
        System.out.println(cell2.getStringCellValue());
    }


    @Test
    public void testCellType() throws IOException {
        FileInputStream inputStream = new FileInputStream(PATH + "basss.xlsx");

        Workbook workbook=new XSSFWorkbook(inputStream);

        Sheet sheet = workbook.getSheetAt(0);
        Row rowTitle = sheet.getRow(0);
        if(rowTitle!=null){
            //获取这一行有多少有内容的列
            int cellCount=rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                Cell cell = rowTitle.getCell(cellNum);
                if(cell!=null){
                    int type = cell.getCellType();
//                    System.out.println("-------type="+type+"----");
                    String stringCellValue = cell.getStringCellValue();
                    System.out.print(stringCellValue+"|");
                }
                //获取表中的内容
                int rows = sheet.getPhysicalNumberOfRows();
                for (int rowNum = 1; rowNum < rows; rowNum++) {
                    Row rowdata = sheet.getRow(rowNum);
                    //读取列
                    int cellCount2=rowTitle.getPhysicalNumberOfCells();
                    for (int cellNum2 = 0; cellNum2 < cellCount2; cellNum2++) {
                        Cell cell1 = rowdata.getCell(cellNum2);
                        //匹配数据类型
                        if(cell1!=null){
                            int cellType = cell1.getCellType();
                            String cellValue="";
                            switch (cellType){
                                case HSSFCell.CELL_TYPE_STRING: //字符串
                                    System.out.print("[String]");
                                    cellValue = cell1.getStringCellValue();
                                    break;
                                case HSSFCell.CELL_TYPE_NUMERIC: //数字类型
                                    System.out.print("[NUMERIC]");
                                    if(HSSFDateUtil.isCellDateFormatted(cell1)){
                                        System.out.println("日期");
                                        Date dateCellValue = cell1.getDateCellValue();
                                        cellValue = new DateTime(dateCellValue).toString("yyyy-MM-dd");
                                    }else {
                                        //不是日期格式防止字符串太长
                                        System.out.println("转换为字符输出");
                                        cell1.setCellType(HSSFCell.CELL_TYPE_STRING);
                                        cellValue = cell1.toString();
                                    }
                                    break;
                                case HSSFCell.CELL_TYPE_BLANK: //空
                                    System.out.println("内容为空");
                                    break;
                                case HSSFCell.CELL_TYPE_BOOLEAN: //布尔值
                                    System.out.println("Boolean");
                                    cellValue = String.valueOf(cell1.getBooleanCellValue());
                                    break;
                                case HSSFCell.CELL_TYPE_ERROR: //错误
                                    System.out.println("数据类型错误");
                                    break;
                            }
                            System.out.println(cellValue);
                        }
                    }
                }
            }
            inputStream.close();
        }
    }
}
