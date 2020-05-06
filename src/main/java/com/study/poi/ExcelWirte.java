package com.study.poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelWirte {

    /**
     * Excel的写入有分03版和07版
     *  区别: 03版写入是先写到缓存,在从缓存写到临时文件 最后写入   速度相对快
     *       03版最多写入65536行数据,07可以写入无限行数
     *       07版是一次性全部写入 速度相对慢
     */

    final  static String PATH= "D:\\dev_workhome\\studyPoi\\test\\";

    /**
     * 03版的Excel测试
     * @throws Exception
     */
    @Test
    public void ExcelWrite03() throws  Exception{
        //1.创建工作簿
        Workbook workbook = new HSSFWorkbook();//HSSFWorkbook为03版的Excel
        //2.创建工作页
        Sheet sheet = workbook.createSheet("03版测试Excel.xls");
        //3.创建行
        Row row = sheet.createRow(0);
        //4.创建列
        Cell cell = row.createCell(0);
        //那么这里就可以确定为 第一行的第一列了
        //5.设置值
        cell.setCellValue("0_0");
        //生成Excel
        try ( FileOutputStream outputStream = new FileOutputStream(PATH+"03版测试Excel.xls")){
            workbook.write(outputStream);
        }
    }

    /**
     * 07版的Excel测试
     * @throws Exception
     */
    @Test
    public void ExcelWrite07() throws  Exception{
        //1.创建工作簿
        Workbook workbook = new XSSFWorkbook();//HSSFWorkbook为03版的Excel
        //2.创建工作页
        Sheet sheet = workbook.createSheet("03版测试Excel.xls");
        //3.创建行
        Row row = sheet.createRow(0);
        //4.创建列
        Cell cell = row.createCell(0);
        //那么这里就可以确定为 第一行的第一列了
        //5.设置值
        cell.setCellValue("0_0");
        //生成Excel
        try ( FileOutputStream outputStream = new FileOutputStream(PATH+"03版测试Excel.xlsx")){
            workbook.write(outputStream);
        }
    }

    /**
     * 03版大数据量的插入
     */
    @Test
    public void ExcelWriteBigData03() throws  Exception{
        //1.创建工作簿
        Workbook workbook = new HSSFWorkbook();//HSSFWorkbook为03版的Excel
        //2.创建工作页
        Sheet sheet = workbook.createSheet("03版测试Excel.xls");

        for(int i=0;i<65536;i++){
            Row row = sheet.createRow(i);
            for(int j=0;j<10;j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }
        try ( FileOutputStream outputStream = new FileOutputStream(PATH+"03大数据版测试Excel.xls")){
            workbook.write(outputStream);
        }

    }

    /**
     * 07版大数据量的插入
     */
    @Test
    public void ExcelWriteBigData07() throws  Exception{
        //1.创建工作簿
        Workbook workbook = new SXSSFWorkbook();//HSSFWorkbook为03版的Excel
        //2.创建工作页
        Sheet sheet = workbook.createSheet("07版测试Excel.xls");

        for(int i=0;i<100000;i++){
            Row row = sheet.createRow(i);
            for(int j=0;j<10;j++){
                Cell cell = row.createCell(j);
                cell.setCellValue(j);
            }
        }
        try ( FileOutputStream outputStream = new FileOutputStream(PATH+"07大数据版测试Excel.xlsx")){
            workbook.write(outputStream);
        }

    }

    /**
     * 读取Excel数据 03版的Excel和07版的Excel是基本相同的
     * 不同的只有创建工作簿的时候获取的类是不同的
     *  03版的为: HSSFWorkbook
     *  07版的为:XSSFWorkbook 或者 SXSSFWorkbook
     */
    @Test
    public void ExcelRead() throws Exception {
        //1.读取数据存放的流
        try(FileInputStream in = new FileInputStream(PATH+"07大数据版测试Excel.xlsx");) {
            Workbook workbook = new XSSFWorkbook(in);
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(0);
            Cell cell0 = row.getCell(0);
            Cell cell1= row.getCell(1);
            System.out.println(cell0);
            System.out.println(cell1);
        }

    }

    /**
     * 判断Excel的不同类型,打印出不同类型之后输出
     * @throws Exception
     */
    @Test
    public void ExcelTypeRead() {
        try (FileInputStream in = new FileInputStream(PATH + "07大数据版测试Excel.xls");) {
            List<String> sheetList = new ArrayList<>();
            Workbook workbook=new XSSFWorkbook(in);
            Sheet sheet = workbook.getSheetAt(0);
            Row rowTitle = sheet.getRow(0);
            if(rowTitle!=null) {
                //获取当前行下不为空的总列数
                int cellCount = rowTitle.getPhysicalNumberOfCells();
                for (int i = 0; i < cellCount; i++) {
                    //获取每一列
                    Cell cell = rowTitle.getCell(i);
                    if(null!=cell){
                        //获取数据类型
                        CellType cellType = cell.getCellType();
                        String cellValue = cell.getStringCellValue();
                        System.out.print(cellValue+"|"+cellType);
                    }
                }
            }
            //获取总行数
            int rowCount = sheet.getPhysicalNumberOfRows();
            //第0行为标题行 不摄入统计
            for(int row =1;row<rowCount;row++){
                Row rowData = sheet.getRow(row);
                if(rowData!=null){
                    int cellCount = rowData.getPhysicalNumberOfCells();
                    for(int cell =0 ;cell<cellCount;cell++){
                        Cell cellData = rowData.getCell(cell);
                        if(cellData!=null){
                            //获得一个枚举类
                            CellType cellType = cellData.getCellType();
                            switch (cellType){
                                case _NONE:
                                    break;
                                case BLANK:
                                    break;
                                case ERROR:
                                    break;
                                case STRING:
                                    sheetList.add(cellData.getStringCellValue());
                                    break;
                                case BOOLEAN:
                                    sheetList.add(String.valueOf(cellData.getBooleanCellValue()));
                                    break;
                                case FORMULA:
                                    //代表是一个Excel公式
                                    sheetList.add(cellData.getCellFormula());
                                    break;
                                case NUMERIC:
                                    if(DateUtil.isCellDateFormatted(cellData)){
                                        sheetList.add( new DateTime(cellData.getDateCellValue()).toString("yyyy-MM-dd"));
                                        break;
                                    }else{
                                        sheetList.add(String.valueOf(cellData.getNumericCellValue()));
                                        break;
                                    }
                            }
                        }
                    }
                }
            }
            System.out.println("数据"+sheetList);
        }catch (Exception e){
            e.getMessage();
        }
    }
}
