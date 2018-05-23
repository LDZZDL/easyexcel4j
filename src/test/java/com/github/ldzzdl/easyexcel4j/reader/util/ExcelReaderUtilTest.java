package com.github.ldzzdl.easyexcel4j.reader.util;

import com.github.ldzzdl.easyexcel4j.metadata.ExcelType;
import com.github.ldzzdl.easyexcel4j.reader.resolver.ReaderResolverV07;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;
import org.xml.sax.SAXException;

import java.io.*;
import java.util.Date;
import java.util.List;
import java.util.Random;

public class ExcelReaderUtilTest {

    @Test
    public void createExcel() throws IOException {
        SXSSFWorkbook wb = null;
        FileOutputStream out = null;
        try{
            Random random = new Random();
            wb = new SXSSFWorkbook(100);
            Sheet sh = wb.createSheet();
            for(int rownum = 0; rownum < 100000; rownum++){
                Row row = sh.createRow(rownum);
                for(int cellnum = 0; cellnum < 20; cellnum++){
                    Cell cell = row.createCell(cellnum);
                    StringBuilder stringBuilder = new StringBuilder();
                    for(int i = 0; i < 32; i++){
                        stringBuilder.append(((char)('a'+ random.nextInt(25))));
                    }
                    cell.setCellValue(stringBuilder.toString());
                }
                for(int cellnum = 20; cellnum < 23; cellnum ++){
                    Cell cell = row.createCell(cellnum);
                    cell.setCellValue(cellnum*100);
                }
                for(int cellnum = 23; cellnum < 25; cellnum ++){
                    Cell cell = row.createCell(cellnum);
                    cell.setCellValue(new Date());
                }
            }
            out = new FileOutputStream("src/test/java/doc/large.xlsx");
            wb.write(out);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if(out != null)
                out.close();
            if(wb != null)
                wb.dispose();
        }
    }

    @Test
    public void test1() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/1.xls");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<String> list =
                excelReaderUtil.readExcel2List(inputStream, ExcelType.XLS,1,1,1);
        for(String string : list){
            System.out.println(string);
        }
    }

    @Test
    public void test2() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/1.xlsx");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<String> list =
                excelReaderUtil.readExcel2List(inputStream, ExcelType.XLSX,1,1,1);
        for(String string : list){
            System.out.println(string);
        }
    }

    @Test
    public void test3() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/1.xls");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<String> list =
                excelReaderUtil.readExcel2List(inputStream, ExcelType.XLS,1,1,0);
        for(String string : list){
            System.out.println(string);
        }
    }

    @Test
    public void test4() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/1.xlsx");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<String> list =
                excelReaderUtil.readExcel2List(inputStream, ExcelType.XLSX,3,1,2);
        for(String string : list){
            System.out.println(string);
        }
    }

    @Test
    public void test5() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/2.xls");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<TestModel> list = excelReaderUtil.readExcel2ModelListByOrder(inputStream, ExcelType.XLS, TestModel.class,2,2,0);
        for(TestModel testModel : list){
            System.out.println(testModel);
        }
    }

    @Test
    public void test6() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/2.xlsx");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<TestModel> list = excelReaderUtil.readExcel2ModelListByOrder(inputStream, ExcelType.XLSX, TestModel.class,2,1,0);
        for(TestModel testModel : list){
            System.out.println(testModel);
        }
    }

    @Test
    public void test7() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/3.xls");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<TestModel> list = excelReaderUtil.readExcel2ModelListByTitle(inputStream, ExcelType.XLS, TestModel.class,1,0);
        for(TestModel testModel : list){
            System.out.println(testModel);
        }
    }

    @Test
    public void test8() throws OpenXML4JException, SAXException, IOException {
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/3.xlsx");
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<TestModel> list = excelReaderUtil.readExcel2ModelListByTitle(inputStream, ExcelType.XLSX, TestModel.class,1,0);
        for(TestModel testModel : list){
            System.out.println(testModel);
        }
    }

    @Test
    public void test9() throws OpenXML4JException, SAXException, IOException {
        ReaderResolverV07 readerResolverV07 = new ReaderResolverV07();
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/large.xlsx");
        readerResolverV07.process(inputStream, new ReadLargeSheet(), null);
    }

//    @Test
//    public void test10() throws IOException {
//        ReaderResolverV03 readerResolverV03 = new ReaderResolverV03();
//        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/large.xls");
//        readerResolverV03.process(inputStream, new ReadLargeSheet(), null);
//    }





}