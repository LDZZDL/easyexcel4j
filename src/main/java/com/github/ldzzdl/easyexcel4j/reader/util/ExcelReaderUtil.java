package com.github.ldzzdl.easyexcel4j.reader.util;

import com.github.ldzzdl.easyexcel4j.metadata.ExcelType;
import com.github.ldzzdl.easyexcel4j.reader.listener.AutoModelExcelReader;
import com.github.ldzzdl.easyexcel4j.reader.listener.DefaultExcelReader;
import com.github.ldzzdl.easyexcel4j.reader.listener.DefaultModelExcelReader;
import com.github.ldzzdl.easyexcel4j.reader.resolver.ReaderResolverV03;
import com.github.ldzzdl.easyexcel4j.reader.resolver.ReaderResolverV07;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * @author LDZZDL
 * 导入的工具类
 */
public class ExcelReaderUtil {

    /**
     * 读取Excel封装成List数组
     * @param path excel的路径
     * @param excelType excel的类型
     * @param startRowNumber 开始读取Excel的行数
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @return List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public List<String> readExcel2List(String path, ExcelType excelType, int startRowNumber, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        FileInputStream fileInputStream = new FileInputStream(new File(path));
        return readExcel2List(fileInputStream, excelType, startRowNumber, startSheetNumber, endSheetNumber);
    }

    /**
     * 读取Excel封装成List数组
     * @param file excel的文件
     * @param excelType excel的类型
     * @param startRowNumber 开始读取Excel的行数
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @return List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public List<String> readExcel2List(File file, ExcelType excelType, int startRowNumber, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        FileInputStream fileInputStream = new FileInputStream(file);
        return readExcel2List(fileInputStream, excelType, startRowNumber, startSheetNumber, endSheetNumber);
    }

    /**
     * 读取Excel封装成List数组
     * @param fileInputStream excel文件的IO流
     * @param excelType excel的类型
     * @param startRowNumber 开始读取Excel的行数
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @return List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public List<String> readExcel2List(InputStream fileInputStream, ExcelType excelType, int startRowNumber, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        DefaultExcelReader defaultExcelReader = new DefaultExcelReader();
        defaultExcelReader.init(startRowNumber, startSheetNumber, endSheetNumber);
        if(ExcelType.XLS.equals(excelType)){
            ReaderResolverV03 readerResolverV03 = new ReaderResolverV03();
            readerResolverV03.process(fileInputStream,defaultExcelReader, null);
        }else if(ExcelType.XLSX.equals(excelType)){
            ReaderResolverV07 readerResolverV07 = new ReaderResolverV07();
            readerResolverV07.process(fileInputStream, defaultExcelReader, null);
        }
        return defaultExcelReader.getDatas();
    }

    /**
     * 读取Excel文件，根据JavaBean的注解的excelOrder，将每一行数据的封装JavaBean
     * @param path excel的路径
     * @param excelType excel的类型
     * @param clazz JavaBean的class
     * @param startRowNumber 开始读取Excel的行数
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @param <E> 泛型
     * @return 模型的List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public <E> List<E> readExcel2ModelListByOrder(String path, ExcelType excelType, Class clazz, int startRowNumber, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        FileInputStream fileInputStream = new FileInputStream(new File(path));
        return readExcel2ModelListByOrder(fileInputStream, excelType, clazz, startRowNumber, startSheetNumber, endSheetNumber);
    }

    /**
     * 读取Excel文件，根据JavaBean的注解的excelOrder，将每一行的数据封装JavaBean
     * @param file excel文件
     * @param excelType excel的类型
     * @param clazz JavaBean的class
     * @param startRowNumber 开始读取Excel的行数
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @param <E> 泛型
     * @return 模型的List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public <E> List<E> readExcel2ModelListByOrder(File file, ExcelType excelType, Class clazz, int startRowNumber, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        FileInputStream fileInputStream = new FileInputStream(file);
        return readExcel2ModelListByOrder(fileInputStream, excelType, clazz, startRowNumber, startSheetNumber, endSheetNumber);
    }

    /**
     * 读取Excel文件，根据JavaBean的注解的excelOrder，将每一行的数据封装JavaBean
     * @param fileInputStream excel的IO流
     * @param excelType excel的类型
     * @param clazz JavaBean的class
     * @param startRowNumber 开始读取Excel的行数
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @param <E> 泛型
     * @return 模型的List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public <E> List<E> readExcel2ModelListByOrder(InputStream fileInputStream, ExcelType excelType, Class clazz, int startRowNumber, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        DefaultModelExcelReader defaultModelExcelReader = new DefaultModelExcelReader();
        defaultModelExcelReader.init(startRowNumber, startSheetNumber, endSheetNumber);
        if(ExcelType.XLS.equals(excelType)){
            ReaderResolverV03 readerResolverV03 = new ReaderResolverV03();
            readerResolverV03.process(fileInputStream, defaultModelExcelReader, clazz);
        }else if(ExcelType.XLSX.equals(excelType)){
            ReaderResolverV07 readerResolverV07 = new ReaderResolverV07();
            readerResolverV07.process(fileInputStream,defaultModelExcelReader, clazz);
        }
        return defaultModelExcelReader.getDatas();
    }

    /**
     * 读取Excel文件，根据JavaBean的注解的excelTitle，将每一行的数据封装JavaBean
     * @param path excel的路径
     * @param excelType excel的类型
     * @param clazz JavaBean的class
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @param <E> 泛型
     * @return 模型的List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public <E> List<E> readExcel2ModelListByTitle(String path, ExcelType excelType, Class clazz, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        FileInputStream fileInputStream = new FileInputStream(new File(path));
        return readExcel2ModelListByTitle(fileInputStream, excelType, clazz, startSheetNumber, endSheetNumber);
    }

    /**
     * 读取Excel文件，根据JavaBean的注解的excelTitle，将每一行的数据封装JavaBean
     * @param file excel的文件
     * @param excelType excel的类型
     * @param clazz JavaBean的class
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @param <E> 泛型
     * @return 模型的List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public <E> List<E> readExcel2ModelListByTitle(File file, ExcelType excelType, Class clazz, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        FileInputStream fileInputStream = new FileInputStream(file);
        return readExcel2ModelListByTitle(fileInputStream, excelType, clazz, startSheetNumber, endSheetNumber);
    }

    /**
     * 读取Excel文件，根据JavaBean的注解的excelTitle，将每一行的数据封装JavaBean
     * @param fileInputStream excel的IO流
     * @param excelType excel的类型
     * @param clazz JavaBean的class
     * @param startSheetNumber 开始读取Sheet的序号（从1开始）
     * @param endSheetNumber 结束读取Sheet的序号（0代表读到末尾）
     * @param <E> 泛型
     * @return 模型的List数组
     * @throws IOException
     * @throws OpenXML4JException
     * @throws SAXException
     */
    public <E> List<E> readExcel2ModelListByTitle(InputStream fileInputStream, ExcelType excelType, Class clazz, int startSheetNumber, int endSheetNumber) throws IOException, OpenXML4JException, SAXException {
        AutoModelExcelReader autoModelExcelReader = new AutoModelExcelReader();
        autoModelExcelReader.init(startSheetNumber, endSheetNumber);
        if(ExcelType.XLS.equals(excelType)){
            ReaderResolverV03 readerResolverV03 = new ReaderResolverV03();
            readerResolverV03.process(fileInputStream, autoModelExcelReader, clazz);
            return autoModelExcelReader.getResult();
        }else if(ExcelType.XLSX.equals(excelType)){
            ReaderResolverV07 readerResolverV07 = new ReaderResolverV07();
            readerResolverV07.process(fileInputStream, autoModelExcelReader, clazz);
            return autoModelExcelReader.getResult();
        }
        return null;
    }

}
