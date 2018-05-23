package com.github.ldzzdl.easyexcel4j.writer.util;

import com.github.ldzzdl.easyexcel4j.annotation.Excel;
import com.github.ldzzdl.easyexcel4j.metadata.ExcelType;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Random;

/**
 * Excel的导出工具
 */
public class ExcelWriterUtil {

    /**
     * 将模型数组导出为Excel
     * @param list 模型数组
     * @param excelWriterContext Excel的导出上下文
     * @param <E> 泛型
     */
    public <E> void writeModelList2Excel(List<E> list, ExcelWriterContext excelWriterContext) throws IOException, IllegalAccessException, NoSuchMethodException, InvocationTargetException {
        ExcelType excelType = excelWriterContext.getExcelType();
        if(ExcelType.XLS.equals(excelType)){
            Workbook wb = new HSSFWorkbook();
            createWorkBook(list, wb, excelWriterContext);
            try  (OutputStream fileOut = new FileOutputStream(getFileName()+".xls")) {
                wb.write(fileOut);
            }
        }else if(ExcelType.XLSX.equals(excelType)){
            Workbook wb = new XSSFWorkbook();
            createWorkBook(list, wb, excelWriterContext);
            try (OutputStream fileOut = new FileOutputStream(getFileName() + ".xlsx")) {
                wb.write(fileOut);
            }
        }
    }

    /**
     * 根据模型数组产生工作簿的内容
     * @param list 模型数组
     * @param wb 工作簿
     * @param excelWriterContext Excel的导出上下文
     * @param <E> 泛型
     */
    private <E> void createWorkBook(List<E> list, Workbook wb, ExcelWriterContext excelWriterContext) throws IllegalAccessException, InvocationTargetException, NoSuchMethodException {
        Sheet sheet = wb.createSheet();
        if(list != null && list.size() >= 1){
            Field[] fields = list.get(0).getClass().getDeclaredFields();
            List<String> excelTitles = new ArrayList<>();
            boolean flag = false;
            for(Field field : fields){
                Excel excel = field.getAnnotation(Excel.class);
                if(excel == null){
                    excelTitles.add("");
                }else{
                    excelTitles.add(excel.excelTitle());
                    flag = true;
                }
            }
            int rowNumber = 0;
            if(flag){
                Row row = sheet.createRow(rowNumber);
                int cellNumber = 0;
                CellStyle titleCellStyle = getTitleCellStyle(wb, excelWriterContext);
                for (String excelTitle : excelTitles){
                    Cell cell = row.createCell(cellNumber);
                    cell.setCellStyle(titleCellStyle);
                    cell.setCellValue(excelTitle);
                    cellNumber ++;
                }
                rowNumber++;
            }

            for(E e : list){
                Row row = sheet.createRow(rowNumber);
                int cellNumber = 0;
                CellStyle dataCellStyle = getDataCellStyle(wb, excelWriterContext);
                for(Field field : fields){
                    Cell cell = row.createCell(cellNumber);
                    cell.setCellValue(BeanUtils.getProperty(e, field.getName()));
                    cell.setCellStyle(dataCellStyle);
                    cellNumber ++;
                }
                rowNumber ++;
            }
        }
    }

    /**
     * 获取显示标题的单元格样式
     * @param workbook 工作簿
     * @param excelWriterContext 导出Excel的上下文环境
     * @return 单元格样式
     */
    private CellStyle getTitleCellStyle(Workbook workbook, ExcelWriterContext excelWriterContext){
        CellStyle cellStyle = getCommonCellStyle(workbook);
        Font font = workbook.createFont();
        font.setFontHeightInPoints(excelWriterContext.getFontHeightInPoints());
        font.setFontName(excelWriterContext.getFontName());
        font.setBold(true);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        return cellStyle;
    }

    /**
     * 获取显示数据的单元格样式
     * @param workbook 工作簿
     * @param excelWriterContext 导出Excel的上下文环境
     * @return 单元格样式
     */
    private CellStyle getDataCellStyle(Workbook workbook, ExcelWriterContext excelWriterContext){
        CellStyle cellStyle = getCommonCellStyle(workbook);
        Font font = workbook.createFont();
        font.setFontHeightInPoints(excelWriterContext.getFontHeightInPoints());
        font.setFontName(excelWriterContext.getFontName());
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        return cellStyle;
    }

    /**
     * 获取Excel文件共同的样式
     * @param workbook 工作簿
     * @return 单元格样式
     */
    private CellStyle getCommonCellStyle(Workbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setWrapText(true);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
        return cellStyle;
    }

    /**
     * 随机生成文件名称
     * @return 文件名称
     */
    private static String getFileName(){
        StringBuffer name = new StringBuffer();
        Random random = new Random();
        for(int i = 0; i < 24; i++){
            name.append(((char)('a'+ random.nextInt(25))));
        }
        return name.toString();
    }
}
