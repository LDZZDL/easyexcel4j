package com.github.easyexcel4j.reader.resolver;

import com.github.easyexcel4j.metadata.ExcelType;
import com.github.easyexcel4j.reader.context.ReaderContext;
import com.github.easyexcel4j.reader.listener.*;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author LDZZDL
 * V07版本的Excel解析器
 */
public class ReaderResolverV07 extends ExcelReaderListenerManager{


    private class ReaderHandlerV07 implements XSSFSheetXMLHandler.SheetContentsHandler {
        private int currentRow = -1;
        private int currentCol = -1;


        @Override
        public void startRow(int rowNum) {
            handlerMissingRows(rowNum-currentRow-1);
            // Prepare for this row
            currentRow = rowNum;
            currentCol = -1;
        }

        private void handlerMissingRows(int number) {
            for (int i = 1; i <= number; i++) {
                readerContext.setBlankRow(true);
                readerContext.setCurrentRowNumber(currentRow + i);
                try {
                    notifyListener(datas, readerContext);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
            readerContext.setBlankRow(false);
        }

        @Override
        public void endRow(int rowNum)  {
            readerContext.setCurrentRowNumber(currentRow);
            readerContext.setLastColumnNumber(currentCol);
            notifyListener(datas, readerContext);
            datas.clear();
        }

        @Override
        public void cell(String cellReference, String formattedValue,
                         XSSFComment comment) {

            // gracefully handle missing CellRef here in a similar way as XSSFCell does
            if(cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }
            // Did we miss any cells?
            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i=0; i < missedCols; i++) {
                datas.add(null);
            }
            currentCol = thisCol;
            datas.add(formattedValue);
        }

        @Override
        public void headerFooter(String s, boolean b, String s1) {

        }
    }

    ///////////////////////////////////////

    private OPCPackage opcPackage;

    private ReaderContext readerContext;

    private List<String> datas;

    private void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings,
        XSSFSheetXMLHandler.SheetContentsHandler sheetHandler,
        InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = SAXHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch(ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    private void process() throws IOException, OpenXML4JException, SAXException {
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.opcPackage);
        XSSFReader xssfReader = new XSSFReader(this.opcPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int index = 0;
        while (sheetIterator.hasNext()) {
            try (InputStream stream = sheetIterator.next()) {
                String sheetName = sheetIterator.getSheetName();
                readerContext.setCurrentSheetIndex(index);
                readerContext.setSheetName(sheetName);
                processSheet(styles, strings, new ReaderHandlerV07(), stream);
            }
            ++index;
        }
    }

    private void init(ExcelReaderListener excelReaderListener, Class clazz){
        this.readerContext = new ReaderContext();
        this.readerContext.setExcelType(ExcelType.XLSX);
        this.datas = new ArrayList<>();
        registerListener(excelReaderListener);
        readerContext.setClazz(clazz);
    }

    public void process(String path, ExcelReaderListener excelReaderListener, Class clazz) throws OpenXML4JException, SAXException, IOException {
        OPCPackage opcPackage = OPCPackage.open(path, PackageAccess.READ);
        this.opcPackage = opcPackage;
        init(excelReaderListener, clazz);
        process();
    }

    public void process(File file, ExcelReaderListener excelReaderListener, Class clazz) throws OpenXML4JException, SAXException, IOException {
        OPCPackage opcPackage = OPCPackage.open(file, PackageAccess.READ);
        this.opcPackage = opcPackage;
        init(excelReaderListener, clazz);
        process();
    }

    public void process(InputStream fileInputStream, ExcelReaderListener excelReaderListener, Class clazz) throws IOException, OpenXML4JException, SAXException {
        OPCPackage opcPackage = OPCPackage.open(fileInputStream);
        this.opcPackage = opcPackage;
        init(excelReaderListener, clazz);
        process();
    }

}
