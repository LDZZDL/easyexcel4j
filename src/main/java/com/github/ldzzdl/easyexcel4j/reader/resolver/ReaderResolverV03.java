package com.github.ldzzdl.easyexcel4j.reader.resolver;

import com.github.ldzzdl.easyexcel4j.metadata.ExcelType;
import com.github.ldzzdl.easyexcel4j.reader.context.ReaderContext;
import com.github.ldzzdl.easyexcel4j.reader.listener.ExcelReaderListener;
import com.github.ldzzdl.easyexcel4j.reader.listener.ExcelReaderListenerManager;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author LDZZDL
 * V03版本的Excel的解析器
 */
public class ReaderResolverV03 extends ExcelReaderListenerManager implements HSSFListener{

    /** POI File System */
    private POIFSFileSystem fs;

    private int lastRowNumber;
    private int lastColumnNumber;

    /** Should we output the formula, or the value it has? */
    private boolean outputFormulaValues = true;

    /** For parsing Formulas */
    private EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    /** Records we pick up as we process */
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /** So we known which sheet we're on */
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBSRs;
    private List<BoundSheetRecord> boundSheetRecords;

    /** For handling formulas with string results */
    private int nextRow;
    private int nextColumn;
    private boolean outputNextStringRecord;

    /** Context for reading excel03*/
    private ReaderContext readerContext;
    /** String ArrayList for each row*/
    private List<String> rowDatas;

    @Override
    public void processRecord(Record record) {
        int thisRow = -1;
        int thisColumn = -1;
        String thisStr = "";

        switch (record.getSid())
        {
            case BoundSheetRecord.sid:

                BoundSheetRecord boundSheetRecord = (BoundSheetRecord) record;
                boundSheetRecords.add(boundSheetRecord);
                readerContext.setSheetName(boundSheetRecord.getSheetname());
                break;
            case BOFRecord.sid:
                BOFRecord bofRecord = (BOFRecord)record;
                if(bofRecord.getType() == BOFRecord.TYPE_WORKSHEET) {
                    // Create sub workbook if required
                    if(workbookBuildingListener != null && stubWorkbook == null) {
                        stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                    }
                    //  Works by ordering the BSRs by the location of
                    //  their BOFRecords, and then knowing that we
                    //  process BOFRecords in byte offset order
                    sheetIndex++;
                    readerContext.setCurrentSheetIndex(sheetIndex);
                    if(orderedBSRs == null) {
                        orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                    }
                }
                break;
            case SSTRecord.sid:
                sstRecord = (SSTRecord) record;
                break;
            case BlankRecord.sid:
                BlankRecord blankRecord = (BlankRecord) record;
                thisRow = blankRecord.getRow();
                thisColumn = blankRecord.getColumn();
                thisStr = null;
                break;
            case BoolErrRecord.sid:
                BoolErrRecord boolErrRecord = (BoolErrRecord) record;
                thisRow = boolErrRecord.getRow();
                thisColumn = boolErrRecord.getColumn();
                Boolean flag = boolErrRecord.getBooleanValue();
                if(flag){
                    thisStr = "TRUE";
                }else{
                    thisStr = "FALSE";
                }
                break;
            case FormulaRecord.sid:
                FormulaRecord formulaRecord = (FormulaRecord) record;

                thisRow = formulaRecord.getRow();
                thisColumn = formulaRecord.getColumn();
                //output formula result
                if(outputFormulaValues) {
                    if(Double.isNaN(formulaRecord.getValue())) {
                        // Formula result is a string
                        // This is stored in the next record
                        outputNextStringRecord = true;
                        nextColumn = formulaRecord.getColumn();
                        nextRow = formulaRecord.getRow();
                    } else {
                        // Formula result is number
                        thisStr = String.valueOf(formulaRecord.getValue());
                    }
                }
                //output formula expression
                else {
                    thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, formulaRecord.getParsedExpression());
                }
                break;
            // String for formula
            case StringRecord.sid:
                if(outputNextStringRecord) {
                    // String for formula
                    StringRecord stringRecord = (StringRecord)record;
                    thisRow = nextRow;
                    thisColumn = nextColumn;
                    thisStr = stringRecord.getString();
                }
                break;
            case LabelRecord.sid:
                LabelRecord labelRecord = (LabelRecord) record;
                thisRow = labelRecord.getRow();
                thisColumn = labelRecord.getColumn();
                thisStr = labelRecord.getValue();
                break;
            // 单元格中的字符串类型
            case LabelSSTRecord.sid:
                LabelSSTRecord labelSSTRecord = (LabelSSTRecord) record;
                thisRow = labelSSTRecord.getRow();
                thisColumn = labelSSTRecord.getColumn();
                if(sstRecord == null) {
                    throw new RuntimeException("No SST Record, can't identify string");
                } else {
                    thisStr = sstRecord.getString(labelSSTRecord.getSSTIndex()).toString();
                }
                break;
            case NoteRecord.sid:
                NoteRecord noteRecord = (NoteRecord) record;
                thisRow = noteRecord.getRow();
                thisColumn = noteRecord.getColumn();
                thisStr = "(TODO)";
                break;
            case NumberRecord.sid:
                NumberRecord numberRecord = (NumberRecord) record;
                thisRow = numberRecord.getRow();
                thisColumn = numberRecord.getColumn();
                thisStr = String.valueOf(numberRecord.getValue());
                break;
            case RKRecord.sid:
                RKRecord pkRecord = (RKRecord) record;
                thisRow = pkRecord.getRow();
                thisColumn = pkRecord.getColumn();
                thisStr = "(TODO)";
                break;
            default:
                break;
        }
        // Handle new row
        if(thisRow != -1 && thisRow != lastRowNumber) {
            lastColumnNumber = -1;
        }
        // Handle missing column
        if(record instanceof MissingCellDummyRecord) {
            MissingCellDummyRecord mc = (MissingCellDummyRecord)record;
            thisRow = mc.getRow();
            thisColumn = mc.getColumn();
            thisStr = "";
            rowDatas.add(null);
        }
        if(!"".equals(thisStr)){
            if(thisStr == null){
                rowDatas.add(null);
            }else {
                rowDatas.add(new String(thisStr));
            }
        }
        // Update column and row count
        if(thisRow > -1)
            lastRowNumber = thisRow;
        if(thisColumn > -1)
            lastColumnNumber = thisColumn;
        // Handle end of row
        if(record instanceof LastCellOfRowDummyRecord) {
            if(lastColumnNumber == -1){
                readerContext.setLastColumnNumber(-1);
                readerContext.setCurrentRowNumber(((LastCellOfRowDummyRecord) record).getRow());
                readerContext.setBlankRow(true);
            }else{
                readerContext.setLastColumnNumber(lastColumnNumber);
                readerContext.setCurrentRowNumber(lastRowNumber);
                readerContext.setBlankRow(false);
            }
            notifyListener(rowDatas, readerContext);
            rowDatas.clear();
            // We're onto a new row
            lastColumnNumber = -1;
        }
    }

    public void process(String excelPath, ExcelReaderListener excelReaderListener, Class clazz) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(new File(excelPath));
        process(fileInputStream, excelReaderListener, clazz);
        fileInputStream.close();
    }

    public void process(File file, ExcelReaderListener excelReaderListener, Class clazz) throws IOException {
        FileInputStream fileInputStream = new FileInputStream(file);
        process(fileInputStream, excelReaderListener, clazz);
        fileInputStream.close();
    }

    public void process(InputStream fileInputStream, ExcelReaderListener excelReaderListener, Class clazz) throws IOException {

        boundSheetRecords = new ArrayList<>();
        readerContext = new ReaderContext();
        rowDatas = new ArrayList<>();
        readerContext.setExcelType(ExcelType.XLS);
        readerContext.setClazz(clazz);
        excelReaderListeners.add(excelReaderListener);

        fs = new POIFSFileSystem(fileInputStream);
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest req = new HSSFRequest();
        if(outputFormulaValues) {
            req.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
            req.addListenerForAllRecords(workbookBuildingListener);
        }
        factory.processWorkbookEvents(req, fs);
        fs.close();
        fileInputStream.close();
    }

}
