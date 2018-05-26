package example;

import com.github.ldzzdl.easyexcel4j.metadata.ExcelType;
import com.github.ldzzdl.easyexcel4j.reader.util.ExcelReaderUtil;
import com.github.ldzzdl.easyexcel4j.writer.util.ExcelWriterContext;
import com.github.ldzzdl.easyexcel4j.writer.util.ExcelWriterUtil;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.junit.Test;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author LDZZDL
 * @create 2018-05-22 21:26
 **/
public class TestPerson {

    @Test
    public void test1() throws OpenXML4JException, SAXException, IOException {
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        //第一个参数Excel文件的路径
        //第二个参数Excel文件的类型，使用枚举类型
        //第三个参数JavaBean的class属性
        //第四个参数开始读取数据的行数
        //第五个参数开始读取sheet的序号（从1开始）
        //第六个参数sheet的结束序号（0代表读到末尾）
        //方法根据JavaBean注解上的excelOrder来读取属性,进行JavaBean的封装
        List<Person> people = excelReaderUtil.readExcel2ModelListByOrder("src/test/java/doc/person.xls", ExcelType.XLS, Person.class, 4,1,0);

    }

    @Test
    public void test2() throws OpenXML4JException, SAXException, IOException {
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        //第一个参数Excel文件的路径
        //第二个参数Excel文件的类型，使用枚举类型
        //第三个参数JavaBean的class属性
        //第四个参数开始读取sheet的序号（从1开始）
        //第五个参数sheet的结束序号（0代表读到末尾）
        //方法根据JavaBean注解上的excelTitle来读取属性,进行JavaBean的封装
        List<Person> people = excelReaderUtil.readExcel2ModelListByTitle("src/test/java/doc/person.xls", ExcelType.XLS, Person.class,1,0);
    }

    @Test
    public void test3() throws OpenXML4JException, SAXException, IOException {
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        List<String> strings = excelReaderUtil.readExcel2List("src/test/java/doc/person.xlsx", ExcelType.XLSX,1,1,0);
    }

//    @Test
//    public void test4() throws OpenXML4JException, SAXException, IOException {
//        ReaderResolverV07 readerResolverV07 = new ReaderResolverV07();
//        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/large.xlsx");
//        readerResolverV07.process(inputStream, new LargeSheetListener(), null);
//    }

    @Test
    public void test5() throws InvocationTargetException, NoSuchMethodException, IllegalAccessException, IOException {
        ExcelWriterUtil excelWriterUtil = new ExcelWriterUtil();
        List<Person> people = new ArrayList<>();
        Person person = new Person();
        person.setAge(12);
        person.setName("jack");
        person.setHobby("篮球");
        people.add(person);
        people.add(person);
        people.add(person);
        ExcelWriterContext excelWriterContext = new ExcelWriterContext();
        excelWriterContext.setFontName("宋体");
        excelWriterContext.setFontHeightInPoints((short)11);
        excelWriterContext.setExcelType(ExcelType.XLSX);
        excelWriterUtil.writeModelList2Excel(people, excelWriterContext);
    }

}
