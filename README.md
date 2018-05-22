# easyexcel4j
基于Apache POI的理海量数据导入和简单导出的Excel工具
# 快速使用
## 导入
### 注解导入

- 范围

<table>
    <tr>
        <td>支持Excel的文件格式</td>
        <td>xls、xlsx</td>
    </tr>
</table>

- 前提

程序将读取到的每一行数据封装成JavaBean,每一列的值作为JavaBean的属性，所以Excel文件内容排版需要类似以下格式
即每一列的值属于同一个属性，每一行的属性值属于同一个Class对象
<table>
    <tr>
        <th>姓名</th>
        <th>年龄</th>
        <th>生日</th>
    </tr>
    <tr>
        <th>Mike</th>
        <th>10</th>
        <th>2008年3月5日</th>
    </tr>
    <tr>
        <th>Kobe</th>
        <th>39</th>
        <th>1995年5月2日</th>
    </tr>
    <tr>
        <th>Jessy</th>
        <th>20</th>
        <th>1996年7月20日</th>
    </tr>
</table>

暂时不支持复杂的Excel的内容格式，如下所示
<table>
    <tr>
        <th>客户名称</th>
        <th>销售时间</th>
        <th>销售单号</th>
        <th>出仓单号</th>
        <th>客户下单时间</th>
    </tr>
    <tr>
        <th>XXX有限公司</th>
        <th>2016/3/1</th>
        <th>cc060301</th>
        <th>06040023</th>
        <th>2016年2月27日</th>
    </tr>
    <tr>
        <th>品牌</th>
        <th>规格</th>
        <th>数量</th>
        <th>单价</th>
        <th>金额</th>
    </tr>
    <tr>
        <th>品牌1</th>
        <th>A</th>
        <th>200</th>
        <th>300.00</th>
        <th>60000.00</th>
    </tr>
    <tr>
        <th>品牌2</th>
        <th>B</th>
        <th>30</th>
        <th>1200.00</th>
        <th>36000.00</th>
    </tr>
        <tr>
        <th>品牌3</th>
        <th>D</th>
        <th>10</th>
        <th>750.00</th>
        <th>7500.00</th>
    </tr>
    <tr>
        <th></th>
        <th></th>
        <th></th>
        <th>总计</th>
        <th>103500.00</th>
    </tr>
    <tr>
        <th>客户</th>
        <th></th>
        <th>主管</th>
        <th></th>
        <th></th>
    </tr>
    <tr>
        <th>签收时间</th>
        <th></th>
        <th>签收时间</th>
        <th></th>
        <th></th>
    </tr>
</table>

- 注解使用

1. 定义

```java
public @interface Excel {
    String excelTitle() default "";
    int excelOrder() default 0;
}
```

2. 作用
<table>
    <tr>
        <th>使用范围</th>
        <th>类的属性</th>
    </tr>
    <tr>
        <th>excelTitle</th>
        <th>class属性所对应的excel列的标题</th>
    </tr>
    <tr>
        <th>excelOrder</th>
        <th>excel属性所对应的excel列的序号，从1开始</th>
    </tr>
</table>

3. 举例
假如你要导入如下图所示的Excel文件，并封装成Person对象

![excel文件](https://picture-1253615005.cos.ap-guangzhou.myqcloud.com/Snipaste_2018-05-22_20-57-51.png)

只需要在Person对象中添加注解
```java
public class Person{
    //2代表第2列
    @Excel(excelOrder = 2, excelTitle = "姓名")
    private String name;
    //3代表第三列
    @Excel(excelOrder = 3, excelTitle = "年龄")
    private Integer age;
    @Excel(excelOrder = 5, excelTitle = "爱好")
    private String hobby;
    @Excel(excelOrder = 6, excelTitle = "生日")
    private Date birthday;
    //省略get、set、toString方法
}
```

- 方法使用

使用`ExcelReaderUtil`工具类中的`readExcel2ModelListByOrder`或者`readExcel2ModelListByTitle`方法
```java
    public class ImportPerson {
    
        @Test
        public void test1() throws OpenXML4JException, SAXException, IOException {
            ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
            String path = ImportPerson.class.getResource("/").getPath();
            //第一个参数Excel文件的路径，第一个参数可以替换为Excel文件的IO流，或者Excel文件的File对象
            //第二个参数Excel文件的类型，使用枚举类型
            //第三个参数JavaBean的class属性
            //第四个参数开始读取数据的行数
            //第五个参数开始读取sheet的序号（从1开始）
            //第六个参数sheet的结束序号（0代表读到末尾）
            //方法根据JavaBean注解上的excelOrder来读取属性,进行JavaBean的封装
            List<Person> people = excelReaderUtil.readExcel2ModelListByOrder(path + "/example/person.xls", ExcelType.XLS, Person.class, 4,1,0);
            for(Person person : people){
                System.out.println(person);
            }
        }
        @Test
        public void test2() throws OpenXML4JException, SAXException, IOException {
                ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
                String path = ImportPerson.class.getResource("/").getPath();
                //第一个参数Excel文件的路径，第一个参数可以替换为Excel文件的IO流，或者Excel文件的File对象
                //第二个参数Excel文件的类型，使用枚举类型
                //第三个参数JavaBean的class属性
                //第四个参数开始读取sheet的序号（从1开始）
                //第五个参数sheet的结束序号（0代表读到末尾）
                //方法根据JavaBean注解上的excelTitle来读取属性,进行JavaBean的封装
                List<Person> people = excelReaderUtil.readExcel2ModelListByTitle(path + "/example/person.xls", ExcelType.XLS, Person.class,1,0);
                for(Person person : people){
                    System.out.println(person);
                }
            }
    
    }
```

### 非注解导入

1. 假如你如要导入如下图所示的Excel文件

![excel文件](https://picture-1253615005.cos.ap-guangzhou.myqcloud.com/Snipaste_2018-05-22_20-57-51.png)

2. 使用`ExcelReaderUtil`工具类中的`readExcel2List`方法

```java
public class ImportPerson {

    @Test
    public void test3() throws OpenXML4JException, SAXException, IOException {
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        String path = ImportPerson.class.getResource("/").getPath();
        //第一个参数为Excel的文件路径，第一个参数可以替换为Excel文件的IO流，或者Excel文件的File对象
        //第二个参数Excel文件的类型，使用枚举类型
        //第三个参数开始读取数据的行数
        //第四个参数开始读取sheet的序号（从1开始）
        //第五个参数sheet的结束序号（0代表读到末尾）
        List<String> strings = excelReaderUtil.readExcel2List(path + "/example/person.xlsx", ExcelType.XLSX,1,1,0);
        for (String string : strings){
            System.out.println(string);
        }
    }

}
```

### 大数量导入
1. 比如你想从Excel中读取10万行的数据，插入的数据库中，或者统计相关信息，这个时候上述的方法可能对你不管用了，
因为你如果将10万行的数据保存到内存中，往往会导致内存溢出错误

2. 当然，Apache POI 提供了`Event API`来解决读取大量数据的问题，`Event API`的原理就是解析一部分Excel文件--->通知用户处理--->再解析一部分Excel文件

3. `Event API`的操作更接近底层，开发难度更大，同时V03和V07的Excel处理方式不一样

4. easyexcel4j对这些操作进行了整和，使得开发更加简单，只需简单的继承接口，即可
```java
//第一步实现ExcelReaderListener接口，该接口返回每行的数据和导入的上下文环境
public class LargeSheetListener implements ExcelReaderListener {
    
    @Override
    public void invoke(List<String> datas, ReaderContext readerContext) {
        if(readerContext.getCurrentRowNumber() % 10000 == 0){
            System.out.println("当前行为：" + readerContext.getCurrentRowNumber() +
                readerContext.getCurrentSheetIndex() + "," +
                readerContext.isBlankRow());
        }
    }
}

//读取100000行的25列的数据，55M左右的Excel文件，解析时间大概为12秒左右，没有进行任何其他操作
//创建V03或者V07解析器，将Excel文件的IO流和ExcelReaderListener的显示传入即可
public class ImportPerson {
    
    @Test
    public void test4() throws OpenXML4JException, SAXException, IOException {
        // ReaderResolverV03 readerResolverV03 = new readerResolverV03(); 
        ReaderResolverV07 readerResolverV07 = new ReaderResolverV07();
        InputStream inputStream = ExcelReaderUtilTest.class.getResourceAsStream("/doc/large.xlsx");
        readerResolverV07.process(inputStream, new LargeSheetListener(), null);
    }
}
```

## 导出
### 简单导出
```java
public class ImportPerson {

    @Test
    public void test5() throws InvocationTargetException, NoSuchMethodException, IllegalAccessException, IOException {
        ExcelWriterUtil excelWriterUtil = new ExcelWriterUtil();
        List<com.github.easyexcel4j.writer.util.Person> people = new ArrayList<>();
        com.github.easyexcel4j.writer.util.Person person = new com.github.easyexcel4j.writer.util.Person();
        person.setAge(12);
        person.setName("jack");
        person.setSport("篮球");
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

```
