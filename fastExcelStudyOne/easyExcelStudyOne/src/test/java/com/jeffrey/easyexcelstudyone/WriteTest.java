package com.jeffrey.easyexcelstudyone;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.data.CommentData;
import com.alibaba.excel.metadata.data.FormulaData;
import com.alibaba.excel.metadata.data.HyperlinkData;
import com.alibaba.excel.metadata.data.ImageData;
import com.alibaba.excel.metadata.data.RichTextStringData;
import com.alibaba.excel.metadata.data.WriteCellData;
import com.alibaba.excel.util.BooleanUtils;
import com.alibaba.excel.util.FileUtils;
import com.alibaba.excel.util.ListUtils;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.handler.context.CellWriteHandlerContext;
import com.alibaba.excel.write.merge.LoopMergeStrategy;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.WriteTable;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.jeffrey.easyexcelstudyone.entity.out.ComplexHeadData;
import com.jeffrey.easyexcelstudyone.entity.out.ConverterData;
import com.jeffrey.easyexcelstudyone.entity.out.DataClass1;
import com.jeffrey.easyexcelstudyone.entity.out.DataClass2;
import com.jeffrey.easyexcelstudyone.entity.out.DemoMergeData;
import com.jeffrey.easyexcelstudyone.entity.out.DemoStyleData;
import com.jeffrey.easyexcelstudyone.entity.out.DemoStyleData2;
import com.jeffrey.easyexcelstudyone.entity.out.DemoTableData;
import com.jeffrey.easyexcelstudyone.entity.out.DogEntity;
import com.jeffrey.easyexcelstudyone.entity.out.DogEntity2;
import com.jeffrey.easyexcelstudyone.entity.out.ImageDemoData;
import com.jeffrey.easyexcelstudyone.entity.out.LongestMatchColumnWidthData;
import com.jeffrey.easyexcelstudyone.entity.out.ProductEntity;
import com.jeffrey.easyexcelstudyone.entity.out.WriteCellDemoData;
import com.jeffrey.easyexcelstudyone.handler.CommentWriteHandler;
import com.jeffrey.easyexcelstudyone.handler.CustomCellWriteHandler;
import com.jeffrey.easyexcelstudyone.handler.CustomSheetWriteHandler;
import com.jeffrey.easyexcelstudyone.style.CustomHeadStyleHandler;
import com.jeffrey.easyexcelstudyone.utils.PathUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

/**
 * @author jeffrey
 * @version 1.0
 * @date 2025/2/18
 * @time 17:15
 * @week 星期二
 * @description 导出的测试类
 **/
@Slf4j
public class WriteTest {

    @Test
    public void writeTest001() throws IOException {

        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest001.xlsx");
        log.info("outputFile:{}", outputFile);

        List<DogEntity> dataList = new ArrayList<>();
        Date date = new Date();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DogEntity dogEntity = new DogEntity();
            dogEntity.setName("dog" + index);
            dogEntity.setId(Long.parseLong(index + ""));
            dogEntity.setPrice(index * 100 + 2.222);
            dogEntity.setColor("black" + index);
            dogEntity.setBirthday(date);
            dataList.add(dogEntity);
        }

        // 创建 Excel 文件
        EasyExcel.write(outputFile, DogEntity.class)    // YourDataClass 是要写入的数据类
                .sheet("Sheet1")                           // Sheet 名称
                .doWrite(dataList);
    }

    @Test
    public  void writeTest002() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest002.xlsx");
        log.info("outputFile:{}", outputFile);
        // 创建 ExcelWriter
        ExcelWriter excelWriter = EasyExcel.write(outputFile).build();

        // 导出第一个 Sheet
        WriteSheet writeSheet1 = EasyExcel.writerSheet("Sheet1").head(DataClass1.class).build();
        List<DataClass1> data1 = createData1();
        excelWriter.write(data1, writeSheet1);

        // 导出第二个 Sheet
        WriteSheet writeSheet2 = EasyExcel.writerSheet("Sheet2").head(DataClass2.class).build();
        List<DataClass2> data2 = createData2();
        excelWriter.write(data2, writeSheet2);

        // 结束写入
        excelWriter.finish();
    }

    private static List<DataClass1> createData1() {
        List<DataClass1> list = new ArrayList<>();
        list.add(new DataClass1("Alice", 30));
        list.add(new DataClass1("Bob", 25));
        return list;
    }

    private static List<DataClass2> createData2() {
        List<DataClass2> list = new ArrayList<>();
        list.add(new DataClass2("Product1", 100));
        list.add(new DataClass2("Product2", 200));
        return list;
    }


    /**
     * 复杂的表头有限，需要生成表后使用poi 进行修改
     * @throws IOException
     */
    @Test
    public void writeTest003() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest003.xlsx");
        log.info("outputFile:{}", outputFile);

        // 创建 ExcelWriter
        ExcelWriter excelWriter = EasyExcel.write(outputFile).build();

        // 创建 Sheet
        WriteSheet writeSheet = EasyExcel.writerSheet("Complex Header").head(getHead()).build();

        // 写入数据
        List<ProductEntity> dataList = createData();
        excelWriter.write(dataList, writeSheet);

        // 结束写入
        excelWriter.finish();
    }

    private static List<List<String>> getHead() {
        List<List<String>> head = new ArrayList<>();

        // 第一行（合并单元格的标题）
        List<String> row1 = new ArrayList<>();
        row1.add("个人信息");
        row1.add("");
        row1.add("商品信息");
        head.add(row1);

        // 第二行（实际列的标题）
        List<String> row2 = new ArrayList<>();
        row2.add("姓名");
        row2.add("年龄");
        row2.add("商品名称");
        row2.add("价格");
        head.add(row2);

        return head;
    }

    private static List<ProductEntity> createData() {
        List<ProductEntity> list = new ArrayList<>();
        list.add(new ProductEntity("Alice", 30, "Product1", 100));
        list.add(new ProductEntity("Bob", 25, "Product2", 200));
        return list;
    }


    /**
     * 自定义样式
     * @throws IOException
     */
    @Test
    public void writeTest004() throws IOException {

        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest004.xlsx");
        log.info("outputFile:{}", outputFile);

        List<DogEntity> dataList = new ArrayList<>();
        Date date = new Date();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DogEntity dogEntity = new DogEntity();
            dogEntity.setName("dog" + index);
            dogEntity.setId(Long.parseLong(index + ""));
            dogEntity.setPrice(index * 100 + 2.222);
            dogEntity.setColor("black" + index);
            dogEntity.setBirthday(date);
            dataList.add(dogEntity);
        }

        List<List<String>> lists = new ArrayList<>();
        for (int i = 0; i < 5; i++) {
            List<String> list = new ArrayList<>();
            list.add("dog" + i);
            lists.add(list);
        }

        // 创建 Excel 文件
        EasyExcel.write(outputFile, DogEntity.class)
                .head(lists)// YourDataClass 是要写入的数据类
                .registerWriteHandler(new CustomHeadStyleHandler())
                .sheet("Sheet1")
                .doWrite(dataList);
    }


    /**
     * 排除部分字段
     * @throws IOException
     */
    @Test
    public void writeTest005() throws IOException {

        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest005_excludeFieldWirte.xlsx");
        log.info("outputFile:{}", outputFile);

        List<DogEntity2> dataList = new ArrayList<>();
        Date date = new Date();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DogEntity2 dogEntity = new DogEntity2();
            dogEntity.setName("dog" + index);
            dogEntity.setId(Long.parseLong(index + ""));
            dogEntity.setPrice(index * 100 + 2.222);
            dogEntity.setColor("black" + index);
            dogEntity.setBirthday(date);
            dataList.add(dogEntity);
        }

        // 创建 Excel 文件
        EasyExcel.write(outputFile, DogEntity2.class)    // YourDataClass 是要写入的数据类
                .sheet("Sheet1")                           // Sheet 名称
                .excludeColumnFieldNames(new ArrayList<String>(Arrays.asList("price")))//排除的方式2
                .includeColumnIndexes(new ArrayList<Integer>(Arrays.asList(0, 1, 2, 3)))//前面排除的，再加入也不会显示出来
                .doWrite(dataList);
    }


    /**
     * 复杂头写入
     * <p>1. 创建excel对应的实体对象 参照{@link ComplexHeadData}
     * <p>2. 使用{@link ExcelProperty}注解指定复杂的头
     * <p>3. 直接写即可
     */
    @Test
    public void complexHeadWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest006_complexHeaderWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        List<ComplexHeadData> complexHeadDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            ComplexHeadData complexHeadData = new ComplexHeadData();
            int index = i + 1;
            complexHeadData.setString("字符串" + index);
            complexHeadData.setDoubleData(index * 1.1);
            complexHeadData.setDate(new Date());
            complexHeadDataList.add(complexHeadData);
        }
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(outputFile, ComplexHeadData.class).sheet("模板").doWrite(complexHeadDataList);
    }


    /**
     * 日期、数字或者自定义格式转换
     * <p>1. 创建excel对应的实体对象 参照{@link ConverterData}
     * <p>2. 使用{@link ExcelProperty}配合使用注解{@link DateTimeFormat}、{@link NumberFormat}或者自定义注解
     * <p>3. 直接写即可
     */
    @Test
    public void converterWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest007_converterWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        List<ConverterData> converterDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            ConverterData converterData = new ConverterData();
            int index = i + 1;
            converterData.setString("字符串" + index);
            converterData.setDoubleData(index * 0.2451);
            converterData.setDate(new Date());
            converterDataList.add(converterData);
        }
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(outputFile, ConverterData.class).sheet("模板").doWrite(converterDataList);
    }


    /**
     * 图片导出
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link ImageDemoData}
     * <p>
     * 2. 直接写即可
     */
    @Test
    public void imageWrite() throws Exception {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest008_imageWrite.xlsx");
        log.info("outputFile:{}", outputFile);

        // 这里注意下 所有的图片都会放到内存 暂时没有很好的解法，大量图片的情况下建议 2选1:
        // 1. 将图片上传到oss 或者其他存储网站: https://www.aliyun.com/product/oss ，然后直接放链接
        // 2. 使用: https://github.com/coobird/thumbnailator 或者其他工具压缩图片

        // 输入的Excel文件路径
        URL systemResource = ClassLoader.getSystemResource("static/image/W0001.png");
        String imagePath = systemResource.getPath();
        log.info("inputFile:{}", imagePath);
        try {
            InputStream inputStream = FileUtils.openInputStream(new File(imagePath));
            List<ImageDemoData> list =  ListUtils.newArrayList();
            ImageDemoData imageDemoData = new ImageDemoData();
            list.add(imageDemoData);
            // 放入五种类型的图片 实际使用只要选一种即可
            imageDemoData.setByteArray(FileUtils.readFileToByteArray(new File(imagePath)));
            imageDemoData.setFile(new File(imagePath));
            imageDemoData.setString(imagePath);
            imageDemoData.setInputStream(inputStream);
            imageDemoData.setUrl(new URL(
                    "https://p3.ssl.qhimgs1.com/t015e7cb7359d4b1ef9.jpg"));

            // 这里演示
            // 需要额外放入文字
            // 而且需要放入2个图片
            // 第一个图片靠左
            // 第二个靠右 而且要额外的占用他后面的单元格
            WriteCellData<Void> writeCellData = new WriteCellData<>();
            imageDemoData.setWriteCellDataFile(writeCellData);
            // 这里可以设置为 EMPTY 则代表不需要其他数据了
            writeCellData.setType(CellDataTypeEnum.STRING);
            writeCellData.setStringValue("额外的放一些文字");

            // 可以放入多个图片
            List<ImageData> imageDataList = new ArrayList<>();
            ImageData imageData = new ImageData();
            imageDataList.add(imageData);
            writeCellData.setImageDataList(imageDataList);
            // 放入2进制图片
            imageData.setImage(FileUtils.readFileToByteArray(new File(imagePath)));
            // 图片类型
            imageData.setImageType(ImageData.ImageType.PICTURE_TYPE_PNG);
            // 上 右 下 左 需要留空
            // 这个类似于 css 的 margin
            // 这里实测 不能设置太大 超过单元格原始大小后 打开会提示修复。暂时未找到很好的解法。
            imageData.setTop(5);
            imageData.setRight(40);
            imageData.setBottom(5);
            imageData.setLeft(5);

            // 放入第二个图片
            imageData = new ImageData();
            imageDataList.add(imageData);
            writeCellData.setImageDataList(imageDataList);
            imageData.setImage(FileUtils.readFileToByteArray(new File(imagePath)));
            imageData.setImageType(ImageData.ImageType.PICTURE_TYPE_PNG);
            imageData.setTop(5);
            imageData.setRight(5);
            imageData.setBottom(5);
            imageData.setLeft(50);
            // 设置图片的位置 假设 现在目标 是 覆盖 当前单元格 和当前单元格右边的单元格
            // 起点相对于当前单元格为0 当然可以不写
            imageData.setRelativeFirstRowIndex(0);
            imageData.setRelativeFirstColumnIndex(0);
            imageData.setRelativeLastRowIndex(0);
            // 前面3个可以不写  下面这个需要写 也就是 结尾 需要相对当前单元格 往右移动一格
            // 也就是说 这个图片会覆盖当前单元格和 后面的那一格
            imageData.setRelativeLastColumnIndex(1);

            // 写入数据
            EasyExcel.write(outputFile, ImageDemoData.class).sheet().doWrite(list);
        } catch (Exception e) {
            log.error("图片导出失败,e={}", e);
        }
    }

    /**
     * 超链接、备注、公式、指定单个单元格的样式、单个单元格多种样式
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link WriteCellDemoData}
     * <p>
     * 2. 直接写即可
     *
     * @since 3.0.0-beta1
     */
    @Test
    public void writeCellDataWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest009_writeCellDataWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        WriteCellDemoData writeCellDemoData = new WriteCellDemoData();

        // 设置超链接
        WriteCellData<String> hyperlink = new WriteCellData<>("官方网站");
        writeCellDemoData.setHyperlink(hyperlink);
        HyperlinkData hyperlinkData = new HyperlinkData();
        hyperlink.setHyperlinkData(hyperlinkData);
        hyperlinkData.setAddress("https://github.com/alibaba/easyexcel");
        hyperlinkData.setHyperlinkType(HyperlinkData.HyperlinkType.URL);

        // 设置备注
        WriteCellData<String> comment = new WriteCellData<>("备注的单元格信息");
        writeCellDemoData.setCommentData(comment);
        CommentData commentData = new CommentData();
        comment.setCommentData(commentData);
        commentData.setAuthor("Jiaju Zhuang");
        commentData.setRichTextStringData(new RichTextStringData("这是一个备注"));
        // 备注的默认大小是按照单元格的大小 这里想调整到4个单元格那么大 所以向后 向下 各额外占用了一个单元格
        commentData.setRelativeLastColumnIndex(1);
        commentData.setRelativeLastRowIndex(1);

        // 设置公式
        WriteCellData<String> formula = new WriteCellData<>();
        writeCellDemoData.setFormulaData(formula);
        FormulaData formulaData = new FormulaData();
        formula.setFormulaData(formulaData);
        // 将 123456789 中的第一个数字替换成 2
        // 这里只是例子 如果真的涉及到公式 能内存算好尽量内存算好 公式能不用尽量不用
        formulaData.setFormulaValue("REPLACE(123456789,1,1,2)");

        // 设置单个单元格的样式 当然样式 很多的话 也可以用注解等方式。
        WriteCellData<String> writeCellStyle = new WriteCellData<>("单元格样式");
        writeCellStyle.setType(CellDataTypeEnum.STRING);
        writeCellDemoData.setWriteCellStyle(writeCellStyle);
        WriteCellStyle writeCellStyleData = new WriteCellStyle();
        writeCellStyle.setWriteCellStyle(writeCellStyleData);
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.
        writeCellStyleData.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景绿色
        writeCellStyleData.setFillForegroundColor(IndexedColors.GREEN.getIndex());

        // 设置单个单元格多种样式
        // 这里需要设置 inMomery=true 不然会导致无法展示单个单元格多种样式，所以慎用
        WriteCellData<String> richTest = new WriteCellData<>();
        richTest.setType(CellDataTypeEnum.RICH_TEXT_STRING);
        writeCellDemoData.setRichText(richTest);
        RichTextStringData richTextStringData = new RichTextStringData();
        richTest.setRichTextStringDataValue(richTextStringData);
        richTextStringData.setTextString("红色绿色默认");
        // 前2个字红色
        WriteFont writeFont = new WriteFont();
        writeFont.setColor(IndexedColors.RED.getIndex());
        richTextStringData.applyFont(0, 2, writeFont);
        // 接下来2个字绿色
        writeFont = new WriteFont();
        writeFont.setColor(IndexedColors.GREEN.getIndex());
        richTextStringData.applyFont(2, 4, writeFont);

        List<WriteCellDemoData> data = new ArrayList<>();
        data.add(writeCellDemoData);
        EasyExcel.write(outputFile, WriteCellDemoData.class).inMemory(true).sheet("模板").doWrite(data);
    }

    /**
     * 注解形式自定义样式
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoStyleData}
     * <p>
     * 3. 直接写即可
     *
     * @since 2.2.0-beta1
     */
    @Test
    public void annotationStyleWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest010_annotationStyleWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        List<DemoStyleData> demoStyleDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DemoStyleData demoStyleData = new DemoStyleData("字符串" + index, new Date(), 0.56 * i);
            demoStyleDataList.add(demoStyleData);
        }
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(outputFile, DemoStyleData.class).sheet("模板").doWrite(demoStyleDataList);
    }


    /**
     * 拦截器形式自定义样式
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>
     * 2. 创建一个style策略 并注册
     * <p>
     * 3. 直接写即可
     */
    @Test
    public void handlerStyleWrite() throws IOException {
        // 方法1 使用已有的策略 推荐
        // HorizontalCellStyleStrategy 每一行的样式都一样 或者隔行一样
        // AbstractVerticalCellStyleStrategy 每一列的样式都一样 需要自己回调每一页
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest011_handlerStyleWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        List<DemoStyleData2> demoStyleDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DemoStyleData2 demoStyleData = new DemoStyleData2("字符串" + index, new Date(), 0.56 * i);
            demoStyleDataList.add(demoStyleData);
        }
        ExcelWriter excelWriter = EasyExcel.write(outputFile).build();
        // 头的策略
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        // 背景设置为红色
        headWriteCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short)20);
        headWriteCellStyle.setWriteFont(headWriteFont);
        // 内容的策略
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND 不然无法显示背景颜色.头默认了 FillPatternType所以可以不指定
        contentWriteCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);
        // 背景绿色
        contentWriteCellStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
        WriteFont contentWriteFont = new WriteFont();
        // 字体大小
        contentWriteFont.setFontHeightInPoints((short)20);
        contentWriteCellStyle.setWriteFont(contentWriteFont);

        // 这个策略是 头是头的样式 内容是内容的样式 其他的策略可以自己实现
        HorizontalCellStyleStrategy horizontalCellStyleStrategy =
                new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);

        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭

        // 导出第一个 Sheet
        WriteSheet writeSheet1 = EasyExcel
                .writerSheet("Sheet1")
                .head(DemoStyleData2.class)
                .registerWriteHandler(horizontalCellStyleStrategy)
                .build();
        excelWriter.write(demoStyleDataList, writeSheet1);


        // 方法2: 使用easyexcel的方式完全自己写 不太推荐 尽量使用已有策略
        // @since 3.0.0-beta2
        WriteSheet writeSheet2 = EasyExcel.writerSheet("Sheet2")
                .head(DemoStyleData2.class)
                .registerWriteHandler(new CellWriteHandler() {
                    @Override
                    public void afterCellDispose(CellWriteHandlerContext context) {
                        // 当前事件会在 数据设置到poi的cell里面才会回调
                        // 判断不是头的情况 如果是fill 的情况 这里会==null 所以用not true
                        if (BooleanUtils.isNotTrue(context.getHead())) {
                            // 第一个单元格
                            // 只要不是头 一定会有数据 当然fill的情况 可能要context.getCellDataList() ,这个需要看模板，因为一个单元格会有多个 WriteCellData
                            WriteCellData<?> cellData = context.getFirstCellData();
                            // 这里需要去cellData 获取样式
                            // 很重要的一个原因是 WriteCellStyle 和 dataFormatData绑定的 简单的说 比如你加了 DateTimeFormat
                            // ，已经将writeCellStyle里面的dataFormatData 改了 如果你自己new了一个WriteCellStyle，可能注解的样式就失效了
                            // 然后 getOrCreateStyle 用于返回一个样式，如果为空，则创建一个后返回
                            WriteCellStyle writeCellStyle = cellData.getOrCreateStyle();
                            writeCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                            // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND
                            writeCellStyle.setFillPatternType(FillPatternType.SOLID_FOREGROUND);

                            // 这样样式就设置好了 后面有个FillStyleCellWriteHandler 默认会将 WriteCellStyle 设置到 cell里面去 所以可以不用管了
                        }
                    }
                }).build();

        excelWriter.write(demoStyleDataList, writeSheet2);

        // 方法3: 使用poi的样式完全自己写 不推荐
        // @since 3.0.0-beta2
        // 坑1：style里面有dataformat 用来格式化数据的 所以自己设置可能导致格式化注解不生效
        // 坑2：不要一直去创建style 记得缓存起来 最多创建6W个就挂了

        WriteSheet writeSheet3 = EasyExcel.writerSheet("Sheet3")
                .head(DemoStyleData2.class)
                .registerWriteHandler(new CellWriteHandler() {
                    @Override
                    public void afterCellDispose(CellWriteHandlerContext context) {
                        // 当前事件会在 数据设置到poi的cell里面才会回调
                        // 判断不是头的情况 如果是fill 的情况 这里会==null 所以用not true
                        if (BooleanUtils.isNotTrue(context.getHead())) {
                            Cell cell = context.getCell();
                            // 拿到poi的workbook
                            Workbook workbook = context.getWriteWorkbookHolder().getWorkbook();
                            // 这里千万记住 想办法能复用的地方把他缓存起来 一个表格最多创建6W个样式
                            // 不同单元格尽量传同一个 cellStyle
                            CellStyle cellStyle = workbook.createCellStyle();
                            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                            // 这里需要指定 FillPatternType 为FillPatternType.SOLID_FOREGROUND
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                            cell.setCellStyle(cellStyle);

                            // 由于这里没有指定dataformat 最后展示的数据 格式可能会不太正确

                            // 这里要把 WriteCellData的样式清空， 不然后面还有一个拦截器 FillStyleCellWriteHandler 默认会将 WriteCellStyle 设置到
                            // cell里面去 会导致自己设置的不一样
                            context.getFirstCellData().setWriteCellStyle(null);
                        }
                    }
                }).build();
        excelWriter.write(demoStyleDataList, writeSheet3);
        // 结束写入
        excelWriter.finish();
    }



    /**
     * 合并单元格
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData} {@link DemoMergeData}
     * <p>
     * 2. 创建一个merge策略 并注册
     * <p>
     * 3. 直接写即可
     *
     * @since 2.2.0-beta1
     */
    @Test
    public void mergeWrite() throws IOException {
        // 方法1 注解
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest012_mergeWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        ExcelWriter excelWriter = EasyExcel.write(outputFile).build();
        List<DemoMergeData> demoMergeDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DemoMergeData demoStyleData = new DemoMergeData("字符串" + index, new Date(), 0.56 * i);
            demoMergeDataList.add(demoStyleData);
        }
        // 在DemoStyleData里面加上ContentLoopMerge注解
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        WriteSheet writeSheet = EasyExcel.writerSheet("Sheet1").head(DemoMergeData.class).build();
        excelWriter.write(demoMergeDataList, writeSheet);

        // 方法2 自定义合并单元格策略
        // 每隔2行会合并 把eachColumn 设置成 3 也就是我们数据的长度，所以就第一列会合并。当然其他合并策略也可以自己写
        LoopMergeStrategy loopMergeStrategy = new LoopMergeStrategy(2, 0);
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        WriteSheet writeSheet2 = EasyExcel.writerSheet("Sheet2").head(DemoMergeData.class).build();
        excelWriter.write(demoMergeDataList, writeSheet2);

        excelWriter.finish();
    }

    /**
     * 使用table去写入
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>
     * 2. 然后写入table即可
     */
    @Test
    public void tableWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest013_tableWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        // 方法1 这里直接写多个table的案例了，如果只有一个 也可以直一行代码搞定，参照其他案
        // 这里 需要指定写用哪个class去写
        try (ExcelWriter excelWriter = EasyExcel.write(outputFile, DemoTableData.class).build()) {
            // 把sheet设置为不需要头 不然会输出sheet的头 这样看起来第一个table 就有2个头了
            WriteSheet writeSheet = EasyExcel.writerSheet("模板").needHead(Boolean.FALSE).build();
            // 这里必须指定需要头，table 会继承sheet的配置，sheet配置了不需要，table 默认也是不需要
            WriteTable writeTable0 = EasyExcel.writerTable(0).needHead(Boolean.TRUE).build();
            WriteTable writeTable1 = EasyExcel.writerTable(1).needHead(Boolean.TRUE).build();
            List<DemoTableData> demoTableDataList = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                int index = i + 1;
                DemoTableData demoStyleData = new DemoTableData("字符串" + index, new Date(), 0.56 * i);
                demoTableDataList.add(demoStyleData);
            }
            // 第一次写入会创建头
            excelWriter.write(demoTableDataList, writeSheet, writeTable0);
            // 第二次写如也会创建头，然后在第一次的后面写入数据
            excelWriter.write(demoTableDataList, writeSheet, writeTable1);
        }
    }

    /**
     * 自动列宽(不太精确)
     * <p>
     * 这个目前不是很好用，比如有数字就会导致换行。而且长度也不是刚好和实际长度一致。 所以需要精确到刚好列宽的慎用。 当然也可以自己参照
     * {@link LongestMatchColumnWidthStyleStrategy}重新实现.
     * <p>
     * poi 自带{@link SXSSFSheet#autoSizeColumn(int)} 对中文支持也不太好。目前没找到很好的算法。 有的话可以推荐下。
     *
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link LongestMatchColumnWidthData}
     * <p>
     * 2. 注册策略{@link LongestMatchColumnWidthStyleStrategy}
     * <p>
     * 3. 直接写即可
     */
    @Test
    public void longestMatchColumnWidthWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest014_longestMatchColumnWidthWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(outputFile, LongestMatchColumnWidthData.class)
                .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy()).sheet("模板").doWrite(dataLong());
    }

    private List<LongestMatchColumnWidthData> dataLong() {
        List<LongestMatchColumnWidthData> list = new ArrayList<LongestMatchColumnWidthData>();
        for (int i = 0; i < 10; i++) {
            LongestMatchColumnWidthData data = new LongestMatchColumnWidthData();
            data.setString("测试很长的字符串测试很长的字符串测试很长的字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(1000000000000.0);
            list.add(data);
        }
        return list;
    }


    /**
     * 下拉，超链接等自定义拦截器（上面几点都不符合但是要对单元格进行操作的参照这个）
     * <p>
     * demo这里实现2点。1. 对第一行第一列的头超链接到:https://github.com/alibaba/easyexcel 2. 对第一列第一行和第二行的数据新增下拉框，显示 测试1 测试2
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>
     * 2. 注册拦截器 {@link CustomCellWriteHandler} {@link CustomSheetWriteHandler}
     * <p>
     * 2. 直接写即可
     */
    @Test
    public void customHandlerWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest015_customHandlerWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        List<DemoTableData> demoTableDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DemoTableData demoStyleData = new DemoTableData("字符串" + index, new Date(), 0.56 * i);
            demoTableDataList.add(demoStyleData);
        }
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(outputFile, DemoTableData.class).registerWriteHandler(new CustomSheetWriteHandler())
                .registerWriteHandler(new CustomCellWriteHandler()).sheet("模板").doWrite(demoTableDataList);
    }


    /**
     * 插入批注
     * <p>
     * 1. 创建excel对应的实体对象 参照{@link DemoData}
     * <p>
     * 2. 注册拦截器 {@link CommentWriteHandler}
     * <p>
     * 2. 直接写即可
     */
    @Test
    public void commentWrite() throws IOException {
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest016_commentWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        List<DemoTableData> demoTableDataList = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            int index = i + 1;
            DemoTableData demoStyleData = new DemoTableData("字符串" + index, new Date(), 0.56 * i);
            demoTableDataList.add(demoStyleData);
        }
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // 这里要注意inMemory 要设置为true，才能支持批注。目前没有好的办法解决 不在内存处理批注。这个需要自己选择。
        EasyExcel.write(outputFile, DemoTableData.class).inMemory(Boolean.TRUE).registerWriteHandler(new CommentWriteHandler())
                .sheet("模板").doWrite(demoTableDataList);
    }


    /**
     * 不创建对象的写
     */
    @Test
    public void noModelWrite() throws IOException {
        // 写法1
        // 输出的Excel文件路径
        String outputFile = PathUtil.createFileUnderResource("out/writeTest017_noModelWrite.xlsx");
        log.info("outputFile:{}", outputFile);
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        EasyExcel.write(outputFile).head(head()).sheet("模板").doWrite(dataList());
    }

    private List<List<String>> head() {
        List<List<String>> list = ListUtils.newArrayList();
        List<String> head0 = ListUtils.newArrayList();
        head0.add("字符串" + System.currentTimeMillis());
        List<String> head1 = ListUtils.newArrayList();
        head1.add("数字" + System.currentTimeMillis());
        List<String> head2 = ListUtils.newArrayList();
        head2.add("日期" + System.currentTimeMillis());
        list.add(head0);
        list.add(head1);
        list.add(head2);
        return list;
    }

    private List<List<Object>> dataList() {
        List<List<Object>> list = ListUtils.newArrayList();
        for (int i = 0; i < 10; i++) {
            List<Object> data = ListUtils.newArrayList();
            data.add("字符串" + i);
            data.add(0.56);
            data.add(new Date());
            list.add(data);
        }
        return list;
    }


}
