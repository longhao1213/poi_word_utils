package com.lh.word;

import com.lh.word.form.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.XDDFColor;
import org.apache.poi.xddf.usermodel.XDDFShapeProperties;
import org.apache.poi.xddf.usermodel.XDDFSolidFillProperties;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: WordUtils.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/02 16:04
 */
public class WordUtils {

    /**
     * 获取图表对象
     *
     * @param document word对象
     * @param width    默认15
     * @param height   默认10
     * @return
     */
    public XWPFChart getChart(XWPFDocument document, Integer width, Integer height) throws IOException, InvalidFormatException {
        if (width == null) {
            width = 15;
        }
        if (height == null) {
            height = 10;
        }
        return document.createChart(width * Units.EMU_PER_CENTIMETER, height * Units.EMU_PER_CENTIMETER);
    }

    /**
     * 创建普通柱状图-簇状柱状图-堆叠柱状图
     *
     * @param chart        图表对象
     * @param barChartForm 数据对象
     */
    public void createBarChart(XWPFChart chart, BarChartForm barChartForm) throws Exception {
        String[] categories = barChartForm.getCategories();
        List<Double[]> tableData = barChartForm.getTableData();
        List<String> colorTitles = barChartForm.getColorTitles();
        String title = barChartForm.getTitle();
        if (colorTitles.size() != tableData.size()) {
            throw new Exception("颜色标题个数,必须和数组个数相同");
        }
        for (Double[] tableDatum : tableData) {
            if (tableDatum.length != categories.length) {
                throw new Exception("每个数组的元素个数,必须和");
            }
        }
        // 设置标题
        chart.setTitleText(title);
        //标题覆盖
        chart.setTitleOverlay(false);

        // 处理对应的数据
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        List<XDDFChartData.Series> seriesList = new ArrayList<>();

        // 创建一些轴
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(barChartForm.getBottomTitle());
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(barChartForm.getBottomTitle());
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        // 创建柱状图的类型
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        // 为图表添加数据
        for (int i = 0; i < tableData.size(); i++) {
            XDDFChartData.Series series = data.addSeries(categoriesData, XDDFDataSourcesFactory.fromArray(
                    tableData.get(i), chart.formatRange(new CellRangeAddress(1, numOfPoints, i, i))));
            seriesList.add(series);
        }
        for (int i = 0; i < seriesList.size(); i++) {
            seriesList.get(i).setTitle(colorTitles.get(i), setTitleInDataSheet(chart, colorTitles.get(i), 1));
        }
        // 指定为簇状柱状图
        if (tableData.size() > 1) {
            ((XDDFBarChartData) data).setBarGrouping(barChartForm.getGrouping());
            chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal(barChartForm.getNewOverlap());
        }

        // 指定系列颜色
        for (BarChartForm.ColorCheck colorCheck : barChartForm.getList()) {
            XDDFSolidFillProperties fillMarker = new XDDFSolidFillProperties(colorCheck.getXddfColor());
            XDDFShapeProperties propertiesMarker = new XDDFShapeProperties();
            // 给对象填充颜色属性
            propertiesMarker.setFillProperties(fillMarker);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(colorCheck.getNum()).addNewSpPr().set(propertiesMarker.getXmlObject());
        }

        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
        // 设置多个柱子之间的间隔
        // 绘制图形数据
        chart.plot(data);
        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);


    }


    /**
     * 创建折线图
     *
     * @param chart         图表对象
     * @param lineChartForm 数据对象
     */
    public void createLineChart(XWPFChart chart, LineChartForm lineChartForm) {
        // 标题
        chart.setTitleText(lineChartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        //分类轴标(X轴),标题位置
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(lineChartForm.getBottomTitle());
        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(lineChartForm.getLeftTitle());
        // 处理数据
        XDDFCategoryDataSource bottomDataSource = XDDFDataSourcesFactory.fromArray(lineChartForm.getBottomData());
        XDDFNumericalDataSource<Integer> leftDataSource = XDDFDataSourcesFactory.fromArray(lineChartForm.getLeftData());

        // 生成数据
        XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);

        // 不自动生成颜色
        data.setVaryColors(lineChartForm.getVaryColors());

        //图表加载数据，折线1
        XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(bottomDataSource, leftDataSource);

        //是否弯曲
        series.setSmooth(lineChartForm.getSmooth());

        //设置标记样式
        series.setMarkerStyle(lineChartForm.getStyle());

        //绘制
        chart.plot(data);
    }

    /**
     * 创建散点图
     *
     * @param chart            图表对象
     * @param scatterChartForm 数据对象
     */
    public void createScatterChart(XWPFChart chart, ScatterChartForm scatterChartForm) {
        // 标题
        chart.setTitleText(scatterChartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        //分类轴标(X轴),标题位置
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setTitle(scatterChartForm.getBottomTitle());
        //值(Y轴)轴,标题位置
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setTitle(scatterChartForm.getLeftTitle());
        XDDFScatterChartData data = null;
        for (int i = 0; i < scatterChartForm.getLists().size(); i++) {
            // 处理数据
            XDDFNumericalDataSource bottomDataSource = XDDFDataSourcesFactory.fromArray(scatterChartForm.getLists().get(i).getBottomData());
            XDDFNumericalDataSource<Integer> leftDataSource = XDDFDataSourcesFactory.fromArray(scatterChartForm.getLists().get(i).getLeftData());
            // 生成数据
            if (data == null) {
                data = (XDDFScatterChartData) chart.createData(ChartTypes.SCATTER, bottomAxis, leftAxis);
                // 是否自动生成颜色
                data.setVaryColors(false);
            }

            //图表加载数据，折线1
            XDDFScatterChartData.Series series = (XDDFScatterChartData.Series) data.addSeries(bottomDataSource, leftDataSource);
            //设置标记样式
            series.setMarkerStyle(scatterChartForm.getStyle());
            series.setMarkerSize(scatterChartForm.getMarkerSize());
            // 设置系列标题
            series.setTitle(scatterChartForm.getLists().get(i).getTitle(), null);
            // 去除连接线
            chart.getCTChart().getPlotArea().getScatterChartArray(0).getSerArray(i).addNewSpPr().addNewLn().addNewNoFill();
            if (scatterChartForm.getLists().get(i).getXddfColor() != null) {
                // 创建一个设置对象
                XDDFSolidFillProperties fillMarker = new XDDFSolidFillProperties(scatterChartForm.getLists().get(i).getXddfColor());
                XDDFShapeProperties propertiesMarker = new XDDFShapeProperties();
                // 给对象填充颜色属性
                propertiesMarker.setFillProperties(fillMarker);
                // 修改系列颜色
                chart.getCTChart().getPlotArea().getScatterChartArray(0).getSerArray(i).getMarker()
                        .addNewSpPr().set(propertiesMarker.getXmlObject());
            }
        }

        //绘制
        chart.plot(data);


    }

    /**
     * 创建饼状图
     *
     * @param chart        图表对象
     * @param pieChartForm 数据对象
     */
    public void createPieChart(XWPFChart chart, PieChartForm pieChartForm) {
        // 标题
        chart.setTitleText(pieChartForm.getTitle());
        //标题覆盖
        chart.setTitleOverlay(false);
        //图例位置
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        // 处理数据
        XDDFCategoryDataSource bottomDataSource = XDDFDataSourcesFactory.fromArray(pieChartForm.getBottomData());
        XDDFNumericalDataSource<Integer> leftDataSource = XDDFDataSourcesFactory.fromArray(pieChartForm.getLeftData());

        // 生成数据
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        // 自动生成颜色
        data.setVaryColors(false);

        //图表加载数据
        XDDFChartData.Series series = data.addSeries(bottomDataSource, leftDataSource);

        //绘制
        chart.plot(data);
    }

    /**
     * 添加word中的标记数据 标记方式为 ${text}
     *
     * @param document word对象
     * @param textMap  需要替换的信息集合
     */
    public void changeParagraphText(XWPFDocument document, Map<String, String> textMap) {
        //获取段落集合
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            //判断此段落时候需要进行替换
            String text = paragraph.getText();
            if (checkText(text)) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (XWPFRun run : runs) {
                    //替换模板原来位置
                    run.setText(changeValue(run.toString(), textMap), 0);
                }
            }
        }
    }

    /**
     * 替换表格中标记的数据 标记方式为 ${text}
     * 这里有个奇怪的问题 输入${}符号的时候需要把输入法切换到中文
     * ${}中间不能用数字,不能有下划线
     *
     * @param document word对象
     * @param textMap  需要替换的信息集合
     */
    public void changeTableText(XWPFDocument document, List<Map<String, String>> tableTextList) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            Map<String, String> textMap = tableTextList.get(i);
            //只处理行数大于等于2的表格
            XWPFTable table = tables.get(i);
            if (table.getRows().size() > 1) {
                //判断表格是需要替换还是需要插入，判断逻辑有$为替换，表格无$为插入
                if (checkText(table.getText())) {
                    List<XWPFTableRow> rows = table.getRows();
                    //遍历表格,并替换模板
                    eachTable(rows, textMap);
                }
            }
        }
    }

    /**
     * 复制表头,插入行数据,这里的样式和表头一样
     *
     * @param document word对象
     * @param list     集合个数和word中的表格个数必须相同
     */
    public void copyHeaderInsertText(XWPFDocument document, List<TableForm> list) {
        //获取表格对象集合
        List<XWPFTable> tables = document.getTables();
        // 循环word中的所有表格
        for (int k = 0; k < tables.size(); k++) {
            // 获取单个表格
            XWPFTable table = tables.get(k);
            // 获取要替换的数据
            TableForm tableForm = list.get(k);
            Integer headerIndex = tableForm.getStartLine();
            List<String[]> tableList = tableForm.getData();
            if (null == tableList) {
                return;
            }
            XWPFTableRow copyRow = table.getRow(headerIndex);
            List<XWPFTableCell> cellList = copyRow.getTableCells();
            if (null == cellList) {
                break;
            }
            //遍历要添加的数据的list
            for (int i = 0; i < tableList.size(); i++) {
                //插入一行
                XWPFTableRow targetRow = table.insertNewTableRow(headerIndex + 1 + i);
                //复制行属性
                targetRow.getCtRow().setTrPr(copyRow.getCtRow().getTrPr());

                String[] strings = tableList.get(i);
                for (int j = 0; j < strings.length; j++) {
                    XWPFTableCell sourceCell = cellList.get(j);
                    //插入一个单元格
                    XWPFTableCell targetCell = targetRow.addNewTableCell();
                    //复制列属性
                    targetCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
                    targetCell.setText(strings[j]);
                }
            }
        }
    }

    public static void main(String[] args) {
//        try (XWPFDocument document = new XWPFDocument(new FileInputStream("D:\\FreeMarker.docx"))) {
//            WordUtils wordUtils = new WordUtils();
//            Map<String, String> paragraphMap = new HashMap<>();
//            paragraphMap.put("number", "10000");
//            paragraphMap.put("date", "2020-03-25");
//            wordUtils.changeParagraphText(document, paragraphMap);
//
//            List<Map<String, String>> tableTextList = new ArrayList<>();
//            Map<String, String> tableMap = new HashMap<>();
//            tableMap.put("name", "赵云");
//            tableMap.put("sexual", "男");
//            tableMap.put("birthday", "2020-01-01");
//            tableMap.put("identify", "123456789");
//            tableMap.put("phone", "18377776666");
//            tableMap.put("address", "王者荣耀");
//            tableMap.put("domicile", "中国-腾讯");
//            tableMap.put("QQ", "是");
//            tableMap.put("chat", "是");
//            tableMap.put("blog", "是");
//            tableTextList.add(tableMap);
//            Map<String, String> tableMap2 = new HashMap<>();
//            tableMap2.put("spring", "sony的名称");
//            tableTextList.add(tableMap2);
//            wordUtils.changeTableText(document, tableTextList);
//
//            List<TableForm> list = new ArrayList<>();
//            TableForm tableForm = new TableForm();
//            tableForm.setStartLine(7);
//            tableForm.getData().add(new String[]{"露娜", "女", "野友", "666", "6660"});
//            tableForm.getData().add(new String[]{"鲁班", "男", "射友", "222", "2220"});
//            tableForm.getData().add(new String[]{"程咬金", "男", "肉友", "999", "9990"});
//            tableForm.getData().add(new String[]{"太乙真人", "男", "辅友", "111", "1110"});
//            tableForm.getData().add(new String[]{"貂蝉", "女", "法友", "888", "8880"});
//            list.add(tableForm);
//            TableForm tableForm2 = new TableForm();
//            tableForm2.setStartLine(1);
//            tableForm2.getData().add(new String[]{"18581588710", "蜘蛛侠", "100"});
//            tableForm2.getData().add(new String[]{"18581588710", "战神", "200"});
//            list.add(tableForm2);
//            wordUtils.copyHeaderInsertText(document,list);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }


//        try (XWPFDocument document = new XWPFDocument()) {
//            WordUtils wordUtils = new WordUtils();
//            XWPFChart chart = wordUtils.getChart(document, null, null);
//            PieChartForm pieChartForm = new PieChartForm();
//            pieChartForm.setTitle("标题");
//            pieChartForm.setBottomData(new String[]{"俄罗斯", "加拿大", "美国", "中国", "巴西", "澳大利亚", "印度"});
//            pieChartForm.setLeftData(new Integer[]{17098242, 9984670, 9826675, 9596961, 8514877, 7741220, 3287263});
//            wordUtils.createPieChart(chart, pieChartForm);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }

//
//        try (XWPFDocument document = new XWPFDocument()) {
//            WordUtils wordUtils = new WordUtils();
//            XWPFChart chart = wordUtils.getChart(document, null, null);
//            ScatterChartForm scatterChartForm = new ScatterChartForm();
//            scatterChartForm.setTitle("测试");
//            scatterChartForm.setBottomTitle("X轴");
//            scatterChartForm.setLeftTitle("Y轴");
//            scatterChartForm.setStyle(MarkerStyle.CIRCLE);
//            scatterChartForm.setMarkerSize((short) 10);
//            scatterChartForm.setVaryColors(false);
//
//            ScatterChartForm.AreaData areaData = new ScatterChartForm.AreaData();
//            areaData.setBottomData(new Integer[]{1, 2, 3, 4, 5, 8, 7});
//            areaData.setLeftData(new Integer[]{5, 5, 5, 4, 5, 6, 7});
//            areaData.setTitle("测试1");
//            scatterChartForm.getLists().add(areaData);
//
//            ScatterChartForm.AreaData areaData2 = new ScatterChartForm.AreaData();
//            areaData2.setBottomData(new Integer[]{6,9});
//            areaData2.setLeftData(new Integer[]{1,9});
//            areaData2.setXddfColor(XDDFColor.from(new byte[]{(byte)0xFF, (byte)0xE1, (byte)0xFF}));
//            areaData2.setTitle("测试2");
//            scatterChartForm.getLists().add(areaData2);
//            wordUtils.createScatterChart(chart, scatterChartForm);
//
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }

//        try (XWPFDocument document = new XWPFDocument()) {
//            WordUtils wordUtils = new WordUtils();
//            XWPFChart chart = wordUtils.getChart(document, null, null);
//            LineChartForm lineChartForm = new LineChartForm();
//            lineChartForm.setTitle("测试");
//            lineChartForm.setBottomTitle("X轴");
//            lineChartForm.setLeftTitle("Y轴");
//            lineChartForm.setStyle(MarkerStyle.STAR);
//            lineChartForm.setMarkerSize((short) 6);
//            lineChartForm.setSmooth(false);
//            lineChartForm.setVaryColors(false);
//            lineChartForm.setBottomData(new String[] {"俄罗斯","加拿大","美国","中国","巴西","澳大利亚","印度"});
//            lineChartForm.setLeftData(new Integer[] {17098242,9984670,9826675,9596961,8514877,7741220,3287263});
//            wordUtils.createLineChart(chart, lineChartForm);
//            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
//                document.write(fileOut);
//            }
//        } catch (Exception e) {
//
//        }


        try (XWPFDocument document = new XWPFDocument()) {
            WordUtils wordUtils = new WordUtils();
            XWPFChart chart = wordUtils.getChart(document, null, null);
            String[] categories = new String[]{"Lang 1", "Lang 2", "Lang 3"};
            Double[] valuesA = new Double[]{10d, 20d, 30d};
            Double[] valuesB = new Double[]{15d, 25d, 35d};
            Double[] valuesC = new Double[]{10d, 8d, 20d};
            List<Double[]> list = new ArrayList<>();
            list.add(valuesA);
            list.add(valuesB);
            list.add(valuesC);
            BarChartForm barChartForm = new BarChartForm();
            barChartForm.setTitle("测试");
            barChartForm.setCategories(categories);
            barChartForm.setTableData(list);
            barChartForm.setColorTitles(Arrays.asList("a", "b", "c"));
            barChartForm.setGrouping(BarGrouping.STACKED);
            barChartForm.setNewOverlap((byte) 100);

            BarChartForm.ColorCheck colorCheck = new BarChartForm.ColorCheck();
            colorCheck.setXddfColor(XDDFColor.from(new byte[]{(byte) 0xFF, (byte) 0x33, (byte) 0x00}));
            colorCheck.setNum(0);
            barChartForm.getList().add(colorCheck);

            BarChartForm.ColorCheck colorCheck2 = new BarChartForm.ColorCheck();
            colorCheck2.setXddfColor(XDDFColor.from(new byte[]{(byte) 0x91, (byte) 0x2C, (byte) 0xEE}));
            colorCheck2.setNum(1);
            barChartForm.getList().add(colorCheck2);

            BarChartForm.ColorCheck colorCheck3 = new BarChartForm.ColorCheck();
            colorCheck3.setXddfColor(XDDFColor.from(new byte[]{(byte) 0x00, (byte) 0x00, (byte) 0x80}));
            colorCheck3.setNum(2);
            barChartForm.getList().add(colorCheck3);


            wordUtils.createBarChart(chart, barChartForm);
            try (FileOutputStream fileOut = new FileOutputStream("CreateWordXDDFChart.docx")) {
                document.write(fileOut);
            }
        } catch (Exception e) {

        }
    }

    /**
     * 判断文本中时候包含$
     *
     * @param text 文本
     * @return 包含返回true, 不包含返回false
     */
    public static boolean checkText(String text) {
        boolean check = false;
        if (text.indexOf("$") != -1) {
            check = true;
        }
        return check;
    }

    /**
     * 匹配传入信息集合与模板
     *
     * @param value   模板需要替换的区域
     * @param textMap 传入信息集合
     * @return 模板需要替换区域信息集合对应值
     */
    public static String changeValue(String value, Map<String, String> textMap) {
        Set<Map.Entry<String, String>> textSets = textMap.entrySet();
        for (Map.Entry<String, String> textSet : textSets) {
            //匹配模板与替换值 格式${key}
            String key = "${" + textSet.getKey() + "}";
            if (value.indexOf(key) != -1) {
                value = textSet.getValue();
            }
        }
        //模板未匹配到区域替换为空
        if (checkText(value)) {
            value = "";
        }
        return value;
    }

    /**
     * 遍历表格,并替换模板
     *
     * @param rows    表格行对象
     * @param textMap 需要替换的信息集合
     */
    public static void eachTable(List<XWPFTableRow> rows, Map<String, String> textMap) {
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                //判断单元格是否需要替换
                if (checkText(cell.getText())) {
                    List<XWPFParagraph> paragraphs = cell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        List<XWPFRun> runs = paragraph.getRuns();
                        for (XWPFRun run : runs) {
                            run.setText(changeValue(run.toString(), textMap), 0);
                        }
                    }
                }
            }
        }
    }

    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null)
            row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null)
            cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }
}