package com.lh.word.form;


import lombok.Data;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: LineChartForm.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/03 10:33
 */
@Data
public class LineChartForm extends ChartFrom {
    // X轴数据
    private String[] bottomData;

    // Y轴数据
    private Integer[] leftData;

    // 标记大小 默认为6
    private Short markerSize = 6;

    // 标记样式 默认圆
    private MarkerStyle style = MarkerStyle.CIRCLE;

    // 是否弯曲 默认不
    private Boolean smooth;

    // 是否自动生成颜色
    private Boolean varyColors;

}