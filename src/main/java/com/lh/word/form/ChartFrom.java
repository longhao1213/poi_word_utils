package com.lh.word.form;


import lombok.Data;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: ChartFrom.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/03 10:35
 */
@Data
public class ChartFrom {
    // 图表标题
    private String title;

    // X轴标题
    private String bottomTitle;

    // Y轴标题
    private String leftTitle;
}