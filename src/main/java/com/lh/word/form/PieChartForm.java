package com.lh.word.form;


import lombok.Data;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: PieChartForm.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/03 11:29
 */
@Data
public class PieChartForm extends ChartFrom{
    // 数据个数
    private String[] bottomData;

    // 数据大小
    private Integer[] leftData;
}