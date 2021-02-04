package com.lh.word.form;


import lombok.Data;

import java.util.ArrayList;
import java.util.List;

/**
 * Copyright (C), 2006-2010, ChengDu ybya info. Co., Ltd.
 * FileName: TableForm.java
 *
 * @author lh
 * @version 1.0.0
 * @Date 2021/02/03 11:50
 */
@Data
public class TableForm {
    // 表头行号,从0开始
    private Integer startLine;

    // 要复制数据 长度为复制的行数
    private List<String[]> data = new ArrayList<>();

}