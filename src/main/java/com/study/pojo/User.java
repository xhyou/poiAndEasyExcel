package com.study.pojo;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;

import java.util.Date;

@Data
public class User {

    @ExcelProperty(value = "姓名",index =0 )
    private String name;

    @ExcelProperty(value = "日期",index =1 )
    private Date date;

    @ExcelProperty(value = "精度数据",index =2 )
    private Double doubleData;

    @ExcelProperty(value = "年龄",index =3 )
    private int age;

    @ExcelProperty(value = "性别",index =4 )
    private boolean sex;
}
