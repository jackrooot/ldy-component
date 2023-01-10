# 霖点雨component

#### 介绍
为了使工作重心放到核心内容上面，一些重复性的工作，例如EXCEL基础化导入导出、CRUD生成、API文档，自动化测试等，进一步封装逐步完善


#### 安装教程
1.      <dependency>
             <groupId>cn.lindianyu</groupId>
             <artifactId>ldy-component</artifactId>
             <version>1.0.1</version>
         </dependency>
#### 使用说明
1. 模板类如: @Excel(name = "编号", orderNum = "1")
             private String no = "123124";
2. 
    导出
    
    Workbook install = new ExcelDownBuild<TestExcel>()
             .buildSheetEndNum(200000)
             .buildExcel(testExcels, TestExcel.class, null);
    //输出流
    install.write(outputStream);
    
    读取.xlsx
    ：InputStream inputStream= new FileInputStream(new File("F:\\4.xlsx"));
             List<TestExcel> dataList = new ExcelXlsxReadBuild<TestExcel>()
                     .buildFile(inputStream, TestExcel.class)
                     .buildSheetNum(0)
                     .build();
    读取.xls
    ：FileInputStream inputStream = new FileInputStream(new File("F:\\4.xls"));
            List<TestExcel> dataList = new ExcelXlsSaxReadBuild<TestExcel>()
                    .buildFile(inputStream, TestExcel.class)
                    .buildStartRow(1)
                    .buildJumpBlankRow(true)
                    .buildSheetNum(0)
                    .build();
    
#### 参与贡献

1.  POI
2.  lombok
3.  commons-lang3