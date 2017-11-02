package org.phyi.render;

import org.phyi.render.excel.ExcelRender;

/**
 * ExcelRenderTest
 *
 * @author czhouyi@gmail.com
 */
public class ExcelRenderTest {

    public static void main(String[] args) throws Exception {
        ExcelRender render = new ExcelRender("C:\\Users\\admin\\Desktop\\workwork\\销售项目\\销售系统-开发计划.xlsx");
        render.render();
    }
}
