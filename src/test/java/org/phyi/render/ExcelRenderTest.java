package org.phyi.render;

import org.junit.Test;
import org.phyi.render.excel.ExcelRender;

/**
 * ExcelRenderTest
 *
 * @author czhouyi@gmail.com
 */
public class ExcelRenderTest {

    @Test
    public void testRender() throws Exception {
        ExcelRender render = new ExcelRender("C:\\Users\\admin\\Desktop\\excel-sample.xlsx");
        render.render();
    }
}
