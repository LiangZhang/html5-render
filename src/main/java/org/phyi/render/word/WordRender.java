package org.phyi.render.word;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.phyi.render.Html5Render;
import org.phyi.render.util.Utils;

import java.util.ArrayList;
import java.util.List;

/**
 * Render Word File to Html5 string
 *
 * @author czhouyi@gmail.com
 */
public class WordRender implements Html5Render {

    private HWPFDocument document;
    private XWPFDocument xdocument;

    /**
     * 回车符ASCII码
     */
    private static final short ENTER_ASCII = 13;
    /**
     * 空格符ASCII码
     */
    private static final short SPACE_ASCII = 32;
    /**
     * 水平制表符ASCII码
     */
    private static final short TABULATION_ASCII = 9;

    public WordRender(String file) throws Exception {
        this.document = new HWPFDocument(Utils.readFile(file));
    }

    @Override
    public Object render() {
        StringBuilder htmlBuilder = new StringBuilder();
        try {
            Range range = document.getRange();
            TableIterator it = new TableIterator(range);
            List<WordTable> tableList = parseTable(it);
            int tableSize = tableList.size();

            PicturesTable pTable = document.getPicturesTable();

            htmlBuilder.append("<html>\n");
            htmlBuilder.append("<head><title>");
            htmlBuilder.append(document.getSummaryInformation().getTitle());
            htmlBuilder.append("</title></head>\n");
            htmlBuilder.append("<body>");

            // 创建临时字符串,好加以判断一串字符是否存在相同格式
            int tableIndex = 0;
            StringBuilder styleGroup = new StringBuilder();
            for (int i = 0, length = document.characterLength(); i < length - 1; i++) {
                // 整篇文章的字符通过一个个字符的来判断,range为得到文档的范围
                Range r = new Range(i, i + 1, document);
                CharacterRun characterRun = r.getCharacterRun(0);
                if (tableSize > tableIndex) {
                    WordTable wordTable = tableList.get(tableIndex);
                    if (i == wordTable.getStartIndex()) {
                        htmlBuilder.append(wordTable.getHtml());
                        i = wordTable.getEndIndex() - 1;
                        tableIndex++;
                        continue;
                    }
                }
                if (pTable.hasPicture(characterRun)) {
                    readPicture(pTable, characterRun);
                    continue;
                }
                Range range2 = new Range(i + 1, i + 2, document);
                // 第二个字符
                CharacterRun cr2 = range2.getCharacterRun(0);
                char c = characterRun.text().charAt(0);

                if (c == SPACE_ASCII) {
                    // 判断是否为空格符
                    styleGroup.append("&nbsp;");
                } else if (c == TABULATION_ASCII) {
                    // 判断是否为水平制表符
                    styleGroup.append("&nbsp;&nbsp;");
                }
                // 比较前后2个字符是否具有相同的格式
                boolean flag = compareCharStyle(characterRun, cr2);
                if (flag && c != ENTER_ASCII) {
                    styleGroup.append(characterRun.text());
                } else {
                    String fontStyle = "<span style=\"font-family:" + characterRun.getFontName() + ";font-size:" + characterRun.getFontSize() / 2
                            + "pt;color:" + Utils.getHexColor(characterRun.getIco24()) + ";";
                    if (characterRun.isBold()) {
                        fontStyle += "font-weight:bold;";
                    }
                    if (characterRun.isItalic()) {
                        fontStyle += "font-style:italic;";
                    }

                    htmlBuilder.append(fontStyle).append("\">").append(styleGroup).append(characterRun.text());
                    htmlBuilder.append("</span>\n");
                    styleGroup = new StringBuilder();
                }
                // 判断是否为回车符
                if (c == ENTER_ASCII) {
                    htmlBuilder.append("<br/>\n");
                }
            }

            htmlBuilder.append("</body>\n</html>");
        } catch (Exception ignored) {

        }
        return htmlBuilder.toString();
    }

    /**
     * 读写文档中的表格
     *
     * @param it
     */
    private List<WordTable> parseTable(TableIterator it) {
        List<WordTable> tableList = new ArrayList<>();
        while (it.hasNext()) {
            Table tb = it.next();
            WordTable wordTable = new WordTable();
            wordTable.setStartIndex(tb.getStartOffset());
            wordTable.setEndIndex(tb.getEndOffset());

            StringBuilder tableBuilder = new StringBuilder();
            tableBuilder.append("<table border>\n");
            for (int i = 0; i < tb.numRows(); i++) {
                TableRow tr = tb.getRow(i);

                tableBuilder.append("<tr>\n");
                //迭代列，默认从0开始
                for (int j = 0; j < tr.numCells(); j++) {
                    TableCell td = tr.getCell(j);//取得单元格
                    int cellWidth = td.getWidth();
                    tableBuilder.append("<td width=").append(cellWidth).append(">");
                    //取得单元格的内容
                    for (int k = 0; k < td.numParagraphs(); k++) {
                        if (k > 0) {
                            tableBuilder.append("<br/>");
                        }
                        Paragraph para = td.getParagraph(k);
                        String s = para.text().trim();
                        if (Utils.isBlank(s)) {
                            s = " ";
                        }
                        tableBuilder.append(s);
                    }
                    tableBuilder.append("</td>\n");
                }
                tableBuilder.append("</tr>\n");
            }
            tableBuilder.append("</table>\n");
            wordTable.setHtml(tableBuilder.toString());
            tableList.add(wordTable);
        }

        return tableList;
    }

    /**
     * 读写文档中的图片
     *
     * @param pTable
     * @param cr
     * @throws Exception
     */
    public static void readPicture(PicturesTable pTable, CharacterRun cr) throws Exception {
        // 提取图片
//        Picture pic = pTable.extractPicture(cr, false);
//        // 返回POI建议的图片文件名
//        String afileName = pic.suggestFullFileName();
//
//        File file = new File(wordImageFilePath());
//        System.out.println(file.mkdirs());
//        OutputStream out = new FileOutputStream(new File( wordImageFilePath()+ File.separator + afileName));
//        pic.writeImageContent(out);
//        htmlText += "<img src='"+wordImgeWebPath()+ afileName
//                + "' mce_src='"+wordImgeWebPath()+ afileName + "' />";
    }

    private boolean compareCharStyle(CharacterRun cr1, CharacterRun cr2) {
        return cr1.isBold() == cr2.isBold() && cr1.isItalic() == cr2.isItalic()
                && cr1.getFontName().equals(cr2.getFontName())
                && cr1.getFontSize() == cr2.getFontSize() && cr1.getColor() == cr2.getColor();
    }
}
