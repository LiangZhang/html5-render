package org.phyi.render.excel;

import org.phyi.render.util.Utils;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * ExcelCell
 *
 * @author czhouyi@gmail.com
 */
public class Text {
    private String value;
    private String color;
    private String fontFamily;
    private short fontSize;
    private boolean bold;
    private boolean strikeout;

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public short getFontSize() {
        return fontSize;
    }

    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public boolean isStrikeout() {
        return strikeout;
    }

    public void setStrikeout(boolean strikeout) {
        this.strikeout = strikeout;
    }

    private CharSequence style() {
        Map<String, String> styleMap = new LinkedHashMap<>();
        if (Utils.isNotBlank(this.color)) {
            styleMap.put("color", this.color);
        }
        if (this.fontSize > 0) {
            styleMap.put("font-size", String.format("%dpx", this.fontSize + 3));
        }
        if (Utils.isNotBlank(this.fontFamily)) {
            styleMap.put("font-family", this.fontFamily);
        }
        if (this.bold) {
            styleMap.put("font-weight", "bold");
        }

        StringBuilder styleBuffer = new StringBuilder();
        styleMap.forEach((k, v) -> styleBuffer.append(k).append(": ").append(v).append("; "));
        return styleBuffer;
    }

    @Override
    public String toString() {
        StringBuilder spanBuffer = new StringBuilder();
        spanBuffer.append("<span ");
        spanBuffer.append("style=\"");
        spanBuffer.append(style());
        spanBuffer.append("\">");
        if (this.strikeout) {
            spanBuffer.append("<s>");
        }
        spanBuffer.append(getValue());
        if (this.strikeout) {
            spanBuffer.append("</s>");
        }
        spanBuffer.append("</span>");
        return spanBuffer.toString();
    }
}
