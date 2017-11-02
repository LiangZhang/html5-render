package org.phyi.render.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Utils
 *
 * @author czhouyi@gmail.com
 */
public class Utils {

    public static SimpleDateFormat SDF = new SimpleDateFormat("yyyy/MM/dd");

    public static boolean isNotBlank(String input) {
        return input != null && input.length() > 0;
    }

    public static boolean isBlank(String input) {
        return input == null || input.length() <= 0;
    }

    public static String sBlank(String input) {
        return input == null ? "" : input;
    }

    public static InputStream readFile(String path) throws IOException {
        return new FileInputStream(new File(path));
    }

    public static Object getCellValue(Cell cell) {
        CellType cellType = cell.getCellTypeEnum();
        XSSFCell xCell = (XSSFCell) cell;
        if (cellType == CellType.NUMERIC) {
            String value = xCell.getRawValue();
            if (value != null && value.length() == 5) {
                Date date = xCell.getDateCellValue();
                value = SDF.format(date);
            }
            return value;
        } else if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cellType == CellType.BLANK) {
            return "";
        }
        return "";
    }

    private static int red(int c) {
        return c & 0XFF;
    }

    private static int green(int c) {
        return (c >> 8) & 0XFF;
    }

    private static int blue(int c) {
        return (c >> 16) & 0XFF;
    }

    private static int rgb(int c) {
        return (red(c) << 16) | (green(c) << 8) | blue(c);
    }

    private static String rgbToSix(String rgb) {
        int length = 6 - rgb.length();
        StringBuilder str = new StringBuilder();
        while (length > 0) {
            str.append("0");
            length--;
        }
        return str + rgb;
    }

    public static String getHexColor(int color) {
        color = color == -1 ? 0 : color;
        int rgb = rgb(color);
        return "#" + rgbToSix(Integer.toHexString(rgb));
    }
}
