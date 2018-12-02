package boostpoi.core;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * 对 {@link Cell} 进行包装
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 11:26
 */
public class BoostCell{

    private XSSFCell cell;

    public XSSFCell getCell() {
        return cell;
    }

    BoostCell setCell(XSSFCell cell) {
        this.cell = cell;
        return this;
    }

    public BoostCell setCellValue(boolean value) {
        cell.setCellValue(value);
        return this;
    }

    public BoostCell setCellValue(Calendar value) {
        cell.setCellValue(value);
        return this;
    }

    public BoostCell setCellValue(Date value) {
        cell.setCellValue(value);
        return this;
    }

    public BoostCell setCellValue(double value) {
        cell.setCellValue(value);
        return this;
    }

    public BoostCell setCellValue(RichTextString value) {
        cell.setCellValue(value);
        return this;
    }

    public BoostCell setCellValue(String value) {
        cell.setCellValue(value);
        return this;
    }

    public BoostCell setCellStyle(XSSFCellStyle cellStyle) {
        cell.setCellStyle(cellStyle);
        return this;
    }
}
