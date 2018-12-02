package boostpoi.core;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * 设置指针指向的本行下一列的 cell 的 value or style，并且指针指向到本行下一列的 cell
 *
 * @author zhixiang.yuan
 * @data 2018/12/01 18:59
 */
public interface ISetNextColumnCellValueOrStyle<T> {

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellValue(boolean value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellValue(Calendar value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellValue(Date value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellValue(double value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellValue(RichTextString value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellValue(String value);

    /**
     * 设置指针指向的本行下一列的 cell 的样式，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setNextColumnCellStyle(XSSFCellStyle value);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的值
     *
     * @return this
     */
    T setNextColumnCellValue(boolean value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的值
     *
     * @return this
     */
    T setNextColumnCellValue(Calendar value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的值
     *
     * @return this
     */
    T setNextColumnCellValue(Date value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的值
     *
     * @return this
     */
    T setNextColumnCellValue(double value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的值
     *
     * @return this
     */
    T setNextColumnCellValue(RichTextString value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的值
     *
     * @return this
     */
    T setNextColumnCellValue(String value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向后第 n 列的单元格的样式
     *
     * @return this
     */
    T setNextColumnCellStyle(XSSFCellStyle value, int n);
}
