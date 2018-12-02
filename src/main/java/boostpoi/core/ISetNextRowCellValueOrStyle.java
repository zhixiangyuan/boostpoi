package boostpoi.core;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * 设置指针指向的本列下一行的 cell 的 value or style，并且指针指向到本列下一行的 cell
 *
 * @author zhixiang.yuan
 * @data 2018/12/01 18:55
 */
public interface ISetNextRowCellValueOrStyle<T> {
    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellValue(boolean value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellValue(Calendar value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellValue(Date value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellValue(double value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellValue(RichTextString value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellValue(String value);

    /**
     * 设置指针指向的本列下一行的 cell 的样式，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setNextRowCellStyle(XSSFCellStyle value);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setNextRowCellValue(boolean value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setNextRowCellValue(Calendar value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setNextRowCellValue(Date value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setNextRowCellValue(double value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setNextRowCellValue(RichTextString value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setNextRowCellValue(String value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的样式
     *
     * @return this
     */
    T setNextRowCellStyle(XSSFCellStyle value, int n);
}
