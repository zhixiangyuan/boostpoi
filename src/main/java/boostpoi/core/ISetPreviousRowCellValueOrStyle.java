package boostpoi.core;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * 设置指针指向的本列下一行的 cell 的 value or style，并且指针指向到本列下一行的 cell
 *
 * @author zhixiang.yuan
 * @data 2018/12/01 19:02
 */
public interface ISetPreviousRowCellValueOrStyle<T> {
    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellValue(boolean value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellValue(Calendar value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellValue(Date value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellValue(double value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellValue(RichTextString value);

    /**
     * 设置指针指向的本列下一行的 cell 的值，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellValue(String value);

    /**
     * 设置指针指向的本列下一行的 cell 的样式，并且指针指向到本列下一行的 cell
     *
     * @return this
     */
    T setPreviousRowCellStyle(XSSFCellStyle value);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellValue(boolean value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellValue(Calendar value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellValue(Date value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellValue(double value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellValue(RichTextString value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellValue(String value, int n);

    /**
     * 设置从本行(本行为 0)开始数，向下第 n 行的值
     *
     * @return this
     */
    T setPreviousRowCellStyle(XSSFCellStyle value, int n);
}
