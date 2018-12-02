package boostpoi.core;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * @author zhixiang.yuan
 * @data 2018/12/01 19:05
 */
public interface ISetPreviousColumnValueOrStyle<T> {

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellValue(boolean value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellValue(Calendar value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellValue(Date value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellValue(double value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellValue(RichTextString value);

    /**
     * 设置指针指向的本行下一列的 cell 的值，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellValue(String value);

    /**
     * 设置指针指向的本行下一列的 cell 的样式，并且指针指向到本行下一列的 cell
     *
     * @return this
     */
    T setPreviousColumnCellStyle(XSSFCellStyle value);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的值
     *
     * @return this
     */
    T setPreviousColumnCellValue(boolean value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的值
     *
     * @return this
     */
    T setPreviousColumnCellValue(Calendar value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的值
     *
     * @return this
     */
    T setPreviousColumnCellValue(Date value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的值
     *
     * @return this
     */
    T setPreviousColumnCellValue(double value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的值
     *
     * @return this
     */
    T setPreviousColumnCellValue(RichTextString value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的值
     *
     * @return this
     */
    T setPreviousColumnCellValue(String value, int n);

    /**
     * 设置从本列（本列为 0 ）开始数，向前第 n 列的单元格的样式
     *
     * @return this
     */
    T setPreviousColumnCellStyle(XSSFCellStyle value, int n);
}
