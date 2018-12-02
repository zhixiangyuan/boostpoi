package boostpoi.core;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * 设置当前指针指向的 cell 的 value or style
 *
 * @author zhixiang.yuan
 * @data 2018/12/01 18:46
 */
public interface ISetCurrentCellValueOrStyle<T> {
    /**
     * 设置当前 cell 的值，指针不移动
     *
     * @return this
     */
    T setCurrentCellValue(boolean value);

    /**
     * 设置当前 cell 的值，指针不移动
     *
     * @return this
     */
    T setCurrentCellValue(Calendar value);

    /**
     * 设置当前 cell 的值，指针不移动
     *
     * @return this
     */
    T setCurrentCellValue(Date value);

    /**
     * 设置当前 cell 的值，指针不移动
     *
     * @return this
     */
    T setCurrentCellValue(double value);

    /**
     * 设置当前 cell 的值，指针不移动
     *
     * @return this
     */
    T setCurrentCellValue(RichTextString value);

    /**
     * 设置当前 cell 的值，指针不移动
     *
     * @return this
     */
    T setCurrentCellValue(String value);

    /**
     * 设置当前 cell 的样式，指针不移动
     *
     * @return this
     */
    T setCurrentCellStyle(XSSFCellStyle value);
}
