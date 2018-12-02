package boostpoi.core;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * 增强的功能
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 16:04
 */
public interface IBoost<T> {

    /**
     * 设置默认单元格格式
     *
     * @param defaultCellStyle 默认单元格格式
     */
    T setDefaultCellStyle(XSSFCellStyle defaultCellStyle);

    /**
     * 生成文件
     */
    T createFile();

    /**
     * 合并单元格
     *
     * @param firstRow 首行
     * @param lastRow  最后一行
     * @param firstCol 首列
     * @param lastCol  最后一列
     * @return this
     */
    T mergedRegion(int firstRow, int lastRow, int firstCol, int lastCol);

    /**
     * 获取一个空的单元格样式
     *
     * @return 空的单元格样式
     */
    XSSFCellStyle getEmptyCellStyle();

    /**
     * 获取一个空的字体
     *
     * @return 空的字体
     */
    Font getEmptyFont();
}
