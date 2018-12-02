package boostpoi.core;

/**
 * 控制指针接口
 *
 * @author zhixiang.yuan
 * @data 2018/12/01 19:19
 */
public interface IControlIndex<T> extends IMoveIndex<T> {

    /**
     * 获取当前行指针
     *
     * @return 行指针的值
     */
    int getRowIndex();

    /**
     * 设置当前行指针的值
     *
     * @param value 设置的值
     * @return this
     */
    T setRowIndex(int value);

    /**
     * 获取当前列指针
     *
     * @return 列指针的值
     */
    int getColumnIndex();

    /**
     * 设置当前列指针的值
     *
     * @param value 设置的值
     * @return this
     */
    T setColumnIndex(int value);

    /**
     * 重置 columnIndex
     *
     * @return this
     */
    T resetColumnIndex();

    /**
     * 重置 rowIndex
     *
     * @return this
     */
    T resetRowIndex();

    /**
     * 设置重置 columnIndex 时， columnIndex 的值
     *
     * @param value
     * @return this
     */
    T setResetColumnIndex(int value);

    /**
     * 设置重置 rowIndex 时，rowIndex 的值
     *
     * @param value
     * @return this
     */
    T setResetRowIndex(int value);

}
