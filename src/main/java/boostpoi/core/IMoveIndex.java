package boostpoi.core;

/**
 * 移动单元格
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 15:23
 */
public interface IMoveIndex<T> {

    /**
     * 列指针加一
     *
     * @return this
     */
    T moveToNextColumn();

    /**
     * 行指针加一
     *
     * @return this
     */
    T moveToNextRow();

    /**
     * 列指针减一
     *
     * @return this
     */
    T moveToPreviousColumn();

    /**
     * 行指针减一
     *
     * @return this
     */
    T moveToPreviousRow();

    /**
     * 列指针加 n
     *
     * @return this
     */
    T moveToNextColumn(int n);

    /**
     * 行指针加 n
     *
     * @return this
     */
    T moveToNextRow(int n);

    /**
     * 列指针减 n
     *
     * @return this
     */
    T moveToPreviousColumn(int n);

    /**
     * 行指针减 n
     *
     * @return this
     */
    T moveToPreviousRow(int n);




}
