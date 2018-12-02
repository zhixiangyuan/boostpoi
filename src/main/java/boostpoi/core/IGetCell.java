package boostpoi.core;

/**
 * 获取单元格
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 15:47
 */
public interface IGetCell {

    /**
     * 获取本行下一列的单元格
     *
     * @return this
     */
    BoostCell getNextColumnCell();

    /**
     * 获取本行前一列的单元格
     *
     * @return this
     */
    BoostCell getPreviousColumnCell();

    /**
     * 获取本列下一行的单元格
     *
     * @return this
     */
    BoostCell getNextRowCell();

    /**
     * 获取本列上一行的单元格
     *
     * @return this
     */
    BoostCell getPreviousRowCell();

    /**
     * 获取从本列开始向后数第 n 列的单元格
     *
     * @return this
     */
    BoostCell getNextColumnCell(int n);

    /**
     * 获取从本列开始向前数第 n 列的单元格
     *
     * @return this
     */
    BoostCell getPreviousColumnCell(int n);

    /**
     * 获取从本行开始向后数第 n 行的单元格
     *
     * @return this
     */
    BoostCell getNextRowCell(int n);

    /**
     * 获取从本行开始向前数第 n 行的单元格
     *
     * @return this
     */
    BoostCell getPreviousRowCell(int n);

    /**
     * 获取当前指针指向的单元格
     *
     * @return this
     */
    BoostCell getCurrentBoostCell();
}
