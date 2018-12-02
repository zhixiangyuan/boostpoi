package boostpoi.core;

/**
 * poi 增强类
 *
 * @author zhixiang.yuan
 * @create 2018/11/22 19:49
 */
public interface IBoostPOI {

    /**
     * @return  本行后一列的单元格
     */
    BoostCell getNextColumnCell();

    /**
     * @return  本行前一列的单元格
     */
    BoostCell getPreviousColumnCell();

    /**
     * @return  本列下一行的单元格
     */
    BoostCell getNextRowCell();

    /**
     * @return  本列上一行的单元格
     */
    BoostCell getPreviousRowCell();

    /**
     * @return  当前所指单元格
     */
    BoostCell getCurrentCell();


}
