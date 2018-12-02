package boostpoi.core;

/**
 * 继承对单元格控制的接口，使得代码结构化，易于管理
 *
 * @author zhixiang.yuan
 * @data 2018/12/01 18:45
 */
public interface IControlCell<T> extends IGetCell, ISetCurrentCellValueOrStyle<T>, ISetNextColumnCellValueOrStyle<T>
        , ISetNextRowCellValueOrStyle<T>, ISetPreviousColumnValueOrStyle<T>, ISetPreviousRowCellValueOrStyle<T> {
}
