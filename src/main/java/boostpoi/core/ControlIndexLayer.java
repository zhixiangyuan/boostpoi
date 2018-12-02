package boostpoi.core;

/**
 * 指针控制层
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 15:34
 */
public class ControlIndexLayer<T extends ControlIndexLayer> extends PoiBase<T> implements IControlIndex<T> {

    protected ControlIndexLayer() {
        this.resetColumnIndex = this.columnIndex;
        this.resetRowIndex = this.rowIndex;
    }

    /**
     * 行指针
     */
    protected int rowIndex;
    /**
     * 列指针
     */
    protected int columnIndex;

    private int resetColumnIndex;

    private int resetRowIndex;

    @Override
    public T moveToNextColumn() {
        ++this.columnIndex;
        return super.getSelf();
    }

    @Override
    public T moveToNextRow() {
        this.rowIndex++;
        return super.getSelf();
    }

    @Override
    public T moveToPreviousColumn() {
        this.columnIndex--;
        return super.getSelf();
    }

    @Override
    public T moveToPreviousRow() {
        this.rowIndex--;
        return super.getSelf();
    }

    @Override
    public T moveToNextColumn(int n) {
        this.columnIndex = this.columnIndex + n;
        return super.getSelf();
    }

    @Override
    public T moveToNextRow(int n) {
        this.rowIndex = this.rowIndex + n;
        return super.getSelf();
    }

    @Override
    public T moveToPreviousColumn(int n) {
        this.columnIndex = this.columnIndex - n;
        return super.getSelf();
    }

    @Override
    public T moveToPreviousRow(int n) {
        this.rowIndex = this.rowIndex - n;
        return super.getSelf();
    }

    @Override
    public int getRowIndex() {
        return rowIndex;
    }

    @Override
    public T setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
        return super.getSelf();
    }

    @Override
    public int getColumnIndex() {
        return columnIndex;
    }

    @Override
    public T setColumnIndex(int columnIndex) {
        this.columnIndex = columnIndex;
        return super.getSelf();
    }

    @Override
    public T resetColumnIndex() {
        this.columnIndex = this.resetColumnIndex;
        return super.getSelf();
    }

    @Override
    public T resetRowIndex() {
        this.rowIndex = this.resetRowIndex;
        return super.getSelf();
    }

    @Override
    public T setResetColumnIndex(int value) {
        this.resetColumnIndex = value;
        return super.getSelf();
    }

    @Override
    public T setResetRowIndex(int value) {
        this.resetRowIndex = value;
        return super.getSelf();
    }
}
