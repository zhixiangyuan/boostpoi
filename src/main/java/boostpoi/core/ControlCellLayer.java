package boostpoi.core;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.Calendar;
import java.util.Date;

/**
 * 单元格控制层
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 15:50
 */
public class ControlCellLayer<T extends ControlCellLayer> extends ControlIndexLayer<T> implements IControlCell<T> {

    protected ControlCellLayer(){
    }

    @Override
    public BoostCell getNextColumnCell() {
        super.moveToNextColumn();
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getPreviousColumnCell() {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getNextRowCell() {
        super.moveToNextRow();
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getPreviousRowCell() {
        super.moveToPreviousRow();
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getNextColumnCell(int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getPreviousColumnCell(int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getNextRowCell(int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getPreviousRowCell(int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell();
        return super.getCurrentBoostCell();
    }

    @Override
    public BoostCell getCurrentBoostCell() {
        super.getCurrentBoostCell().setCell(getCell(super.rowIndex, super.columnIndex));
        return super.getCurrentBoostCell();
    }

    private XSSFCell getCell(int rowIndex, int columnIndex) {
        if (super.getSheet() == null) {
            super.setSheet(super.getWorkbook().createSheet());
        }
        if (super.getSheet().getRow(rowIndex) == null) {
            super.getSheet().createRow(rowIndex);
        }
        if (super.getSheet().getRow(rowIndex).getCell(columnIndex) == null) {
            super.getSheet().getRow(rowIndex).createCell(columnIndex);
        }
        return super.getSheet().getRow(rowIndex).getCell(columnIndex);
    }


    @Override
    public T setCurrentCellValue(boolean value) {
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setCurrentCellValue(Calendar value) {
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setCurrentCellValue(Date value) {
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setCurrentCellValue(double value) {
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setCurrentCellValue(RichTextString value) {
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setCurrentCellValue(String value) {
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setCurrentCellStyle(XSSFCellStyle value) {
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(boolean value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(Calendar value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(Date value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(double value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(RichTextString value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(String value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellStyle(XSSFCellStyle value) {
        super.moveToNextColumn();
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(boolean value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(Calendar value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(Date value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(double value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(RichTextString value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellValue(String value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextColumnCellStyle(XSSFCellStyle value, int n) {
        super.moveToNextColumn(n);
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(boolean value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(Calendar value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(Date value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(double value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(RichTextString value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(String value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellStyle(XSSFCellStyle value) {
        super.moveToNextRow();
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(boolean value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(Calendar value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(Date value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(double value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(RichTextString value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellValue(String value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setNextRowCellStyle(XSSFCellStyle value, int n) {
        super.moveToNextRow(n);
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(boolean value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(Calendar value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(Date value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(double value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(RichTextString value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(String value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellStyle(XSSFCellStyle value) {
        super.moveToPreviousColumn();
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(boolean value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(Calendar value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(Date value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(double value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(RichTextString value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellValue(String value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousColumnCellStyle(XSSFCellStyle value, int n) {
        super.moveToPreviousColumn(n);
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(boolean value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(Calendar value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(Date value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(double value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(RichTextString value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(String value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellStyle(XSSFCellStyle value) {
        super.moveToPreviousRow();
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(boolean value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(Calendar value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(Date value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(double value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(RichTextString value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellValue(String value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellValue(value);
        return super.getSelf();
    }

    @Override
    public T setPreviousRowCellStyle(XSSFCellStyle value, int n) {
        super.moveToPreviousRow(n);
        this.getCurrentBoostCell().setCellStyle(value);
        return super.getSelf();
    }
}
