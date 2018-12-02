package boostpoi.core;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author zhixiang.yuan
 * @data 2018/12/01 19:13
 */
public class PoiBase<T> {

    protected final Logger log = LoggerFactory.getLogger(this.getClass());

    private BoostCell currentBoostCell;

    private XSSFWorkbook workbook;

    private XSSFSheet sheet;

    private T self;

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public T setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
        return this.self;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public T setSheet(XSSFSheet sheet) {
        this.sheet = sheet;
        return this.self;
    }

    protected T getSelf() {
        return self;
    }

    protected void setSelf(T self) {
        this.self = self;
    }

    protected BoostCell getCurrentBoostCell() {
        return currentBoostCell;
    }

    public T setCurrentBoostCell(BoostCell currentCell) {
        this.currentBoostCell = currentCell;
        return this.self;
    }
}
