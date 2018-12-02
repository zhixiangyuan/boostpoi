package boostpoi.core;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 为 poi 装上刀把
 *
 * TODO 添加全局列宽、行高
 *
 * @author zhixiang.yuan
 * @date 2018/11/22 19:45
 */
public class BoostPOI extends NewFeatureLayer<BoostPOI> {

    private BoostPOI(Builder builder) {
        super();
        super.setSelf(this);
        super.setRowIndex(builder.getRowIndex());
        super.setColumnIndex(builder.getColumnIndex());
        super.setWorkbook(builder.getWorkbook());
        super.setSheet(builder.getSheet());
        super.setCurrentBoostCell(builder.getBoostCell());
        super.setOutputPath(builder.getOutputPath());
        if (super.getSheet() == null) {
            super.setSheet(super.getWorkbook().createSheet());
        }
    }

    public static class Builder {

        private XSSFWorkbook workbook;

        private XSSFSheet sheet;

        private String outputPath;

        private Integer rowIndex;

        private Integer columnIndex;

        private BoostCell boostCell;

        BoostCell getBoostCell() {
            return boostCell;
        }

        public Builder setWorkBook(XSSFWorkbook workbook) {
            this.workbook = workbook;
            return this;
        }

        XSSFWorkbook getWorkbook() {
            return workbook;
        }

        public Builder setOutputPath(String outputPath) {
            this.outputPath = outputPath;
            return this;
        }

        String getOutputPath() {
            return outputPath;
        }

        public Builder setSheet(XSSFSheet sheet) {
            this.sheet = sheet;
            return this;
        }

        XSSFSheet getSheet() {
            return sheet;
        }

        Integer getRowIndex() {
            return rowIndex;
        }

        public Builder setRowIndex(Integer rowIndex) {
            this.rowIndex = rowIndex;
            return this;
        }

        Integer getColumnIndex() {
            return columnIndex;
        }

        public Builder setColumnIndex(Integer columnIndex) {
            this.columnIndex = columnIndex;
            return this;
        }

        public BoostPOI build() {
            // 默认创建 XSSFWorkbook 格式的 Workbook
            if (this.getWorkbook() == null) {
                this.setWorkBook(new XSSFWorkbook());
            }
            // rowIndex 默认指向 0
            if (this.getRowIndex() == null) {
                this.setRowIndex(0);
            }
            // columnIndex 默认指向 0
            if (this.getColumnIndex() == null) {
                this.setColumnIndex(0);
            }
            boostCell = new BoostCell();
            return new BoostPOI(this);
        }

    }
}
