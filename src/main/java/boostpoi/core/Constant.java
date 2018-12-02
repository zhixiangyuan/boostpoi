package boostpoi.core;

/**
 * @author zhixiang.yuan
 * @data 2018/11/28 16:20
 */
public class Constant {

    public enum DefaultCellStyle {

        /**
         * 默认创建 Cell 时选择的 CellStyle 类型
         */
        DEFAULT_SYSTEM_CELL_STYLE(0);

        int code;

        DefaultCellStyle(int code){
            this.code = code;
        }

        public int getCode() {
            return code;
        }
    }

}
