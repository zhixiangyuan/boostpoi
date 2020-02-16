package boostpoi.example;

import boostpoi.core.BoostPOI;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * @author zhixiang.yuan
 * @data 2018/11/28 12:21
 */
public class Example {

    public static void main(String[] args) {
        Example example = new Example();
        example.example_1();
    }

    private final static String OUT_PUT_PATH = "/Users/zhixiangyuan/workspace/tmp/boostpoi/src/main/java/boostpoi/example/Sample-2.xlsx";

    /**
     * 第一个使用示例，先看文件，再看代码
     */
    private void example_1() {
        /** 1. 使用构造器创建 {@link BoostPOI} */
        BoostPOI boostpoi = new BoostPOI.Builder().setOutputPath(OUT_PUT_PATH).build();
        /** 2. 设置默认列宽 */
        boostpoi.getSheet().setDefaultColumnWidth(1);
        boostpoi.getSheet().setDefaultRowHeight((short)250);
        /** 3. 创建单元格样式 */
        XSSFCellStyle cellStyle = boostpoi.getEmptyCellStyle();
        cellStyle.setFillBackgroundColor((short)0);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        /** 3. 向单元格中放入值 */
        boostpoi.moveToNextRow(10)
                .moveToNextColumn(15)
                .setResetColumnIndex(boostpoi.getColumnIndex())
                .setResetRowIndex(boostpoi.getRowIndex())
                // 第一行
                .setNextColumnCellStyle(cellStyle,18)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                // 第二行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,17)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                // 第三行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,16)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle,3)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                // 第四行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,16)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle,3)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                // 第五行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,15)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle,3)
                .setNextColumnCellStyle(cellStyle,7)
                // 第六行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,16)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                // 第七行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,16)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                // 第八行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,15)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle,3)
                // 第九行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,12)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,9)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle)
                // 第十行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,11)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,13)
                .setNextColumnCellStyle(cellStyle,10)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                // 第十一行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,11)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle,17)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle)
                // 第十二行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,11)
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,12)
                .setNextColumnCellStyle(cellStyle,7)
                // 第十三行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,10)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,12)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                // 第十四行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,9)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle,11)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle)
                // 第十五行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,16)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                // 第十六行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,11)
                .setNextColumnCellStyle(cellStyle,10)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                // 第十七行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle,10)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,9)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                // 第十八行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,3)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle,9)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                // 第十九行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                // 第二十行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,3)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,2)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,5)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,7)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,6)
                .setNextColumnCellStyle(cellStyle)
                // 第二一行
                .resetColumnIndex().moveToNextRow()
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,8)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,4)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle,9)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle)
                .setNextColumnCellStyle(cellStyle);
        /** 4. 创建文件 */
        boostpoi.createFile();
    }
}
