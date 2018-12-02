package boostpoi.core;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;

/**
 * 为 poi 增加新特性
 *
 * @author zhixiang.yuan
 * @data 2018/11/28 16:06
 */
public class NewFeatureLayer<T extends NewFeatureLayer> extends ControlCellLayer<T> implements IBoost<T> {

    protected NewFeatureLayer() {
    }

    /**
     * 默认全局 {@link CellStyle}
     */
    private XSSFCellStyle defaultCellStyle;

    /**
     * 文件输出路径，不带文件扩展名
     */
    private String outputPath;

    private static final String FILE_SUFFIX = ".xlsx";

    @Override
    public T setDefaultCellStyle(XSSFCellStyle defaultCellStyle) {
        // # 遍历所有单元格，将未设置样式的单元格设置为默认样式
        this.defaultCellStyle = defaultCellStyle;
        return super.getSelf();
    }

    @Override
    public T createFile() {
        if (outputPath == null) {
            outputPath = "./Sample" + FILE_SUFFIX;
            log.info("未设置输出路径，故在当前目录下创建 {}", outputPath);
        }
        if (super.getWorkbook() instanceof XSSFWorkbook) {
            if (!outputPath.contains(FILE_SUFFIX)) {
                outputPath += FILE_SUFFIX;
            }
        }
        // 设置默认单元格格式
        if (this.defaultCellStyle != null) {
            for (Row row : super.getSheet()) {
                for (Cell cell : row) {
                    short cellStyleIndex = cell.getCellStyle().getIndex();
                    if (cellStyleIndex == Constant.DefaultCellStyle.DEFAULT_SYSTEM_CELL_STYLE.getCode()) {
                        cell.setCellStyle(this.defaultCellStyle);
                    }
                }
            }
        }
        try {
            FileOutputStream fileOut = new FileOutputStream(outputPath);
            // 把相应的 Excel 工作簿存盘
            super.getWorkbook().write(fileOut);
            fileOut.flush();
            fileOut.close();
            log.info("文件生成...");
            log.info("文件创建在 {} 位置。", this.outputPath);
        } catch (Exception e) {
            log.info("创建文件时出现异常");
            e.printStackTrace();
        }
        return super.getSelf();
    }

    @Override
    public T mergedRegion(int firstRow, int lastRow, int firstCol, int lastCol) {
        super.getSheet().addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        return super.getSelf();
    }

    @Override
    public XSSFCellStyle getEmptyCellStyle() {
        return super.getWorkbook().createCellStyle();
    }

    @Override
    public XSSFFont getEmptyFont() {
        return super.getWorkbook().createFont();
    }

    public T setOutputPath(String outputPath) {
        this.outputPath = outputPath;
        return super.getSelf();
    }

    public XSSFCellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }
}
