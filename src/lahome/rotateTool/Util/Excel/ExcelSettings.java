package lahome.rotateTool.Util.Excel;

import org.apache.poi.ss.util.CellReference;

import java.util.Locale;

/**
 * Created by Administrator on 2017/5/6.
 */
public class ExcelSettings {

    public static String rotateFilePath;
    public static int rotateSheetIndex = 1;
    public static int rotateFirstDataRow;
    public static int rotateKitColumn;
    public static int rotatePartColumn;
    public static int rotatePmQtyColumn;
    public static int rotateApQtyColumn;
    public static int rotateRatioColumn;
    public static int rotateApplySetColumn;
    public static int rotateRemarkColumn;

    public static String stockFilePath;
    public static int stockSheetIndex = 1;
    public static int stockFirstDataRow;
    public static int stockKitColumn;
    public static int stockPartColumn;
    public static int stockPoColumn;
    public static int stockStQtyColumn;
    public static int stockLotColumn;
    public static int stockDcColumn;
    public static int stockApQtyColumn;
    public static int stockRemarkColumn;

    public static String purchaseFilePath;
    public static int purchaseSheetIndex = 1;
    public static int purchaseFirstDataRow;
    public static int purchaseKitColumn;
    public static int purchasePartColumn;
    public static int purchasePoColumn;
    public static int purchaseGrDateColumn;
    public static int purchaseGrQtyColumn;
    public static int purchaseApQtyColumn;
    public static int purchaseApSetColumn;
    public static int purchaseRemarkColumn;

    public static void setRotateConfig(String filePath, int firstDataRow, String kitColStr, String partColStr,
                                String pmQtyColStr, String apQtyColStr, String ratioColStr, String applySetColStr,
                                String remarkColStr) {

        rotateFilePath = filePath;
        rotateFirstDataRow = firstDataRow - 1;

        rotateKitColumn = CellReference.convertColStringToIndex(kitColStr.toUpperCase());
        rotatePartColumn = CellReference.convertColStringToIndex(partColStr.toLowerCase());
        rotatePmQtyColumn = CellReference.convertColStringToIndex(pmQtyColStr.toUpperCase());
        rotateApQtyColumn = CellReference.convertColStringToIndex(apQtyColStr.toUpperCase());
        rotateRatioColumn = CellReference.convertColStringToIndex(ratioColStr.toUpperCase());
        rotateApplySetColumn = CellReference.convertColStringToIndex(applySetColStr.toUpperCase());
        rotateRemarkColumn = CellReference.convertColStringToIndex(remarkColStr.toUpperCase());
    }

    public static void setStockConfig(String filePath, int firstDataRow, String kitColStr, String partColStr,
                               String poColStr, String stQtyColStr, String lotColStr, String dcColStr,
                               String apQtyColStr, String remarkColStr) {

        stockFilePath = filePath;
        stockFirstDataRow = firstDataRow - 1;

        stockKitColumn = CellReference.convertColStringToIndex(kitColStr);
        stockPartColumn = CellReference.convertColStringToIndex(partColStr);
        stockPoColumn = CellReference.convertColStringToIndex(poColStr);
        stockStQtyColumn = CellReference.convertColStringToIndex(stQtyColStr);
        stockLotColumn = CellReference.convertColStringToIndex(lotColStr);
        stockDcColumn = CellReference.convertColStringToIndex(dcColStr);
        stockApQtyColumn = CellReference.convertColStringToIndex(apQtyColStr);
        stockRemarkColumn = CellReference.convertColStringToIndex(remarkColStr);
    }

    public static void setPurchaseConfig(String filePath, int firstDataRow, String kitColStr, String partColStr,
                                  String poColStr, String grDateColStr, String grQtyColStr,
                                  String apQtyColStr, String setColStr, String remarkColStr) {

        purchaseFilePath = filePath;
        purchaseFirstDataRow = firstDataRow - 1;

        purchaseKitColumn = CellReference.convertColStringToIndex(kitColStr);
        purchasePartColumn = CellReference.convertColStringToIndex(partColStr);
        purchasePoColumn = CellReference.convertColStringToIndex(poColStr);
        purchaseGrDateColumn = CellReference.convertColStringToIndex(grDateColStr);
        purchaseGrQtyColumn = CellReference.convertColStringToIndex(grQtyColStr);
        purchaseApQtyColumn = CellReference.convertColStringToIndex(apQtyColStr);
        purchaseApSetColumn = CellReference.convertColStringToIndex(setColStr);
        purchaseRemarkColumn = CellReference.convertColStringToIndex(remarkColStr);
    }


}
