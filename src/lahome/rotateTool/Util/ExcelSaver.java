package lahome.rotateTool.Util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lahome.rotateTool.module.PurchaseItem;
import lahome.rotateTool.module.RotateCollection;
import lahome.rotateTool.module.RotateItem;
import lahome.rotateTool.module.StockItem;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.util.CellReference;

import java.util.Comparator;
import java.util.List;

/**
 * Created by Administrator on 2017/5/6.
 */
public class ExcelSaver {
    private static final Logger log = LogManager.getLogger(PurchaseItem.class.getName());

    private RotateCollection collection;
    private Dispatch currentSheet;

    public ExcelSaver(RotateCollection collection) {
        this.collection = collection;
    }

    public void writeCell(int row, int col, String data) {
        try {
            String pos = ExcelCommon.convertNumToColString(col) + row;
            Dispatch cell = Dispatch.invoke(currentSheet, "Range", Dispatch.Get, new Object[]{pos}, new int[1]).toDispatch();
            Dispatch.put(cell, "Value", data);
        } catch (Exception e) {
            log.error("failed!", e);
        }
    }

    public void saveRotateToExcel() {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant(ExcelCommon.rotateFilePath)).toDispatch();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{ExcelCommon.rotateSheetIndex}, new int[1]).toDispatch();
            String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            currentSheet = sheet;

            for (RotateItem item : collection.getRotateObsList()) {
                if (item.isDuplicate())
                    continue;

                int rowNum = item.getRowNum();
                writeCell(rowNum, ExcelCommon.rotateApQtyColumn, String.valueOf(item.getStockApplyQtyTotal()));
                writeCell(rowNum, ExcelCommon.rotateApplySetColumn, String.valueOf(item.getApplySet()));
                writeCell(rowNum, ExcelCommon.rotateRemarkColumn, item.getRemark());
            }

            // Saves and closes
//            Dispatch.call(workBook, "Save");
//            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
//            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //excel.invoke("Quit", new Variant[0]);
            ComThread.Release();
        }
    }

    public void saveStockToExcel() {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant(ExcelCommon.stockFilePath)).toDispatch();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{ExcelCommon.stockSheetIndex}, new int[1]).toDispatch();
            String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            currentSheet = sheet;

            for (StockItem item : collection.getStockItemList()) {
                if (item.getRotateItem().isDuplicate())
                    continue;

                int rowNum = item.getRowNum();
                writeCell(rowNum, ExcelCommon.stockApQtyColumn, String.valueOf(item.getApplyQty()));
                writeCell(rowNum, ExcelCommon.stockRemarkColumn, item.getRemark());
            }

            // Saves and closes
//            Dispatch.call(workBook, "Save");
//            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
//            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //excel.invoke("Quit", new Variant[0]);
            ComThread.Release();
        }
    }

    public void savePurchaseToExcel() {
        ComThread.InitSTA(true);
        final ActiveXComponent excel = new ActiveXComponent("Excel.Application");
        try {
            //excel.setProperty("EnableEvents", new Variant(false));
            Dispatch.put(excel, "Visible", new Variant(true));

            Dispatch workbooks = excel.getProperty("Workbooks").toDispatch();
            Dispatch workbook = Dispatch.call(workbooks, "Open", new Variant(ExcelCommon.purchaseFilePath)).toDispatch();
            Dispatch sheets = Dispatch.get(workbook, "Worksheets").toDispatch();//獲得所有的Sheet
            int SheetCount = Dispatch.get(sheets, "Count").getInt();//獲得有多少個sheet
            Dispatch sheet = Dispatch.invoke(sheets, "Item", Dispatch.Get, new Object[]{ExcelCommon.purchaseSheetIndex}, new int[1]).toDispatch();
            String sheetName = Dispatch.get(sheet, "Name").toString();//獲得sheet的名字
            Dispatch userRange = Dispatch.call(sheet, "UsedRange").toDispatch();//獲取Excel使用的sheet
            Dispatch row = Dispatch.call(userRange, "Rows").toDispatch();
            int rowCount = Dispatch.get(row, "Count").getInt();//excel的使用的行數

            currentSheet = sheet;

            for (PurchaseItem item : collection.getPurchaseItemList()) {
                if (item.getRotateItem().isDuplicate())
                    continue;

                int rowNum = item.getRowNum();
                writeCell(rowNum, ExcelCommon.purchaseApQtyColumn, String.valueOf(item.getApplyQty()));
                writeCell(rowNum, ExcelCommon.purchaseApSetColumn, String.valueOf(item.getApplySet()));
                writeCell(rowNum, ExcelCommon.purchaseRemarkColumn, item.getRemark());
            }

            // Saves and closes
//            Dispatch.call(workBook, "Save");
//            com.jacob.com.Variant f = new com.jacob.com.Variant(true);
//            Dispatch.call(workBook, "Close", f);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            //excel.invoke("Quit", new Variant[0]);
            ComThread.Release();
        }

    }
}
